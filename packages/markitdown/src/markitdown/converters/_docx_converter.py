import sys
import io
import os
import base64
import mimetypes
import tempfile
import hashlib
from pathlib import Path
from warnings import warn

from typing import BinaryIO, Any

# PIL for image conversion
try:
    from PIL import Image
    _pil_available = True
except ImportError:
    _pil_available = False

from ._html_converter import HtmlConverter
from ..converter_utils.docx.pre_process import pre_process_docx
from .._base_converter import DocumentConverterResult
from .._stream_info import StreamInfo
from .._exceptions import MissingDependencyException, MISSING_DEPENDENCY_MESSAGE

# Try loading optional (but in this case, required) dependencies
# Save reporting of any exceptions for later
_dependency_exc_info = None
try:
    import mammoth

except ImportError:
    # Preserve the error and stack trace for later
    _dependency_exc_info = sys.exc_info()

# DEPRECATED: docxlatex import no longer needed
# OMML equation handling is now done directly in mammoth
# See mammoth/docx/body_xml.py:omath() for implementation


ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
]

ACCEPTED_FILE_EXTENSIONS = [".docx"]


# DEPRECATED: These functions are no longer needed
# mammoth now handles OMML formulas directly via integrated docxlatex.OMMLParser
# See mammoth/docx/body_xml.py:omath() for implementation
#
# def _extract_text_with_equations(file_path_or_stream):
#     """Extract text with embedded equations from DOCX using docxlatex."""
#     ...
#
# def _extract_formulas_with_context(docxlatex_text, context_words=10):
#     """Extract formulas from docxlatex output with surrounding context."""
#     ...
#
# def _match_formulas_by_context(mammoth_text, formulas_with_context):
#     """Match formulas to image placeholders in mammoth text using context anchors."""
#     ...


class DocxImageWriter:
    """
    Custom image writer for saving DOCX images to a directory.
    Compatible with mammoth's convert_image callback.
    """

    def __init__(self, output_dir):
        """
        Initialize the image writer.

        Args:
            output_dir: Directory to save images to
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.image_counter = 0

    def __call__(self, image):
        """
        Process a single image from the DOCX document.
        Automatically converts WMF/EMF to PNG for Markdown compatibility.

        Args:
            image: mammoth image object with .open(), .content_type, .alt_text

        Returns:
            dict: HTML attributes for the <img> tag, containing 'src' path
        """
        # Determine file extension from content type
        content_type = image.content_type or "image/png"
        extension = mimetypes.guess_extension(content_type)
        if not extension:
            # Default to .png if unable to determine
            extension = '.png'

        # Check if this is WMF/EMF - these need to be converted to PNG
        is_wmf_emf = (extension.lower() in ['.wmf', '.emf'] or
                      'wmf' in content_type.lower() or
                      'emf' in content_type.lower())

        # Generate sequential filename
        self.image_counter += 1

        if is_wmf_emf:
            # For WMF/EMF, save as WMF first, then convert to PNG
            temp_filename = f"image_{self.image_counter:03d}{extension}"
            temp_filepath = self.output_dir / temp_filename

            # Save original WMF/EMF
            with image.open() as image_bytes:
                with open(temp_filepath, 'wb') as f:
                    f.write(image_bytes.read())

            # Convert to PNG
            png_filename = f"image_{self.image_counter:03d}.png"
            png_filepath = self.output_dir / png_filename

            try:
                import subprocess
                # Use ImageMagick to convert WMF/EMF to PNG
                result = subprocess.run(
                    [
                        'magick',
                        '-density', '600',      # High DPI
                        str(temp_filepath),     # Input WMF/EMF
                        '-background', 'white',
                        '-alpha', 'remove',
                        '-colorspace', 'RGB',
                        '-quality', '100',
                        str(png_filepath)       # Output PNG
                    ],
                    capture_output=True,
                    text=True,
                    timeout=30
                )

                if result.returncode == 0 and png_filepath.exists():
                    # Conversion successful - delete WMF/EMF and use PNG
                    temp_filepath.unlink()
                    filename = png_filename
                    filepath = png_filepath
                else:
                    # Conversion failed - keep WMF/EMF
                    print(f"[WARN] WMF/EMF conversion failed, keeping original: {result.stderr}")
                    filename = temp_filename
                    filepath = temp_filepath
            except Exception as e:
                # Conversion failed - keep WMF/EMF
                print(f"[WARN] WMF/EMF conversion error: {e}")
                filename = temp_filename
                filepath = temp_filepath
        else:
            # Regular image (PNG, JPEG, etc.) - save directly
            filename = f"image_{self.image_counter:03d}{extension}"
            filepath = self.output_dir / filename

            # Save image data to file
            with image.open() as image_bytes:
                with open(filepath, 'wb') as f:
                    f.write(image_bytes.read())

        # Return relative path for the markdown
        return {
            "src": f"{self.output_dir.name}/{filename}"
        }


class DocxImageWriterWith_llm(DocxImageWriter):
    """
    Enhanced image writer with OCR support for mathematical formulas.
    Automatically recognizes formula images and converts them to LaTeX.
    """

    def __init__(self, output_dir,
                 ocr_formulas=True,
                 formula_mode='auto',
                 cache_ocr_results=True,
                 llm_client=None,
                 llm_model=None):
        """
        Initialize enhanced image writer with LLM-based OCR capabilities.

        Args:
            output_dir: Directory to save images to
            ocr_formulas: Enable formula OCR recognition (default: True)
            formula_mode: Formula type - 'auto', 'inline', 'display' (default: 'auto')
            cache_ocr_results: Cache OCR results to avoid re-processing (default: True)
            llm_client: LLM client for formula recognition
            llm_model: LLM model name
        """
        super().__init__(output_dir)

        # OCR configuration
        self.ocr_formulas = ocr_formulas
        self.formula_mode = formula_mode

        # LLM configuration
        self.llm_client = llm_client
        self.llm_model = llm_model

        # OCR result cache
        self.cache_ocr_results = cache_ocr_results
        self.ocr_cache = {}  # {image_hash: latex_string}

        # Temporary files for cleanup
        self.temp_files = []

        # Statistics
        self.stats = {
            'total_images': 0,
            'formula_detected': 0,
            'ocr_success': 0,
            'ocr_failed': 0,
            'ocr_cached': 0
        }

    def __del__(self):
        """Cleanup: remove temporary files"""
        # Cleanup temporary PNG files
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except:
                pass


    def _is_likely_formula(self, filepath, content_type=None):
        """
        Heuristic to detect if an image is likely a mathematical formula.

        Args:
            filepath: Path to image file
            content_type: MIME type of the image

        Returns:
            bool: True if image is likely a formula
        """
        filepath = Path(filepath)

        # Rule 1: WMF/EMF formats strongly indicate formulas
        if filepath.suffix.lower() in ['.wmf']:
            return True

        # Rule 2: Check MIME type
        if content_type:
            if 'wmf' in content_type.lower() or 'emf' in content_type.lower():
                return True

        # Rule 3: File size check (formulas typically < 100KB)
        try:
            file_size = filepath.stat().st_size
            if file_size > 100 * 1024:  # 100KB
                return False
        except:
            pass

        # Default: conservative approach for non-WMF files
        return False

    def _convert_to_png_if_needed(self, filepath):
        """
        Convert WMF/EMF to PNG for OCR processing using ImageMagick.

        Args:
            filepath: Path to image file

        Returns:
            str: Path to PNG file (may be temporary)
        """
        filepath = Path(filepath)

        if filepath.suffix.lower() not in ['.wmf', '.emf']:
            return str(filepath)  # No conversion needed

        try:
            import subprocess

            # Create temporary PNG file
            temp_png = tempfile.mktemp(suffix='.png', dir=self.output_dir)

            # Use ImageMagick to convert with high quality settings
            # Correct parameter order: input options -> input file -> processing options -> output file
            # -density 600: Use 600 DPI for high resolution (default is 72)
            # -background white -alpha remove: Ensure white background (after reading input)
            # -quality 100: Maximum PNG quality
            # -colorspace RGB: Ensure proper color space
            result = subprocess.run(
                [
                    'magick',
                    '-density', '600',      # Input option: High DPI
                    str(filepath),          # Input file
                    '-background', 'white', # Processing: white background
                    '-alpha', 'remove',     # Processing: remove transparency
                    '-colorspace', 'RGB',   # Processing: RGB colorspace
                    '-quality', '100',      # Output option: maximum quality
                    temp_png                # Output file
                ],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode == 0 and os.path.exists(temp_png):
                self.temp_files.append(temp_png)
                return temp_png
            else:
                print(f"[WARN] ImageMagick conversion failed for {filepath.name}: {result.stderr}")
                return str(filepath)

        except FileNotFoundError:
            print("[WARN] ImageMagick (magick) not found in PATH. Cannot convert WMF/EMF.")
            print("  Install ImageMagick: https://imagemagick.org/")
            return str(filepath)
        except subprocess.TimeoutExpired:
            print(f"[WARN] Conversion timeout for {filepath.name}")
            return str(filepath)
        except Exception as e:
            print(f"[WARN] Failed to convert {filepath.name} to PNG: {e}")
            return str(filepath)  # Return original path

    def _get_image_hash(self, filepath):
        """Calculate MD5 hash of image for caching"""
        try:
            with open(filepath, 'rb') as f:
                return hashlib.md5(f.read()).hexdigest()
        except:
            return None


    def _recognize_formula(self, image_path):
        """
        Recognize formula from image using LLM.

        Args:
            image_path: Path to image file

        Returns:
            str: LaTeX formula string, or None if failed
        """
        # Check if LLM is available
        if not self.llm_client or not self.llm_model:
            print("[WARN] No LLM client configured. Formula OCR disabled.")
            self.stats['ocr_failed'] += 1
            return None

        # Check cache
        img_hash = None
        if self.cache_ocr_results:
            img_hash = self._get_image_hash(image_path)
            if img_hash and img_hash in self.ocr_cache:
                self.stats['ocr_cached'] += 1
                return self.ocr_cache[img_hash]

        # Convert to PNG if needed
        try:
            png_path = self._convert_to_png_if_needed(image_path)
        except Exception as e:
            print(f"  [FAIL] Image conversion failed: {e}")
            self.stats['ocr_failed'] += 1
            return None

        # Use LLM to recognize formula
        try:
            # Encode image to base64
            with open(png_path, 'rb') as f:
                image_data = base64.b64encode(f.read()).decode('utf-8')

            # Prepare messages for LLM
            messages = [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{image_data}"
                            }
                        },
                        {
                            "type": "text",
                            "text": "Convert the formula image to LaTeX. Rule: Use $...$ if the resulting formula is under 10 characters; otherwise, use $$...$$.  Only output the LaTeX code without any explanation."
                        }
                    ]
                }
            ]

            # Call LLM
            response = self.llm_client.chat.completions.create(
                model=self.llm_model,
                messages=messages,
                max_tokens=500
            )

            latex = response.choices[0].message.content.strip()

            if latex:
                # Remove markdown code blocks if present
                #latex = latex.replace('```latex', '').replace('```', '').strip()

                # Cache result
                if self.cache_ocr_results and img_hash:
                    self.ocr_cache[img_hash] = latex

                self.stats['ocr_success'] += 1
                print(f"  [OK] LLM OCR: {Path(image_path).name} -> {latex[:100]}{'...' if len(latex) > 100 else ''}")
                return latex

        except Exception as e:
            print(f"  [WARN] LLM OCR error: {e}")
            self.stats['ocr_failed'] += 1
            return None

        self.stats['ocr_failed'] += 1
        return None

    def __call__(self, image):
        """
        Process image: save and optionally perform OCR.

        Args:
            image: mammoth image object

        Returns:
            dict: {"src": "path"} or {"src": "LATEX_FORMULA:..."}
        """
        self.stats['total_images'] += 1

        # First, save the image using parent class logic
        result = super().__call__(image)
        filepath = Path(result["src"])

        # If OCR is disabled or image is not a formula, return as-is
        content_type = image.content_type or "image/png"

        if not self.ocr_formulas:
            return result

        if not self._is_likely_formula(filepath, content_type):
            return result

        # This is likely a formula - attempt OCR
        self.stats['formula_detected'] += 1

        latex = self._recognize_formula(filepath)

        if latex:
            # LLM already returns formatted LaTeX with $ or $$
            # Return special marker for post-processing
            return {
                "src": f"LATEX_FORMULA:{latex}"
            }

        # OCR failed - return image path
        return result

    def print_stats(self):
        """Print OCR processing statistics"""
        print("\n" + "="*60)
        print("OCR Processing Statistics:")
        print(f"  Total images: {self.stats['total_images']}")
        print(f"  Formulas detected: {self.stats['formula_detected']}")
        print(f"  OCR successful: {self.stats['ocr_success']}")
        print(f"  OCR failed: {self.stats['ocr_failed']}")
        print(f"  OCR cached: {self.stats['ocr_cached']}")
        print("="*60)


class DocxConverter(HtmlConverter):
    """
    Converts DOCX files to Markdown. Style information (e.g.m headings) and tables are preserved where possible.
    """

    def __init__(self, llm_client=None, llm_model=None):
        super().__init__()
        self._html_converter = HtmlConverter()
        self.llm_client = llm_client
        self.llm_model = llm_model

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension in ACCEPTED_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> DocumentConverterResult:
        # Check: the dependencies
        if _dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".docx",
                    feature="docx",
                )
            ) from _dependency_exc_info[
                1
            ].with_traceback(  # type: ignore[union-attr]
                _dependency_exc_info[2]
            )

        style_map = kwargs.get("style_map", None)
        images_dir = kwargs.get("save_images_dir")
        keep_data_uris = kwargs.get("keep_data_uris", False)

        # OCR configuration
        ocr_formulas = kwargs.get("ocr_formulas", True)  # Default: enabled
        formula_mode = kwargs.get("formula_mode", "auto")  # auto, inline, or display

        # Native equation extraction configuration
        extract_native_equations = kwargs.get("extract_native_equations", True)  # Default: enabled

        # Setup image conversion strategy
        convert_image = None
        image_writer = None

        if images_dir:
            # Choose image writer based on OCR setting
            if ocr_formulas:
                # Use enhanced writer with LLM-based OCR support
                image_writer = DocxImageWriterWith_llm(
                    images_dir,
                    ocr_formulas=True,
                    formula_mode=formula_mode,
                    llm_client=self.llm_client,
                    llm_model=self.llm_model
                )
            else:
                # Use basic image writer
                image_writer = DocxImageWriter(images_dir)

            convert_image = mammoth.images.img_element(image_writer)

        elif not keep_data_uris:
            # Default: use base64 data URIs (mammoth's default behavior)
            convert_image = mammoth.images.data_uri
        # If keep_data_uris is True, convert_image stays None, using mammoth's default

        pre_process_stream = pre_process_docx(file_stream)

        # DEPRECATED: OMML equation extraction is now handled directly by mammoth
        # See mammoth/docx/body_xml.py:omath() for implementation
        # text_with_native_equations = ""  # No longer needed

        # Convert DOCX to HTML with image handling
        if convert_image is not None:
            html_result = mammoth.convert_to_html(
                pre_process_stream,
                style_map=style_map,
                convert_image=convert_image
            )
        else:
            html_result = mammoth.convert_to_html(
                pre_process_stream,
                style_map=style_map
            )

        # Convert HTML to Markdown
        markdown_result = self._html_converter.convert_string(
            html_result.value,
            **kwargs,
        )

        text_content = markdown_result.text_content

        # Post-processing 0: Decode OMML placeholders from mammoth
        # Format: ⟨OMML:$:base64⟩ or ⟨OMML:$$:base64⟩
        import re
        import base64

        def decode_omml_placeholder(match):
            delimiter = match.group(1)  # $ or $$
            latex_b64 = match.group(2)
            try:
                # Remove markdown escaping that markdownify may have added
                # markdownify escapes + as \+ in base64 strings
                latex_b64_clean = latex_b64.replace(r'\+', '+').replace(r'\/', '/')
                latex = base64.b64decode(latex_b64_clean).decode('utf-8')
                # For display equations ($$), add blank line after
                if delimiter == "$$":
                    result = f"{delimiter}{latex}{delimiter}"
                    return result + chr(10) + chr(10)
                else:
                    return f"{delimiter}{latex}{delimiter}"
            except Exception as e:
                # Return original if decode fails
                print(f"[WARNING] Failed to decode OMML placeholder: {e}")
                return match.group(0)

        # Replace all ⟨OMML:...:...⟩ markers
        # Pattern needs to handle escaped characters from markdownify
        text_content = re.sub(
            r'⟨OMML:(\$+):([A-Za-z0-9+/=\\]+)⟩',
            decode_omml_placeholder,
            text_content
        )

        # Post-processing 1: DEPRECATED - docxlatex fallback no longer needed
        # mammoth now handles OMML formulas directly via integrated docxlatex.OMMLParser
        # See mammoth/docx/body_xml.py:omath() for implementation
        # if text_with_native_equations:
        #     # ... old docxlatex post-processing code removed ...

        # Post-processing 2: Replace LaTeX formula markers from OCR
        if ocr_formulas and image_writer and isinstance(image_writer, DocxImageWriterWith_llm):
            import re

            # Replace LATEX_FORMULA markers with actual LaTeX
            # Pattern: ![](LATEX_FORMULA:$formula$) or ![](LATEX_FORMULA:$$formula$$)
            # NOTE: Cannot use (.*?) because it stops at first ) inside LaTeX formulas
            # Use custom function to handle nested parentheses
            def replace_latex_formula(match_obj):
                """Replace LATEX_FORMULA markers, handling nested parentheses in LaTeX"""
                text = match_obj.string
                start = match_obj.end()  # Position after "![](LATEX_FORMULA:"

                # Find the matching closing ) by counting parentheses
                depth = 1
                pos = start
                while pos < len(text) and depth > 0:
                    if text[pos] == '(':
                        depth += 1
                    elif text[pos] == ')':
                        depth -= 1
                    pos += 1

                # Extract the LaTeX content
                latex = text[start:pos-1]  # -1 to exclude the final )
                return latex

            # Use custom replacement function
            text_content = re.sub(
                r'!\[\]\(LATEX_FORMULA:',
                lambda m: replace_latex_formula(m),
                text_content
            )

            # Print OCR statistics
            image_writer.print_stats()

        # Create final result
        if text_content != markdown_result.text_content:
            markdown_result = DocumentConverterResult(
                text_content,  # First positional parameter is 'markdown'
                title=markdown_result.title
            )

        return markdown_result
