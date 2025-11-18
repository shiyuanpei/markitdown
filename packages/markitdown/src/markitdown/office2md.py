#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Office2Markdown CLI Tool

Convert Office documents (DOCX, PPTX) to Markdown with enhanced equation support.
Supports image extraction, OMML equations, and LLM-based OCR.
"""

import sys
import os
import argparse
from pathlib import Path
import glob

# Ensure UTF-8 encoding on Windows
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

try:
    from markitdown import MarkItDown
    from openai import OpenAI
except ImportError as e:
    print(f"ERROR: Required package not installed: {e}")
    print("Please install required packages")
    sys.exit(1)


def convert_file(input_path, output_path=None, images_dir=None,
                 enable_llm=True, enable_omml=True, verbose=True):
    """
    Convert a single Office document to Markdown with image extraction.

    Args:
        input_path: Path to input file
        output_path: Path to output file (auto-generated if None)
        images_dir: Directory to save extracted images (auto-generated if None)
        enable_llm: Enable LLM for formula images
        enable_omml: Enable native OMML equation extraction
        verbose: Print progress messages

    Returns:
        Path to output file or None if failed
    """
    input_path = Path(input_path)

    if not input_path.exists():
        print(f"ERROR: File not found: {input_path}")
        return None

    # Auto-generate output filename if not provided
    if output_path is None:
        output_path = input_path.with_suffix('.md')
    else:
        output_path = Path(output_path)

    # Auto-generate images directory if not provided
    if images_dir is None:
        images_dir = output_path.parent / f"{output_path.stem}_images"
    else:
        images_dir = Path(images_dir)

    try:
        if verbose:
            print(f"Converting: {input_path.name}")

        # Check for API key if LLM is enabled
        api_key = os.getenv('OPENROUTER_API_KEY')

        if enable_llm and api_key:
            # Create OpenAI client for OpenRouter
            client = OpenAI(
                base_url='https://openrouter.ai/api/v1',
                api_key=api_key
            )
            md = MarkItDown(
                llm_client=client,
                llm_model='google/gemini-2.5-flash'
            )
        else:
            # No LLM client
            md = MarkItDown()
            if enable_llm and not api_key and verbose:
                print("  Note: OPENROUTER_API_KEY not set, LLM disabled")

        # Convert with all advanced options
        result = md.convert(
            str(input_path),
            save_images_dir=str(images_dir),
            extract_native_equations=enable_omml,
            llm_formulas=enable_llm and bool(api_key),
            formula_mode='auto'
        )

        # Save to file
        output_path.write_text(result.text_content, encoding='utf-8')

        if verbose:
            print(f"  → Saved to: {output_path.name}")

            # Count images
            if images_dir.exists():
                image_count = len(list(images_dir.glob('*.*')))
                if image_count > 0:
                    print(f"  → Extracted {image_count} images to: {images_dir.name}/")

            # Count formulas
            import re
            inline_formulas = result.text_content.count('$') // 2
            display_formulas = result.text_content.count('$$') // 2
            total_formulas = inline_formulas + display_formulas

            if total_formulas > 0:
                print(f"  → Found {total_formulas} formulas")

        return output_path

    except Exception as e:
        print(f"ERROR converting {input_path.name}: {e}")
        if verbose:
            import traceback
            traceback.print_exc()
        return None


def main():
    """Main CLI entry point"""
    parser = argparse.ArgumentParser(
        description='Convert Office documents to Markdown with enhanced equation support',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  # Convert single DOCX (with image extraction)
  office2md document.docx
  office2md document.docx -o output.md

  # Specify custom images directory
  office2md document.docx -d images/

  # Convert without LLM (faster, free)
  office2md document.docx --no-llm

  # Batch convert
  office2md *.docx
  office2md *.pptx

Supported Formats:
  - DOCX (Microsoft Word documents)
  - PPTX (PowerPoint presentations)
  - PDF, XLSX, HTML, Images, etc.

Features:
  - Native OMML equation support (converted to LaTeX)
  - Automatic image extraction to separate directory
  - LLM-based OCR for MathType/WMF images (requires API key)
  - Image format conversion (WMF/EMF to PNG)
  - Batch processing support

LLM Configuration:
  Set environment variable to enable formula OCR:
    Windows: set OPENROUTER_API_KEY=your_key_here
    Linux/Mac: export OPENROUTER_API_KEY=your_key_here

  Get API key from: https://openrouter.ai/
        '''
    )

    parser.add_argument(
        'input',
        nargs='+',
        help='Input file(s) - supports wildcards'
    )

    parser.add_argument(
        '-o', '--output',
        help='Output file (only valid for single input)'
    )

    parser.add_argument(
        '-d', '--images-dir',
        help='Images directory (default: <output_name>_images/)'
    )

    parser.add_argument(
        '--no-llm',
        action='store_true',
        help='Disable LLM for formula images'
    )

    parser.add_argument(
        '--no-omml',
        action='store_true',
        help='Disable native OMML equation extraction'
    )

    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Suppress progress messages'
    )

    parser.add_argument(
        '--version',
        action='version',
        version='office2markdown 1.0.2'
    )

    args = parser.parse_args()

    # Expand wildcards
    input_files = []
    for pattern in args.input:
        matches = glob.glob(pattern)
        if matches:
            input_files.extend(matches)
        else:
            # Not a pattern, add as-is
            input_files.append(pattern)

    if not input_files:
        print("ERROR: No input files specified")
        return 1

    # Validate arguments
    if args.output and len(input_files) > 1:
        print("ERROR: --output can only be used with a single input file")
        return 1

    if args.images_dir and len(input_files) > 1:
        print("ERROR: --images-dir can only be used with a single input file")
        return 1

    verbose = not args.quiet
    enable_llm = not args.no_llm
    enable_omml = not args.no_omml

    # Process files
    success_count = 0
    fail_count = 0

    for input_file in input_files:
        output_file = args.output if args.output else None
        images_dir = args.images_dir if args.images_dir else None

        result = convert_file(
            input_file,
            output_file,
            images_dir,
            enable_llm=enable_llm,
            enable_omml=enable_omml,
            verbose=verbose
        )

        if result:
            success_count += 1
        else:
            fail_count += 1

    # Summary
    if verbose and len(input_files) > 1:
        print(f"\nSummary: {success_count} succeeded, {fail_count} failed")

    return 0 if fail_count == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
