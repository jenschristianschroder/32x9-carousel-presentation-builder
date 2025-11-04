"""
End-to-end carousel builder: Extract slides and build carousel in one command.

This script:
1. Exports slides from input PowerPoint as images
2. Extracts the definition
3. Builds a carousel presentation based on template design

Usage:
    python create_carousel_end_to_end.py input.pptx template.json output_carousel.pptx
"""

import argparse
import json
import shutil
from pathlib import Path
import subprocess
import sys


def run_command(cmd: list, description: str) -> bool:
    """Run a command and return success status."""
    print(f"\n{'='*60}")
    print(f"{description}")
    print(f"{'='*60}")
    try:
        result = subprocess.run(cmd, check=True, capture_output=False)
        return result.returncode == 0
    except subprocess.CalledProcessError as e:
        print(f"Error: Command failed with exit code {e.returncode}")
        return False


def main():
    parser = argparse.ArgumentParser(
        description="End-to-end carousel builder from PowerPoint"
    )
    parser.add_argument(
        "input_pptx",
        help="Input PowerPoint file to convert to carousel"
    )
    parser.add_argument(
        "template_definition",
        help="Template definition JSON file (e.g., 'Carousel Presentation Template_definition.json')"
    )
    parser.add_argument(
        "output_pptx",
        help="Output carousel PowerPoint file"
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep temporary files (definition and images)"
    )
    
    args = parser.parse_args()
    
    input_path = Path(args.input_pptx)
    template_path = Path(args.template_definition)
    output_path = Path(args.output_pptx)
    
    # Validate inputs
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        return 1
    
    if not template_path.exists():
        print(f"Error: Template definition not found: {template_path}")
        return 1
    
    # Derive temp paths
    definition_path = input_path.parent / f"{input_path.stem}_definition.json"
    images_folder = input_path.parent / f"{input_path.stem}_images"
    
    print(f"\n{'='*60}")
    print(f"CAROUSEL BUILDER - End-to-End")
    print(f"{'='*60}")
    print(f"Input: {input_path}")
    print(f"Template: {template_path}")
    print(f"Output: {output_path}")
    print(f"Temp definition: {definition_path}")
    print(f"Temp images: {images_folder}")
    
    # Step 1: Export slides as images and create definition
    if not run_command(
        [sys.executable, "pptx_to_definition.py", str(input_path), "--export-images"],
        "Step 1: Exporting slides as images..."
    ):
        return 1
    
    # Step 2: Build carousel from template
    if not run_command(
        [sys.executable, "build_carousel_from_template.py", 
         str(template_path), str(images_folder), str(output_path)],
        "Step 2: Building carousel presentation..."
    ):
        return 1
    
    # Step 3: Cleanup temp files if requested
    if not args.keep_temp:
        print(f"\n{'='*60}")
        print("Step 3: Cleaning up temporary files...")
        print(f"{'='*60}")
        
        try:
            if definition_path.exists():
                definition_path.unlink()
                print(f"  Deleted: {definition_path}")
            
            if images_folder.exists():
                shutil.rmtree(images_folder)
                print(f"  Deleted: {images_folder}/")
        except Exception as e:
            print(f"  Warning: Could not clean up temp files: {e}")
    else:
        print(f"\n✓ Keeping temporary files:")
        print(f"  - {definition_path}")
        print(f"  - {images_folder}/")
    
    print(f"\n{'='*60}")
    print(f"✅ SUCCESS!")
    print(f"{'='*60}")
    print(f"Carousel presentation created: {output_path}")
    print()
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
