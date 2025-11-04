"""
Add Morph transitions to all slides in a PowerPoint presentation.

This script modifies the XML directly to add morph transitions.

Usage:
    python add_morph_transitions.py input.pptx output.pptx
"""

import argparse
import sys
from pathlib import Path
import zipfile
import shutil
from lxml import etree
import tempfile


def add_morph_transitions(input_path: Path, output_path: Path):
    """Add morph transitions to all slides by modifying XML."""
    
    print(f"Processing presentation: {input_path}")
    
    # Create temporary directory for extraction
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        
        # Extract PPTX (which is a ZIP file)
        with zipfile.ZipFile(input_path, 'r') as zip_ref:
            zip_ref.extractall(temp_path)
        
        # Find all slide XML files
        slides_dir = temp_path / 'ppt' / 'slides'
        slide_files = sorted(slides_dir.glob('slide*.xml'))
        
        print(f"Adding Morph transitions to {len(slide_files)} slides...")
        
        # Namespaces
        ns = {
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
            'p159': 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'
        }
        
        # Register namespaces
        for prefix, uri in ns.items():
            etree.register_namespace(prefix, uri)
        
        # Process each slide
        for i, slide_file in enumerate(slide_files, 1):
            try:
                # Parse slide XML
                tree = etree.parse(str(slide_file))
                root = tree.getroot()
                
                # Remove existing transitions
                for trans in root.findall('.//{%s}transition' % ns['p']):
                    root.remove(trans)
                
                # Create morph transition with AlternateContent structure
                # (same as template)
                alt_content = etree.Element('{%s}AlternateContent' % ns['mc'])
                
                # Choice element (for PowerPoint 2016+)
                choice = etree.SubElement(alt_content, '{%s}Choice' % ns['mc'])
                choice.set('Requires', 'p159')
                
                transition_choice = etree.SubElement(choice, '{%s}transition' % ns['p'])
                transition_choice.set('spd', 'slow')
                transition_choice.set('{%s}dur' % ns['p14'], '2000')
                
                morph = etree.SubElement(transition_choice, '{%s}morph' % ns['p159'])
                morph.set('option', 'byObject')
                
                # Fallback element (for older PowerPoint versions)
                fallback = etree.SubElement(alt_content, '{%s}Fallback' % ns['mc'])
                
                transition_fallback = etree.SubElement(fallback, '{%s}transition' % ns['p'])
                transition_fallback.set('spd', 'slow')
                
                fade = etree.SubElement(transition_fallback, '{%s}fade' % ns['p'])
                
                # Insert after clrMapOvr element
                clr_map = root.find('.//{%s}clrMapOvr' % ns['p'])
                if clr_map is not None:
                    clr_map_idx = list(root).index(clr_map)
                    root.insert(clr_map_idx + 1, alt_content)
                
                # Write back
                tree.write(str(slide_file), encoding='UTF-8', xml_declaration=True, standalone=True)
                
                if i % 5 == 0:
                    print(f"  Processed {i}/{len(slide_files)} slides...")
                
            except Exception as e:
                print(f"  Warning: Could not process slide {i}: {e}")
        
        print(f"Saving presentation: {output_path}")
        
        # Repack as PPTX
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for file_path in temp_path.rglob('*'):
                if file_path.is_file():
                    arcname = str(file_path.relative_to(temp_path))
                    zip_out.write(file_path, arcname)
        
        print(f"✅ Morph transitions added successfully!")
        print(f"   Output: {output_path}")
    
    return 0


def parse_args():
    parser = argparse.ArgumentParser(
        description="Add Morph transitions to PowerPoint presentation"
    )
    parser.add_argument(
        "input",
        help="Input PowerPoint file (.pptx)"
    )
    parser.add_argument(
        "output",
        help="Output PowerPoint file (.pptx)"
    )
    return parser.parse_args()


def main():
    args = parse_args()
    
    input_path = Path(args.input)
    output_path = Path(args.output)
    
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        return 1
    
    return add_morph_transitions(input_path, output_path)


if __name__ == "__main__":
    sys.exit(main())
