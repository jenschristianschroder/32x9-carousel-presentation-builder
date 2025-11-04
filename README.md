# PowerPoint Carousel Presentation Builder

A complete toolkit for creating carousel-style PowerPoint presentations from existing decks.

## What It Does

1. **Extracts** slides from any PowerPoint as high-quality images
2. **Analyzes** presentation structure (shapes, text, layouts)
3. **Builds** carousel presentations matching a template design
4. **All-in-one** command for end-to-end conversion

## Features

### PPTX Structure Extractor (`pptx_to_definition.py`)
- Slide + shape geometry (inches)
- Text paragraphs and runs (font styles, colors, sizes)
- Tables (cell text, size)
- **Export each slide as a PNG image** (uses PowerPoint COM)
- Group shape children flattened
- Minimal chart stub metadata
- Notes text per slide
- Slide layout name + overall dimensions

## Install
```powershell
pip install -r requirements.txt
```

## Quick Start

### End-to-End Carousel Creation (Recommended)
```powershell
# One command does it all!
python create_carousel_end_to_end.py "Innovate with AI Apps and Agents - Deck.PPTX" "Carousel Presentation Template_definition.json" "Output_Carousel.pptx"

# Keep intermediate files for inspection
python create_carousel_end_to_end.py "MyDeck.pptx" "Carousel Presentation Template_definition.json" "Output.pptx" --keep-temp
```

## Individual Tools

### 1. Extract Slides & Definition
```powershell
# Export slides as images + create JSON definition
python pptx_to_definition.py "MyDeck.pptx" --export-images

# YAML output instead
python pptx_to_definition.py "MyDeck.pptx" -f yaml --export-images

# Limit to first 5 slides
python pptx_to_definition.py "MyDeck.pptx" --max-slides 5 --export-images
```

### 2. Build Carousel from Template
```powershell
# Using template design
python build_carousel_from_template.py "Carousel Presentation Template_definition.json" "MyDeck_images/" "Carousel_Output.pptx"
```

### 3. Simple Carousel (Grid Layout)
```powershell
# 4 slides per page in a grid
python build_carousel.py "MyDeck_images/" "GridCarousel.pptx" --slides-per-page 4

# 6 slides per page
python build_carousel.py "MyDeck_images/" "GridCarousel.pptx" --slides-per-page 6
```

### 4. Rebuild Presentation from Definition
```powershell
# Recreate presentation from definition + images
python build_from_definition.py "MyDeck_definition.json" "MyDeck_images/" "Rebuilt.pptx"

# Custom dimensions (32:9 format)
python build_from_definition.py "MyDeck_definition.json" "MyDeck_images/" "Rebuilt.pptx" --width 16 --height 9
```

## Output Schema (simplified)
```yaml
source_file: MyDeck.pptx
metadata:
  slide_width_inches: 16.0
  slide_height_inches: 9.0
  slide_count: 20
slides:
  - index: 1
    layout_name: Title Slide
    notes: "Welcome notes"
    shapes:
      - id: 3
        name: Title 1
        type: placeholder
        left_inches: 1.0
        top_inches: 1.2
        width_inches: 10.0
        height_inches: 1.4
        rotation: 0.0
        has_text_frame: true
        text: "Project Overview"
        paragraphs:
          - alignment: CENTER
            runs:
              - text: Project Overview
                font:
                  name: Calibri
                  size_pt: 44.0
                  bold: true
                  italic: false
                  underline: null
                  color_hex: "#FFFFFF"
```

## File Structure

```
32x9 carousel presentation builder/
├── pptx_to_definition.py              # Extract slides + structure
├── build_carousel_from_template.py    # Build using template design  
├── build_carousel.py                  # Build with grid layout
├── build_from_definition.py           # Rebuild from definition
├── create_carousel_end_to_end.py      # All-in-one script
├── quick_test.py                      # Quick validation tool
├── requirements.txt                   # Python dependencies
├── Carousel Presentation Template.pptx         # Template file
├── Carousel Presentation Template_definition.json  # Template design
└── README.md
```

## How It Works

1. **`pptx_to_definition.py`**: Uses PowerPoint COM automation to export each slide as a PNG image, and python-pptx to extract structure
2. **`build_carousel_from_template.py`**: Analyzes template layout and replicates the design with your slides
3. **`create_carousel_end_to_end.py`**: Orchestrates both steps automatically

## Template Design

The `Carousel Presentation Template.pptx` defines:
- Ultra-wide format (12.6" x 3.54")
- Horizontal scrolling carousel effect
- 3 slides visible per page: left (partial), center (full), right (partial)
- Decorative side panels
- Smooth progression (advances 1 slide at a time)

## Test Harness
```powershell
# Quick validation (first 2 slides by default)
python quick_test.py MyDeck.pptx

# Test first 5 slides
python quick_test.py MyDeck.pptx 5
```

## Requirements

- Python 3.7+
- Microsoft PowerPoint (for slide export via COM automation)
- Windows OS (due to COM automation dependency)

## Limitations

- Chart data values not extracted (python-pptx limitation)
- Advanced SmartArt, animations, media represented only by type
- Theme color palette not fully exposed
- Requires PowerPoint installed for slide export

## Troubleshooting

**PowerPoint COM Error**: Ensure PowerPoint is installed and you have appropriate permissions  
**Image Export Issues**: Check that the input PowerPoint file is not corrupted  
**Memory Issues**: For very large presentations, use `--max-slides` to process in batches

---
Created for flexible carousel presentation generation.
