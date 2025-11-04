# PowerPoint 32:9 Carousel Presentation Builder

Create stunning ultra-wide carousel presentations from any PowerPoint deck with a single command.

## Quick Start (One Command!)

```powershell
python create_carousel_end_to_end.py "MyPresentation.pptx" "Output_Carousel.pptx"
```

This single command:
1. ✅ Extracts all slides as high-quality PNG images
2. ✅ Builds an ultra-wide (12.6" × 3.54") carousel presentation with:
   - Black background
   - Gradient overlays (transparent to black on edges)
   - Each page centers a different slide with surrounding slides visible
   - Smooth horizontal scrolling effect

### Add Morph Transitions (Takes 3 seconds!)
1. Open `Output_Carousel.pptx` in PowerPoint
2. Select all slides (`Ctrl+A`)
3. Go to **Transitions** tab → Click **Morph**
4. Done! ✨

## Installation
```powershell
pip install -r requirements.txt
```

**Requirements:**
- Python 3.7+
- Microsoft PowerPoint (for slide export)
- Windows OS (PowerPoint COM automation)

## What You Get

The carousel presentation features:
- **Ultra-wide format**: 12.6" × 3.54" (32:9 aspect ratio)
- **Cinematic layout**: Center slide is prominent, surrounding slides fade with gradients
- **Progressive navigation**: Each page advances one slide through the carousel
- **Professional styling**: Black background with directional gradients on edges

## Advanced Usage

### Export Specific Slides
```powershell
# Export slides 3-12 only
python create_carousel_end_to_end.py "MyPresentation.pptx" "Output_Carousel.pptx" --range "3-12"

# Export from slide 10 to the end
python create_carousel_end_to_end.py "MyPresentation.pptx" "Output_Carousel.pptx" --range "10.."

# Export first 5 slides
python create_carousel_end_to_end.py "MyPresentation.pptx" "Output_Carousel.pptx" --range "..5"

# Export multiple ranges (slides 1-5 and 7-9)
python create_carousel_end_to_end.py "MyPresentation.pptx" "Output_Carousel.pptx" --range "..5,7-9"

# Export specific slides only
python create_carousel_end_to_end.py "MyPresentation.pptx" "Output_Carousel.pptx" --range "1,3,5,7"
```

### Keep Intermediate Files
```powershell
# Keep the extracted images and definition files for inspection
python create_carousel_end_to_end.py "MyPresentation.pptx" "Output_Carousel.pptx" --keep-temp
```

### Use Custom Template
```powershell
# Use a different template design
python create_carousel_end_to_end.py "MyPresentation.pptx" "Output_Carousel.pptx" --template "MyCustomTemplate_definition.json"
```

### Manual Two-Step Process

If you need more control, run the steps separately:

#### Step 1: Extract Slides
```powershell
python pptx_to_definition.py "MyPresentation.pptx" --export-images
```

#### Step 2: Build Carousel
```powershell
python build_carousel_from_template.py "Carousel Presentation Template_definition.json" "MyPresentation_images/" "Output_Carousel.pptx"
```

### Extract Options
```powershell
# Export specific slide range
python pptx_to_definition.py "MyDeck.pptx" --export-images --range "5-15"

# Export as YAML instead of JSON
python pptx_to_definition.py "MyDeck.pptx" -f yaml --export-images
```

### Alternative Layouts

#### Grid-Based Carousel
```powershell
# 4 slides per page in a grid
python build_carousel.py "MyDeck_images/" "GridCarousel.pptx" --slides-per-page 4

# 6 slides per page
python build_carousel.py "MyDeck_images/" "GridCarousel.pptx" --slides-per-page 6
```

#### Rebuild from Definition
```powershell
# Recreate original presentation from definition + images
python build_from_definition.py "MyDeck_definition.json" "MyDeck_images/" "Rebuilt.pptx"
```

## Project Structure

```
32x9 carousel presentation builder/
├── pptx_to_definition.py              # Extract slides as images
├── build_carousel_from_template.py    # Build carousel from template  
├── build_carousel.py                  # Alternative grid layout
├── build_from_definition.py           # Rebuild from definition
├── create_carousel_end_to_end.py      # All-in-one script
├── requirements.txt                   # Python dependencies
├── Carousel Presentation Template.pptx            # Template file
├── Carousel Presentation Template_definition.json # Template structure
└── README.md                          # This file
```

## Output Definition Schema (Optional)

The extractor can also create JSON/YAML definitions with full presentation structure:

```yaml
source_file: MyDeck.pptx
metadata:
  slide_width_inches: 10.0
  slide_height_inches: 7.5
  slide_count: 21
slides:
  - index: 1
    layout_name: Title Slide
    slide_image: slide_001.png
    shapes:
      - name: Title 1
        type: placeholder
        left_inches: 1.0
        top_inches: 1.2
        width_inches: 8.0
        height_inches: 1.5
        text: "My Presentation"
```

This definition enables advanced scenarios like rebuilding presentations or analyzing structure.

## How It Works

1. **Slide Export**: Uses PowerPoint COM automation to export each slide as a high-quality PNG image
2. **Template Analysis**: Reads the carousel template design (positions, gradients, layout)
3. **Carousel Generation**: Creates ultra-wide presentation with slides arranged in carousel pattern
4. **Manual Polish**: Apply Morph transitions in PowerPoint for smooth animations (3 seconds!)

## Why Manual Morph Transitions?

The `python-pptx` library doesn't support adding transitions to slides. The fastest, most reliable method is:
- Select all slides (Ctrl+A) → Transitions tab → Click Morph
- Takes literally 3 seconds and works perfectly every time

### Known Issue: Morph Transition Behavior

**Issue**: When input slides are very similar, PowerPoint's Morph transition may morph individual slide images instead of creating the smooth carousel scrolling effect.

**Cause**: PowerPoint detects similar content across slides and tries to morph matching elements rather than the overall layout.

**Solution**: Ensure your input slides are visually distinct enough:
- Use different layouts, colors, or content on each slide
- Avoid slides with identical backgrounds or repeated elements
- Add unique visual elements (logos, page numbers, different images) to each slide

**Alternative**: If slides must be similar, you can manually adjust the Morph transition settings in PowerPoint or use a Fade transition instead for consistent animation.

## Template Design

The included `Carousel Presentation Template.pptx` defines:
- **Format**: 12.6" × 3.54" ultra-wide (32:9)
- **Layout**: Horizontal carousel with 3-5 slides visible per page
- **Styling**: Black background with transparent-to-black gradients on edges
- **Pattern**: Each page centers a different slide, creating smooth scrolling effect

### Creating Your Own Template

To create a custom carousel template:

1. **Design your template** in PowerPoint with your desired layout, dimensions, and styling
2. **Generate the template definition**:
   ```powershell
   python pptx_to_definition.py "MyCustomTemplate.pptx"
   ```
   This creates `MyCustomTemplate_definition.json`
3. **Use your custom template**:
   ```powershell
   python create_carousel_end_to_end.py "Input.pptx" "Output.pptx" --template "MyCustomTemplate_definition.json"
   ```

**Template Requirements:**
- Include sample slides showing the carousel pattern progression
- Each slide should contain picture placeholders positioned where slide images will appear
- Optional: Add gradient rectangles or decorative elements that will be replicated

## Troubleshooting

**"Package not found" or file can't be opened**: Input PowerPoint files must have **General** sensitivity label (no encryption). Remove sensitivity labels from files before processing.  
**"PowerPoint COM Error"**: Ensure PowerPoint is installed and not running in protected mode  
**"Permission denied saving file"**: Close the output file if it's already open in PowerPoint  
**"No images found"**: Make sure to use `--export-images` flag when extracting slides  
**"Module not found"**: Run `pip install -r requirements.txt` to install dependencies

## System Requirements

- **Python**: 3.7 or higher
- **PowerPoint**: Microsoft PowerPoint (any recent version)
- **OS**: Windows (required for PowerPoint COM automation)
- **Memory**: Sufficient for loading presentation images (~50MB per deck)

## Performance

- **Speed**: Convert a 21-slide deck to carousel in under 30 seconds
- **Transitions**: Add Morph transitions manually in 3 seconds (select all → Transitions → Morph)
- **File Size**: Output presentations are typically 5-15 MB depending on image quality and slide count

---

*Built with python-pptx and PowerPoint COM automation*
