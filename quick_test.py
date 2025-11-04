"""Quick test harness for pptx_to_definition.
Run: python quick_test.py <presentation.pptx>
"""
from pathlib import Path
import sys
import json
from pptx_to_definition import extract_presentation_definition


def main():
    if len(sys.argv) < 2:
        print("Usage: python quick_test.py <presentation.pptx> [max_slides]")
        sys.exit(1)
    pptx_path = Path(sys.argv[1])
    max_slides = int(sys.argv[2]) if len(sys.argv) > 2 else 2

    if not pptx_path.exists():
        print(f"File not found: {pptx_path}")
        sys.exit(2)

    definition = extract_presentation_definition(pptx_path, max_slides=max_slides)
    summary = {
        "source": definition["source_file"],
        "slide_count_total": definition["metadata"]["slide_count"],
        "slides_processed": len(definition["slides"]),
        "first_slide_shapes": len(definition["slides"][0]["shapes"]) if definition["slides"] else 0,
    }
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
