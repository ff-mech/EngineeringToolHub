# SW Batch Update — VB Macro

SolidWorks VB macro for batch-updating custom properties and exporting DXF flat patterns.

## Files

| File | Description |
|------|-------------|
| `SoldworksBatchUpdate.bas` | VB macro — run inside SolidWorks |

## How to Run

1. Open SolidWorks
2. Go to **Tools → Macro → Run**
3. Browse to `SoldworksBatchUpdate.bas` and open it
4. Follow the prompts in the dialog

## What It Does

- Updates `DrawnBy` and `DwgDrawnBy` custom properties across a batch of parts
- Exports flat-pattern DXF files

> **Requirement:** SolidWorks must be installed. This macro cannot be run standalone.

## Notes

The Engineering Tool Hub **SW Batch Update** panel references this macro and displays
the file path and run instructions directly in the app UI.
