# AUK Technical Guidance Panel - Deployment Guide

## Overview
Context-aware guidance system that displays technical standards and workflows directly in Revit. Searches and filters guidance based on current view type and keywords.

## Quick Start

### 1. File Deployment
```
pyRevit Extension:
├── Foundry.extension/
│   └── Foundry.tab/
│       └── Documentation.panel/
│           └── Guidance Panel.pushbutton/
│               └── script.py  (rename guidance_panel.py to script.py)

Network Resources:
├── I:/002_BIM/001_Document Reference Library/
│   ├── guidance_metadata.csv
│   └── Technical_Guidance/
│       ├── Partition_Types_RIBA_Stage_3.pdf
│       ├── Drawing_Packages_Standards.pdf
│       └── [other PDF guides]
```

### 2. CSV Setup
1. Copy `guidance_metadata.csv` to: `I:\002_BIM\001_Document Reference Library\`
2. Update paths in script if your network drive differs:
   ```python
   GUIDANCE_CSV_PATH = r"I:\002_BIM\001_Document Reference Library\guidance_metadata.csv"
   GUIDANCE_PDF_FOLDER = r"I:\002_BIM\001_Document Reference Library\Technical_Guidance"
   ```

### 3. Add PDF Guides
- Place PDF files in `Technical_Guidance` folder
- Ensure filenames in CSV match actual PDF filenames exactly

## CSV Structure

### Required Columns:
- **guidance_id**: Unique identifier (e.g., G001, G002)
- **title**: Display title in panel
- **context_type**: View type filter (FloorPlan, Section, Sheet, ThreeDView, All)
- **riba_stage**: RIBA stage(s) (e.g., "2-3", "4", "All")
- **keywords**: Comma-separated search terms
- **pdf_file**: Filename only (not full path)
- **page**: Starting page number (optional, for future PDF viewer)
- **summary**: Brief description (2-3 sentences)

### Context Types:
Use Revit API ViewType names:
- `FloorPlan` - Floor plans and ceiling plans
- `Section` - Building sections
- `Elevation` - Elevations
- `ThreeDView` - 3D views
- `Sheet` - Sheets
- `Schedule` - Schedules
- `All` - Shows in all contexts

Multiple contexts: Use semicolon separator (e.g., `FloorPlan;Section`)

## Usage

### Opening the Panel
1. Click "Guidance Panel" button in Documentation panel
2. Panel opens as modeless window (stays on top)
3. Shows guidance relevant to current view context

### Search Features
- **Keyword search**: Enter terms in search box, click "Search"
- **Show All**: Clears search, shows all guidance
- **Refresh**: Reloads CSV (use after updating guidance database)

### Context Awareness
- Panel detects current view type automatically
- Top label shows: "Current context: FloorPlan - Level 1"
- Filter guidance by searching for view type keywords

### Opening PDFs
- Click "View Guide (PDF)" button to open full document
- Opens in default PDF viewer
- If PDF missing, shows error with expected path

## Maintenance Workflow

### Adding New Guidance
1. Create/update PDF guide
2. Save to `Technical_Guidance` folder
3. Add row to CSV:
   ```csv
   G011,New Topic,FloorPlan,3,"keywords,here",New_Guide.pdf,1,"Brief summary here."
   ```
4. Click "Refresh" in panel (no Revit restart needed)

### Updating Existing Guidance
1. Edit CSV row (change title, summary, keywords, etc.)
2. Save CSV
3. Click "Refresh" in panel

### Removing Guidance
1. Delete CSV row
2. Click "Refresh" in panel

## Advanced Customization

### Custom Network Paths
Edit script constants:
```python
GUIDANCE_CSV_PATH = r"X:\Your\Custom\Path\guidance_metadata.csv"
GUIDANCE_PDF_FOLDER = r"X:\Your\Custom\Path\PDFs"
```

### Panel Position/Size
Adjust XAML:
```xml
<Window ... Width="400" Height="600" 
        WindowStartupLocation="Manual" ...>
```

### Styling
Colors/fonts follow AUK Style Guide. To customize:
```xml
<Border Background="#F0F0F0">  <!-- Header background -->
<TextBlock Foreground="#666666">  <!-- Context label -->
```

## Troubleshooting

### Panel Opens But Shows "Loading..."
- Check CSV path exists: `I:\002_BIM\001_Document Reference Library\guidance_metadata.csv`
- Verify CSV format (use template provided)
- Check Python output window for errors

### PDFs Not Opening
- Verify PDF exists in `Technical_Guidance` folder
- Check filename in CSV matches actual file exactly (case-sensitive)
- Ensure Windows file associations set for .pdf files

### Search Returns No Results
- Check keywords column in CSV
- Search is case-insensitive and matches partial words
- Try broader search terms

### Context Label Blank
- Normal if view type not detected
- Panel still functional, shows all guidance

## Future Enhancements

### Phase 2 (Potential):
- Auto-filter by active view type (checkbox)
- Recent/favorite guidance tracking
- Direct page navigation in PDFs
- Usage analytics

### Phase 3 (Potential):
- AI-powered search integration
- Dynamic content generation
- Multi-language support
- Integration with project templates

## Support
For issues or feature requests, contact BIM team or check pyRevit output window for error messages.
