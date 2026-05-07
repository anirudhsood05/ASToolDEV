# CSV Migration Guide - Simple to Enhanced Structure

## Overview
This guide helps you migrate from the basic CSV structure to the enhanced structure that supports CI/SfB codes, related packages, NBS specifications, and more.

## What's New

### New Columns Added:
1. **cisfb_code** - Building element classification (e.g., (22), (32))
2. **discipline** - Architecture, MEP, Structure, etc.
3. **related_packages** - Cross-references to related guidance (semicolon-separated)
4. **nbs_specs** - NBS specification codes (semicolon-separated)
5. **brief_summary** - Short summary for card display
6. **detailed_summary** - Comprehensive summary for tooltips

### Enhanced Columns:
- **context_type** - Now supports multiple values (FloorPlan;Section;Detail)
- **riba_stage** - Now supports ranges (3-5 means stages 3, 4, AND 5)

## Migration Steps

### Step 1: Add New Columns
Open your existing CSV in Excel and add these columns (in order):

```
guidance_id,cisfb_code,title,discipline,riba_stage,context_type,related_packages,nbs_specs,keywords,pdf_file,page,brief_summary,detailed_summary
```

### Step 2: Fill CI/SfB Codes
Based on your guidance topic, assign CI/SfB codes:

| Code | Element |
|------|---------|
| (08) | Fire Strategy |
| (22) | Internal Walls/Partitions |
| (32) | Doors |
| (35) | Suspended Ceilings |
| (42) | Wall Finishes |
| (43) | Floor Finishes |

**Example:**
- Partition guidance → `(22)`
- Door guidance → `(32)`
- Multi-topic → Use most relevant code

### Step 3: Assign Disciplines
Set discipline for each entry:
- Architecture (most common)
- Structure
- MEP
- Landscape
- Fire
- Acoustic
- All (cross-disciplinary)

### Step 4: Enhance Context Types
Change single values to semicolon-separated lists:

**Old:** `FloorPlan`  
**New:** `FloorPlan;Section` (if applicable to multiple view types)

**Options:**
- FloorPlan
- CeilingPlan
- Section
- Elevation
- ThreeDView
- Detail
- Schedule
- Sheet
- All

### Step 5: Add Related Packages
List packages that must be coordinated:

**Format:** `(32) Doors;(35) RCPs;BWIC`

**Common related packages for partitions:**
- (32) Doors
- (35) RCPs (Reflected Ceiling Plans)
- (42) Wall Finishes
- (43) Floor Finishes
- (08) Fire Strategy
- BWIC (Built Work in Connection)

### Step 6: Add NBS Specifications
List relevant NBS codes:

**Format:** `K10;F10;G20`

**Common NBS codes:**
- K10 - Dry Lining
- F10 - Brick/Block Walling
- F30 - Accessories to Brick/Block
- G20 - Carpentry/Timber Framing
- P12 - Fire Stopping
- P20 - Unframed Isolated Trims
- E10 - In-situ Concrete
- E20 - Formwork

### Step 7: Split Summary into Brief and Detailed
Your existing `summary` column should be split:

**brief_summary** (1-2 sentences):
- Displays in panel cards
- Max ~200 characters
- Focus on key deliverable

**detailed_summary** (2-4 sentences):
- Future tooltips/expanded view
- Max ~500 characters
- Include specific requirements and coordination points

**Example:**
```
OLD summary: "Define partition arrangements and details with special attention to interfaces with doors, ceilings, finishes, and fire strategy. Complete guide covering plans, details, and specifications."

NEW brief_summary: "Define partition arrangements and details with special attention to interfaces with doors, ceilings, finishes, and fire strategy."

NEW detailed_summary: "Complete guide to (22) Internal Walls package preparation. Covers plan layouts (1:100/1:50), wall build-up details (1:2/1:5), setting out coordination, and critical interfaces with related packages. Includes Stage 3-5 requirements and CDP considerations."
```

## Example Migration

### Before (Simple CSV):
```csv
guidance_id,title,context_type,riba_stage,keywords,pdf_file,page,summary
G001,Partition Type Selection,FloorPlan,2-3,"partition,wall,types",Partition_Guide.pdf,5,"Choose partition types based on acoustic requirements and fire ratings. Standard types available in template."
```

### After (Enhanced CSV):
```csv
guidance_id,cisfb_code,title,discipline,riba_stage,context_type,related_packages,nbs_specs,keywords,pdf_file,page,brief_summary,detailed_summary
G022-001,(22),Partition Type Selection,Architecture,2-3,"FloorPlan;Detail","(08) Fire Strategy;(42) Wall Finishes","K10;F10","partition,wall,types,acoustic,fire",Partition_Guide.pdf,5,"Choose partition types based on acoustic requirements and fire ratings.","Guide to selecting appropriate partition types based on performance requirements including acoustic ratings, fire compartmentation, and structural considerations. References standard template types and coordination with finishes and services."
```

## Quick Tips

### Semicolon-Separated Values
Always use semicolons (`;`) to separate multiple values:
- context_type: `FloorPlan;Section`
- related_packages: `(32) Doors;(35) RCPs`
- nbs_specs: `K10;F10;G20`

**No trailing semicolons:** `K10;F10` not `K10;F10;`

### RIBA Stage Ranges
- Single stage: `3`
- Range: `3-5` (includes 3, 4, AND 5)
- All stages: `All`

### Guidance IDs
Update IDs to include CI/SfB code:
- Old: `G001, G002, G003`
- New: `G022-001, G022-002, G032-001`

Pattern: `G{cisfb_code_without_brackets}-{sequential_number}`

## Testing Your Migration

1. Save enhanced CSV to: `I:\002_BIM\001_Document Reference Library\guidance_metadata.csv`
2. Open Guidance Panel in Revit
3. Click "Refresh" button
4. Verify:
   - CI/SfB codes display as badges
   - Related packages show
   - NBS specs display
   - Search works across new fields
   - Brief summaries display cleanly

## Backward Compatibility

The enhanced CSV structure is **NOT backward compatible** with the simple structure. You must update the Python script to use the new version.

**Required:** Use the updated `guidance_panel.py` script that includes enhanced CSV parsing.

## Need Help?

- See `CSV_Column_Reference.md` for detailed column specifications
- See `guidance_metadata_optimized.csv` for complete examples
- Test with a few entries first before migrating entire database

## Column Order (Important)

Excel columns must be in this exact order:
1. guidance_id
2. cisfb_code
3. title
4. discipline
5. riba_stage
6. context_type
7. related_packages
8. nbs_specs
9. keywords
10. pdf_file
11. page
12. brief_summary
13. detailed_summary
