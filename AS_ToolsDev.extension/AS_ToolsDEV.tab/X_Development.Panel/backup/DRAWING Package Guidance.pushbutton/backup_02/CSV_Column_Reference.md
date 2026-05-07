# Guidance Metadata CSV - Column Reference Guide

## Column Definitions

### **guidance_id** (Required)
- **Format:** Gxxx-yyy (e.g., G022-001)
- **Purpose:** Unique identifier for each guidance entry
- **Pattern:** G{CI/SfB_code}-{sequential_number}
- **Example:** G022-001 (First entry for CI/SfB code 22)

### **cisfb_code** (Required)
- **Format:** (xx) where xx is CI/SfB code
- **Purpose:** Structured categorization by building element
- **Examples:** 
  - (22) = Internal Walls/Partitions
  - (32) = Doors
  - (35) = Suspended Ceilings
  - (43) = Floor Finishes
  - (08) = Fire Strategy
- **Enables:** Hierarchical browsing, package grouping

### **title** (Required)
- **Format:** Free text (max 100 chars recommended)
- **Purpose:** Display title in guidance panel
- **Guidelines:** 
  - Start with element/package name
  - Include specific aspect (e.g., "Stage 4 Requirements")
  - Be descriptive but concise

### **discipline** (Required)
- **Format:** Single discipline code
- **Options:** 
  - Architecture
  - Structure
  - MEP
  - Landscape
  - Interiors
  - Fire
  - Acoustic
  - All (cross-disciplinary)
- **Purpose:** Filter by discipline specialty

### **riba_stage** (Required)
- **Format:** Single stage (3) or range (3-5)
- **Options:**
  - Single: 0, 1, 2, 3, 4, 5, 6, 7
  - Range: 2-3, 3-4, 3-5, 4-5, etc.
  - All: "All" (applies to all stages)
- **Purpose:** Filter by project stage

### **context_type** (Required)
- **Format:** Semicolon-separated list
- **Options:**
  - FloorPlan
  - CeilingPlan
  - Section
  - Elevation
  - ThreeDView
  - Detail
  - Schedule
  - Sheet
  - All
- **Examples:**
  - "FloorPlan;Section" (applies to both)
  - "Detail" (detail views only)
  - "All" (any view type)
- **Purpose:** Context-aware filtering by active view

### **related_packages** (Optional)
- **Format:** Semicolon-separated CI/SfB codes and/or keywords
- **Examples:**
  - "(32) Doors;(35) RCPs;BWIC"
  - "(08) Fire Strategy;(42) Wall Finishes"
- **Purpose:** Cross-reference related guidance, coordination reminders
- **Future:** Could enable clickable links to related guidance

### **nbs_specs** (Optional)
- **Format:** Semicolon-separated NBS specification codes
- **Examples:**
  - "K10;F10;G20"
  - "E10;E20;E30"
- **Purpose:** Link to specification requirements
- **Usage:** Quick reference for specification writers

### **keywords** (Required)
- **Format:** Comma-separated keywords (lowercase recommended)
- **Purpose:** Full-text search optimization
- **Guidelines:**
  - Include element type (partition, wall, door)
  - Include action words (coordination, setting out, detail)
  - Include stage references (stage3, stage4)
  - Include technical terms (acoustic, fire, cdp)
- **Examples:**
  - "partition,wall,internal,package,coordination"
  - "base detail,head detail,track,deflection"

### **pdf_file** (Required)
- **Format:** Filename only (not full path)
- **Purpose:** Link to full PDF guidance document
- **Guidelines:**
  - Use clear, descriptive filenames
  - Include package/topic in filename
  - Use underscores for spaces
- **Examples:**
  - "Partition_Types_Package_Guide.pdf"
  - "Drawing_Standards_RIBA_Stage_4.pdf"

### **page** (Optional)
- **Format:** Integer (page number)
- **Purpose:** Direct navigation to specific PDF page
- **Usage:** Future enhancement for PDF viewer integration
- **Current:** Display as reference only

### **brief_summary** (Required)
- **Format:** 1-2 sentences (max 200 chars)
- **Purpose:** Display in guidance panel cards
- **Guidelines:**
  - Concise, actionable
  - Answer "what does this cover?"
  - Include key deliverables or requirements

### **detailed_summary** (Optional but Recommended)
- **Format:** 2-4 sentences (max 500 chars)
- **Purpose:** Tooltip, expanded view, or search preview
- **Guidelines:**
  - Comprehensive overview
  - Include specific requirements
  - Mention critical coordination points
  - Reference related packages

---

## CSV Maintenance Best Practices

### Adding New Guidance
1. Determine CI/SfB code → Set guidance_id pattern
2. Choose appropriate discipline and RIBA stage(s)
3. List all applicable view contexts (use semicolons)
4. Identify related packages for coordination
5. Add relevant NBS specification codes
6. Write comprehensive keywords for searchability
7. Link to PDF (ensure file exists in Technical_Guidance folder)
8. Write brief summary for card display
9. Write detailed summary for full context

### Organizing Entries
- Group by CI/SfB code (all (22) entries together)
- Within each code, order by RIBA stage progression
- Use consistent guidance_id numbering

### Search Optimization
- Include synonyms in keywords (e.g., "partition,wall,internal")
- Add common misspellings if relevant
- Include both technical and colloquial terms
- Reference related codes in keywords for cross-discovery

### Quality Checks
- Verify PDF file exists at expected path
- Test keywords by searching for them in panel
- Ensure brief_summary displays cleanly (no line breaks)
- Check semicolon separators (no trailing semicolons)

---

## Example: Complete Guidance Entry

```csv
G022-004,(22),Wall Build-Up Details,Architecture,3-4,"Detail","(42) Wall Finishes;(08) Fire Strategy","K10;F10;F30;G20;P12","detail,wall type,build-up,fire,acoustic,pattress",Partition_Types_Package_Guide.pdf,5,"1:2 or 1:5 wall details showing linings, framework, pattressing, and performance specifications.","Detailed wall build-up drawings (1:2 or 1:5 scale) showing construction layers, framework systems, pattressing requirements, fire ratings, acoustic performance, and security specifications. Essential for accurate pricing and coordination."
```

**Breakdown:**
- **ID:** G022-004 (4th entry for CI/SfB 22)
- **Code:** (22) Internal Walls
- **Title:** Wall Build-Up Details
- **Discipline:** Architecture
- **Stage:** 3-4 (applicable to both Stage 3 and 4)
- **Context:** Detail views
- **Related:** Wall Finishes and Fire Strategy packages
- **NBS:** Multiple specification references
- **Keywords:** Searchable terms including technical vocabulary
- **PDF:** Links to main partition guide
- **Page:** Section starts on page 5
- **Brief:** Card display summary
- **Detailed:** Comprehensive tooltip/preview text

---

## CSV File Location
`I:\002_BIM\001_Document Reference Library\guidance_metadata.csv`

## Related PDF Location
`I:\002_BIM\001_Document Reference Library\Technical_Guidance\`
