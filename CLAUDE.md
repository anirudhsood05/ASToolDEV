# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Is

A **pyRevit extension** (`AS_ToolsDev.extension`) for Autodesk Revit. Scripts run inside Revit via the pyRevit framework using **IronPython 2.7** (CPython 3 is not available). There is no build step, test runner, or local execution — all tools execute inside the Revit host application.

## Extension Structure

pyRevit uses a strict folder-naming convention to auto-discover UI elements:

```
AS_ToolsDev.extension/
  AS_ToolsDEV.tab/          # Ribbon tab (bundle.yaml controls tab layout)
    Admin.Panel/            # Ribbon panel
      QA_QC.pulldown/       # Pulldown button group
        ToolName.pushbutton/
          script.py         # Main entry point — pyRevit runs this
          bundle.yaml       # Button metadata: title, tooltip, author, context
          icon.png          # Optional button icon
    X_Development.Panel/    # In-development / experimental tools
  lib/
    Snippets/               # Shared Revit API helpers (imported as `from Snippets import ...`)
  pyrevitlib/               # Vendored pyRevit library overrides (rpw, rjm, etc.)
  docs/
    pull_request_template.md
```

The `X_Development.Panel` holds work-in-progress tools. `Admin.Panel` holds production-ready tools. Both panels are inside the `AS_ToolsDEV` tab.

## Scripting Conventions

### Standard script header pattern
Every `script.py` starts with:
```python
# -*- coding: utf-8 -*-
__title__ = 'Button\nLabel'     # Two-line label shown on ribbon button
__doc__   = """Tooltip text."""
__author__ = "Author Name"

from pyrevit import revit, DB, forms, script
doc   = revit.doc
uidoc = revit.uidoc
app   = __revit__.Application
rvt_year = int(app.VersionNumber)  # Used for API version branching (e.g. < 2022)
```

### Key imports
- **Revit API**: `from Autodesk.Revit.DB import *` or specific imports via `from pyrevit import DB`
- **UI API**: `from Autodesk.Revit.UI import *`
- **WPF dialogs**: `clr.AddReference('PresentationFramework')` then `from System.Windows import ...`
- **Shared snippets**: `from Snippets._views import ...`, `from Snippets._elements import ...`, etc.
- **Transactions**: always wrap model changes in `with revit.Transaction("description"):` or explicit `Transaction(doc, "name")`

### IronPython 2.7 constraints
- No f-strings — use `.format()` or `%` formatting
- No type hints
- `print` is a statement: `print "text"` or use `output = script.get_output(); output.print_md(...)`
- `.NET` collections require `from System.Collections.Generic import List; lst = List[ElementId]()`

### Version branching
`FilterStringRule` constructor changed in Revit 2022 — always branch on `rvt_year`:
```python
if rvt_year < 2022:
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), value, caseSensitive)
else:
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), value)
```

## Shared Library (`lib/Snippets/`)

| Module | Contents |
|---|---|
| `_elements.py` | `dict_name_element()`, element type helpers |
| `_views.py` | `create_string_equals_filter()`, `create_3D_view()`, sheet↔view helpers |
| `_sheets.py` | `get_views_on_sheet()`, `get_titleblock_on_sheet()` |
| `_filters.py` / `_filtered_element_collector.py` | Collector shortcuts |
| `_selection.py` | Selection helpers |
| `_annotations.py`, `_text.py`, `_lines.py` | Annotation/detail element utilities |
| `_convert.py`, `_vectors.py`, `_boundingbox.py` | Geometry utilities |
| `_groups.py`, `_overrides.py`, `_revisions.py` | Group/graphics/revision helpers |
| `_excel.py` | Excel read/write via `openpyxl` |
| `_variables.py`, `_context_manager.py` | Misc utilities |

## `bundle.yaml` Keys

| Key | Purpose |
|---|---|
| `title` | Button label (use `\n` for two lines) |
| `tooltip` | Hover text |
| `author` | Author name |
| `context` | Revit context filter: `doc-project`, `active-section-view`, etc. |

## Development Workflow

Since scripts execute inside Revit, the development loop is:
1. Edit `script.py` (or shared lib files) in this repo.
2. In Revit, use the pyRevit **Reload** button (or `pyRevit > Reload`) to pick up changes without restarting Revit.
3. Click the button in the Revit ribbon to run.

New tools go in `X_Development.Panel` during development, then move to `Admin.Panel` (or a new panel) when production-ready. Keep `backup/` folders for old iterations — do not delete them.

## Branch

Active development branch: `claude/init-project-Vktor`
