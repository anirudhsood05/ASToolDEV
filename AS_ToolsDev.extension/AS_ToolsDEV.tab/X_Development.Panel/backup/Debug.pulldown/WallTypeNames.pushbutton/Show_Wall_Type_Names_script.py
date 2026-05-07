# -*- coding: utf-8 -*-
"""
DIAGNOSTIC: Show Wall Type Names
=================================
Run this to see what your walls are actually called
"""

__title__ = "Show Wall\nType Names"
__author__ = "Anirudh Sood"

from pyrevit import revit, DB, script
from pyrevit.revit.db import query
from pychilizer import database

output = script.get_output()
doc = revit.doc

output.print_md("# Wall Type Names in Project")
output.print_md("---")

# Get all walls
walls = DB.FilteredElementCollector(doc)\
    .OfCategory(DB.BuiltInCategory.OST_Walls)\
    .WhereElementIsNotElementType()\
    .ToElements()

output.print_md("Found {} walls in project".format(len(walls)))
output.print_md("")

# Get unique type names
type_names = set()
for wall in walls:
    try:
        wall_type = query.get_type(wall)
        if wall_type:
            type_name = database.get_name(wall_type)
            type_names.add(type_name)
    except:
        pass

output.print_md("## Unique Wall Type Names ({} types):".format(len(type_names)))
output.print_md("")

# Sort and display
for type_name in sorted(type_names):
    # Check if matches filter
    matches_partition = "AUK_Wall_Partition_" in type_name
    matches_auk = "AUK_Wall_" in type_name
    
    marker = ""
    if matches_partition:
        marker = " ✓ MATCHES 'AUK_Wall_Partition_'"
    elif matches_auk:
        marker = " ⚠ MATCHES 'AUK_Wall_' (broader)"
    else:
        marker = " ✗ NO MATCH"
    
    output.print_md("- `{}`{}".format(type_name, marker))

output.print_md("")
output.print_md("---")
output.print_md("**Recommendation:**")
output.print_md("If most walls show '✗ NO MATCH', the filter pattern needs updating.")
