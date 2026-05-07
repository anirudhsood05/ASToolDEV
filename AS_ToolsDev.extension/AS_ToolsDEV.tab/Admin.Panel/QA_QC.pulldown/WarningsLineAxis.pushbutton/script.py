# -*- coding: utf-8 -*-
"""Analyze Inaccurate Line Warnings
Identifies which element types and views are causing warnings.
Helps decide if manual cleanup is worthwhile.
"""
__title__ = "Analyze\nWarning Sources"
__author__ = "Ani"

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from pyrevit import script
import clr
clr.AddReference('System')
from collections import defaultdict

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()


def analyze_warnings():
    """Analyze which elements and views are causing inaccuracy warnings"""
    
    # Get all warnings
    warnings = doc.GetWarnings()
    
    # Filter for inaccuracy warnings
    inaccurate_warnings = [w for w in warnings if 
                          'off axis' in w.GetDescriptionText().lower() or
                          'inaccuracies' in w.GetDescriptionText().lower()]
    
    if not inaccurate_warnings:
        output.print_md("# No Inaccuracy Warnings Found")
        return
    
    output.print_md("# Inaccurate Line Warning Analysis")
    output.print_md("**Total warnings:** {}".format(len(inaccurate_warnings)))
    output.print_md("\n---\n")
    
    # Categorize by element type
    category_counts = defaultdict(int)
    family_counts = defaultdict(int)
    elements_by_category = defaultdict(list)
    views_by_category = defaultdict(set)
    detail_line_views = defaultdict(list)
    
    for warning in inaccurate_warnings:
        element_ids = warning.GetFailingElements()
        
        for elem_id in element_ids:
            try:
                elem = doc.GetElement(elem_id)
                if elem:
                    category = elem.Category.Name if elem.Category else "Unknown"
                    category_counts[category] += 1
                    elements_by_category[category].append(elem_id.IntegerValue)
                    
                    # Get view ownership for view-specific elements
                    owner_view_id = elem.OwnerViewId
                    if owner_view_id and owner_view_id.IntegerValue != -1:
                        view = doc.GetElement(owner_view_id)
                        if view:
                            view_name = view.Name
                            views_by_category[category].add(view_name)
                            
                            # Track detail lines by view
                            if "Lines" in category or "Detail" in category:
                                detail_line_views[view_name].append(elem_id.IntegerValue)
                    
                    # Get family name if applicable
                    if hasattr(elem, 'Symbol'):
                        family = elem.Symbol.Family.Name
                        family_counts[family] += 1
            except:
                category_counts["Unknown/Deleted"] += 1
    
    # Display breakdown by category
    output.print_md("## Warnings by Category:")
    for category, count in sorted(category_counts.items(), 
                                  key=lambda x: x[1], reverse=True):
        output.print_md("- **{}**: {} warning(s)".format(category, count))
        
        # Show element IDs for this category (first 10)
        elem_ids = elements_by_category[category][:10]
        if elem_ids:
            output.print_md("  - Element IDs: {}{}".format(
                ', '.join(str(e) for e in elem_ids),
                ' ...' if len(elements_by_category[category]) > 10 else ''
            ))
        
        # Show views containing this category
        if category in views_by_category and views_by_category[category]:
            view_list = sorted(list(views_by_category[category]))
            if len(view_list) <= 5:
                output.print_md("  - Found in views: {}".format(', '.join(view_list)))
            else:
                output.print_md("  - Found in {} views: {} ...".format(
                    len(view_list), 
                    ', '.join(view_list[:5])
                ))
    
    # Detail Lines by View section
    if detail_line_views:
        output.print_md("\n---\n")
        output.print_md("## Detail Lines by View:")
        output.print_md("*Views containing off-axis detail lines (safe to fix)*\n")
        
        for view_name, elem_ids in sorted(detail_line_views.items(), 
                                          key=lambda x: len(x[1]), 
                                          reverse=True):
            output.print_md("### {}".format(view_name))
            output.print_md("- **Count**: {} detail line(s)".format(len(elem_ids)))
            
            # Show first 20 element IDs
            if len(elem_ids) <= 20:
                output.print_md("- **Element IDs**: {}".format(', '.join(str(e) for e in elem_ids)))
            else:
                output.print_md("- **Element IDs**: {} ... +{} more".format(
                    ', '.join(str(e) for e in elem_ids[:20]),
                    len(elem_ids) - 20
                ))
            output.print_md("")
    
    output.print_md("\n---\n")
    
    # Families breakdown if applicable
    if family_counts:
        output.print_md("## Warnings by Family:")
        for family, count in sorted(family_counts.items(), 
                                    key=lambda x: x[1], reverse=True)[:20]:
            output.print_md("- **{}**: {} warning(s)".format(family, count))
    
    output.print_md("\n---\n")
    
    # Recommendations
    output.print_md("## Recommendations:")
    
    # Check for detail lines
    if "Lines" in category_counts or "Detail Items" in category_counts:
        detail_count = category_counts.get("Lines", 0) + category_counts.get("Detail Items", 0)
        view_count = len(detail_line_views)
        
        output.print_md("### Detail Lines Detected")
        output.print_md("- **Count**: {} warning(s) across {} view(s)".format(detail_count, view_count))
        output.print_md("- **Safe to fix**: Detail lines can be deleted/recreated")
        output.print_md("- **Risk**: Low - no dependencies")
        
        if view_count <= 10:
            output.print_md("- **Action**: Consider manual cleanup in these {} views".format(view_count))
        else:
            output.print_md("- **Action**: Use suppression (too many views for manual cleanup)")
    
    # Check for model lines
    if "Model Lines" in category_counts:
        output.print_md("### Model Lines Detected")
        output.print_md("- **Cannot fix**: Model lines in families not editable in project")
        output.print_md("- **Action**: Edit source families or suppress warnings")
    
    # Check for walls/floors/roofs
    model_elements = ["Walls", "Floors", "Roofs", "Ceilings"]
    has_model = any(cat in category_counts for cat in model_elements)
    
    if has_model:
        output.print_md("### Model Elements Detected")
        model_count = sum(category_counts.get(cat, 0) for cat in model_elements)
        output.print_md("- **Count**: {} model element warning(s)".format(model_count))
        output.print_md("- **High risk**: Walls/Floors/Roofs have dependencies")
        output.print_md("- **Recommendation**: Suppress warnings instead of rebuilding")
    
    output.print_md("\n---\n")
    
    # Final verdict
    detail_only = ("Lines" in category_counts or "Detail Items" in category_counts) and not has_model
    view_count = len(detail_line_views)
    
    if detail_only and view_count <= 5:
        output.print_md("**Verdict**: Manual cleanup feasible - only {} view(s) with detail lines".format(view_count))
    elif detail_only and view_count <= 20:
        output.print_md("**Verdict**: Manual cleanup possible but tedious - {} views affected".format(view_count))
    else:
        output.print_md("**Verdict**: Suppression recommended - too many views or risky element types")


if __name__ == '__main__':
    analyze_warnings()
