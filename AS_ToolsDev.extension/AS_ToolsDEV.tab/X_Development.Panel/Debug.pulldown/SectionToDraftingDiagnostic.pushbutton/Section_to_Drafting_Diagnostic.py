# -*- coding: utf-8 -*-
"""
Script:   Section to Drafting - Failure Diagnostic
Desc:     Analyses a section view and reports exactly which elements would
          fail recreation and why. Does NOT modify the model.
Author:   AUK BIM Team
Usage:    Run with a section view active or select one from the list.
Result:   Detailed diagnostic report in pyRevit output window.
"""
from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms

logger = script.get_logger()
output = script.get_output()
doc    = revit.doc
uidoc  = revit.uidoc


def eid_int(eid):
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def get_cat_info(el):
    """Return (cat_int, cat_name) or (None, 'No Category')."""
    if el.Category is None:
        return None, "No Category"
    try:
        cat_int = int(el.Category.Id.IntegerValue)
    except AttributeError:
        cat_int = int(el.Category.Id.Value)
    return cat_int, el.Category.Name


def main():
    # Get section view
    active = uidoc.ActiveView
    if active.ViewType == DB.ViewType.Section and not active.IsTemplate:
        section = active
    else:
        all_sections = DB.FilteredElementCollector(doc)\
                         .OfClass(DB.View).ToElements()
        candidates = [v for v in all_sections
                      if v.ViewType == DB.ViewType.Section
                      and not v.IsTemplate]
        if not candidates:
            forms.alert("No section views found.", exitscript=True)
        name_map = {v.Name: v for v in candidates}
        chosen = forms.SelectFromList.show(
            sorted(name_map.keys()),
            title="Select Section View to Diagnose",
            multiselect=False
        )
        if not chosen:
            script.exit()
        section = name_map[chosen]

    output.print_md("# Diagnostic: {}".format(section.Name))
    output.print_md("---")

    # Collect ALL view-specific elements
    all_in_view = DB.FilteredElementCollector(doc, section.Id)\
                    .WhereElementIsNotElementType()\
                    .ToElements()

    view_specific = [el for el in all_in_view if el.ViewSpecific]
    output.print_md("**Total ViewSpecific elements:** {}".format(
        len(view_specific)))

    # Categorise everything
    output.print_md("## All ViewSpecific Elements by Category")

    cat_elements = {}
    for el in view_specific:
        cat_int, cat_name = get_cat_info(el)
        if cat_name not in cat_elements:
            cat_elements[cat_name] = []
        cat_elements[cat_name].append(el)

    for cat_name in sorted(cat_elements.keys()):
        els = cat_elements[cat_name]
        output.print_md("- **{}**: {} element(s)".format(cat_name, len(els)))

    # Detailed analysis per category
    output.print_md("---")
    output.print_md("## Detailed Element Analysis")

    for cat_name in sorted(cat_elements.keys()):
        els = cat_elements[cat_name]
        output.print_md("### {} ({})".format(cat_name, len(els)))

        for el in els[:5]:  # Sample first 5
            eid = eid_int(el.Id)
            el_type = type(el).__name__

            # Group membership
            in_group = el.GroupId != DB.ElementId.InvalidElementId
            group_info = ""
            if in_group:
                group_el = doc.GetElement(el.GroupId)
                if group_el:
                    group_info = " | IN GROUP: '{}'".format(
                        DB.Element.Name.GetValue(group_el))
                else:
                    group_info = " | IN GROUP (id={})".format(
                        eid_int(el.GroupId))

            # Location info
            loc = el.Location
            loc_info = "No Location"
            if isinstance(loc, DB.LocationPoint):
                pt = loc.Point
                loc_info = "LocationPoint ({:.2f}, {:.2f}, {:.2f})".format(
                    pt.X, pt.Y, pt.Z)
            elif isinstance(loc, DB.LocationCurve):
                crv = loc.Curve
                p0 = crv.GetEndPoint(0)
                p1 = crv.GetEndPoint(1)
                loc_info = "LocationCurve ({:.1f},{:.1f})->({:.1f},{:.1f})".format(
                    p0.X, p0.Y, p1.X, p1.Y)
            elif loc is not None:
                loc_info = "Location type: {}".format(type(loc).__name__)

            # Geometry curve (for CurveElements)
            geom_info = ""
            if isinstance(el, DB.CurveElement):
                try:
                    crv = el.GeometryCurve
                    if crv:
                        crv_type = type(crv).__name__
                        geom_info = " | Curve: {} len={:.3f}".format(
                            crv_type, crv.Length)
                    else:
                        geom_info = " | GeometryCurve=None"
                except Exception as ex:
                    geom_info = " | GeometryCurve ERROR: {}".format(str(ex))

            # Line style (for CurveElements)
            style_info = ""
            if isinstance(el, DB.CurveElement):
                try:
                    ls = el.LineStyle
                    if ls:
                        style_info = " | Style: '{}'".format(
                            DB.Element.Name.GetValue(ls))
                except Exception:
                    style_info = " | Style: ERROR"

            # For Detail Components - check if line-based
            dc_info = ""
            if isinstance(el, DB.FamilyInstance):
                try:
                    fam_name = el.Symbol.Family.Name if el.Symbol else "?"
                    type_name = DB.Element.Name.GetValue(el.Symbol) if el.Symbol else "?"
                    dc_info = " | Family: '{}' Type: '{}'".format(
                        fam_name, type_name)
                    # Check if line-based
                    if isinstance(loc, DB.LocationCurve):
                        dc_info += " [LINE-BASED]"
                    elif isinstance(loc, DB.LocationPoint):
                        dc_info += " [POINT-BASED]"
                    else:
                        dc_info += " [NO LOCATION]"
                except Exception as ex:
                    dc_info = " | FamilyInfo ERROR: {}".format(str(ex))

            # For TextNotes
            tn_info = ""
            if isinstance(el, DB.TextNote):
                try:
                    text = el.Text
                    coord = el.Coord
                    tn_info = " | Text: '{}' Coord: ({:.2f},{:.2f},{:.2f})".format(
                        text[:30] if text else "EMPTY",
                        coord.X, coord.Y, coord.Z)
                except Exception as ex:
                    tn_info = " | TextNote ERROR: {}".format(str(ex))

            # For FilledRegions
            fr_info = ""
            if isinstance(el, DB.FilledRegion):
                try:
                    boundaries = el.GetBoundaries()
                    loop_count = boundaries.Count if boundaries else 0
                    fr_info = " | Loops: {}".format(loop_count)
                except Exception as ex:
                    fr_info = " | Boundaries ERROR: {}".format(str(ex))

            # For Groups
            grp_info = ""
            if isinstance(el, DB.Group):
                try:
                    members = el.GetMemberIds()
                    member_count = len(list(members)) if members else 0
                    grp_info = " | Members: {}".format(member_count)
                    # List member categories
                    member_cats = {}
                    for mid in members:
                        m = doc.GetElement(mid)
                        if m:
                            _, mcat = get_cat_info(m)
                            member_cats[mcat] = member_cats.get(mcat, 0) + 1
                    for mc, cnt in member_cats.iteritems():
                        grp_info += "\n      - {} x {}".format(cnt, mc)
                except Exception as ex:
                    grp_info = " | Group ERROR: {}".format(str(ex))

            output.print_md(
                "  - id={} type=**{}** | {}{}{}{}{}{}{}{}".format(
                    eid, el_type, loc_info,
                    group_info, geom_info, style_info,
                    dc_info, tn_info, fr_info, grp_info))

        if len(els) > 5:
            output.print_md("  ... and {} more".format(len(els) - 5))

    # Summary of potential issues
    output.print_md("---")
    output.print_md("## Potential Issues Summary")

    # Count elements in groups
    in_group_count = sum(
        1 for el in view_specific
        if el.GroupId != DB.ElementId.InvalidElementId
        and not isinstance(el, DB.Group))
    output.print_md("- Elements inside groups (handled via group): {}".format(
        in_group_count))

    # Count elements with no location
    no_loc_count = sum(
        1 for el in view_specific
        if el.Location is None and not isinstance(el, DB.Group)
        and not isinstance(el, DB.FilledRegion))
    output.print_md("- Elements with no Location property: {}".format(
        no_loc_count))

    # Count line-based detail components
    line_based_dc = sum(
        1 for el in view_specific
        if isinstance(el, DB.FamilyInstance)
        and isinstance(el.Location, DB.LocationCurve))
    output.print_md("- Line-based detail components: {}".format(line_based_dc))

    # Count CurveElements with None geometry
    null_geom = sum(
        1 for el in view_specific
        if isinstance(el, DB.CurveElement)
        and (not hasattr(el, 'GeometryCurve') or el.GeometryCurve is None))
    output.print_md("- CurveElements with null geometry: {}".format(null_geom))

    # Count elements with empty category name
    empty_cat = sum(
        1 for el in view_specific
        if el.Category is not None
        and el.Category.Name == "")
    output.print_md("- Elements with empty category name: {}".format(empty_cat))

    # List all .NET types present
    output.print_md("---")
    output.print_md("## .NET Types Present")
    type_counts = {}
    for el in view_specific:
        tn = type(el).__name__
        type_counts[tn] = type_counts.get(tn, 0) + 1
    for tn in sorted(type_counts.keys()):
        output.print_md("- **{}**: {}".format(tn, type_counts[tn]))


main()
