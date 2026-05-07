# -*- coding: utf-8 -*-
"""Diagnose Window References
Debug tool to understand what references are available on window families.
"""

__title__ = "Diagnose\nWindow Refs"
__author__ = "AUK Digital"

from pyrevit import revit, DB, forms, script

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


def diagnose_window(window):
    """Comprehensive diagnostic of window references."""
    
    output.print_md("# Window: {}".format(window.Symbol.FamilyName))
    output.print_md("**ID**: {}".format(output.linkify(window.Id)))
    output.print_md("---\n")
    
    # Test all reference types
    ref_types = [
        ("Left", DB.FamilyInstanceReferenceType.Left),
        ("CenterLeftRight", DB.FamilyInstanceReferenceType.CenterLeftRight),
        ("Right", DB.FamilyInstanceReferenceType.Right),
        ("Front", DB.FamilyInstanceReferenceType.Front),
        ("CenterFrontBack", DB.FamilyInstanceReferenceType.CenterFrontBack),
        ("Back", DB.FamilyInstanceReferenceType.Back),
        ("Bottom", DB.FamilyInstanceReferenceType.Bottom),
        ("CenterElevation", DB.FamilyInstanceReferenceType.CenterElevation),
        ("Top", DB.FamilyInstanceReferenceType.Top),
        ("StrongReference", DB.FamilyInstanceReferenceType.StrongReference),
        ("WeakReference", DB.FamilyInstanceReferenceType.WeakReference),
    ]
    
    output.print_md("## Testing GetReferences() Method")
    for ref_name, ref_type in ref_types:
        try:
            references = window.GetReferences(ref_type)
            ref_count = len(references) if references else 0
            
            if ref_count > 0:
                output.print_md("✓ **{}**: {} reference(s) found".format(ref_name, ref_count))
                
                # Try to get geometry from first reference
                try:
                    ref = references[0]
                    geom_obj = window.GetGeometryObjectFromReference(ref)
                    if geom_obj:
                        geom_type = type(geom_obj).__name__
                        output.print_md("  - Geometry type: **{}**".format(geom_type))
                        
                        if isinstance(geom_obj, DB.Line):
                            start = geom_obj.GetEndPoint(0)
                            end = geom_obj.GetEndPoint(1)
                            output.print_md("  - Line from ({:.2f}, {:.2f}, {:.2f}) to ({:.2f}, {:.2f}, {:.2f})".format(
                                start.X, start.Y, start.Z, end.X, end.Y, end.Z
                            ))
                    else:
                        output.print_md("  - GetGeometryObjectFromReference returned None")
                except Exception as e:
                    output.print_md("  - Error getting geometry: {}".format(str(e)))
            else:
                output.print_md("✗ **{}**: No references".format(ref_name))
                
        except Exception as e:
            output.print_md("✗ **{}**: Error - {}".format(ref_name, str(e)))
    
    # Alternative approach - check location and bounding box
    output.print_md("\n## Alternative Geometry Sources")
    
    try:
        location = window.Location
        if isinstance(location, DB.LocationPoint):
            pt = location.Point
            output.print_md("✓ **LocationPoint**: ({:.2f}, {:.2f}, {:.2f})".format(pt.X, pt.Y, pt.Z))
        elif isinstance(location, DB.LocationCurve):
            output.print_md("✓ **LocationCurve**: Available")
    except Exception as e:
        output.print_md("✗ Location: {}".format(str(e)))
    
    try:
        bbox = window.get_BoundingBox(None)
        if bbox:
            output.print_md("✓ **BoundingBox**: Min({:.2f}, {:.2f}, {:.2f}) Max({:.2f}, {:.2f}, {:.2f})".format(
                bbox.Min.X, bbox.Min.Y, bbox.Min.Z,
                bbox.Max.X, bbox.Max.Y, bbox.Max.Z
            ))
    except Exception as e:
        output.print_md("✗ BoundingBox: {}".format(str(e)))
    
    # Check geometry options
    output.print_md("\n## Direct Geometry Analysis")
    try:
        geom_options = DB.Options()
        geom_options.ComputeReferences = True
        geom = window.get_Geometry(geom_options)
        
        if geom:
            output.print_md("✓ **Geometry Available**")
            geom_count = 0
            for geom_obj in geom:
                geom_count += 1
                geom_type = type(geom_obj).__name__
                output.print_md("  - Geometry object {}: **{}**".format(geom_count, geom_type))
                
                # If it's a GeometryInstance, get symbol geometry
                if isinstance(geom_obj, DB.GeometryInstance):
                    symbol_geom = geom_obj.GetSymbolGeometry()
                    if symbol_geom:
                        output.print_md("    - Symbol geometry available")
                        for sym_obj in symbol_geom:
                            output.print_md("      - **{}**".format(type(sym_obj).__name__))
        else:
            output.print_md("✗ No geometry available")
            
    except Exception as e:
        output.print_md("✗ Geometry error: {}".format(str(e)))
    
    output.print_md("\n" + "="*50 + "\n")


def main():
    """Main execution."""
    
    # Get selection
    selection = uidoc.Selection.GetElementIds()
    
    if not selection:
        forms.alert("Please select at least one window.", exitscript=True)
    
    # Filter to windows
    selected_elements = [doc.GetElement(id) for id in selection]
    windows = [elem for elem in selected_elements 
               if isinstance(elem, DB.FamilyInstance) 
               and elem.Category 
               and elem.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Windows)]
    
    if not windows:
        forms.alert("No windows in selection.", exitscript=True)
    
    output.print_md("# Window Reference Diagnostic Tool")
    output.print_md("**Total Windows Selected**: {}\n".format(len(windows)))
    output.print_md("="*50 + "\n")
    
    # Diagnose each window
    for window in windows:
        diagnose_window(window)
    
    output.print_md("# Diagnostic Complete")


if __name__ == "__main__":
    main()
