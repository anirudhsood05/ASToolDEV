# -*- coding: utf-8 -*-
"""CAD Block Name Discovery Diagnostic

Explores what identifiers are available on nested CAD
GeometryInstance objects to find usable block names.
Checks: GraphicsStyle subcategory names, internal Ids,
nested GI attributes, and geometry grouping patterns.
"""

__title__ = "CAD Block\nName Diag"
__author__ = "Aukett Swanke BIM"

from pyrevit import script, forms, revit, DB
import math

doc = revit.doc
output = script.get_output()

MAX_BLOCKS = 30  # Limit output for readability


def get_gs_full_info(geom_obj):
    """Get full GraphicsStyle info including category hierarchy."""
    info = {"name": "", "category": "", "subcategory": "", "id": -1}
    try:
        gs_id = geom_obj.GraphicsStyleId
        if gs_id and gs_id != DB.ElementId.InvalidElementId:
            gs = doc.GetElement(gs_id)
            if gs:
                info["name"] = gs.Name or ""
                info["id"] = gs_id.IntegerValue
                # GraphicsStyle has a GraphicsStyleCategory
                try:
                    cat = gs.GraphicsStyleCategory
                    if cat:
                        info["subcategory"] = cat.Name or ""
                        if cat.Parent:
                            info["category"] = cat.Parent.Name or ""
                except Exception:
                    pass
    except Exception:
        pass
    return info


def explore_block_gi(block_gi, depth=0, max_depth=3):
    """Recursively explore a block GeometryInstance for naming clues."""
    prefix = "  " * depth
    results = []

    # Basic transform info
    try:
        origin = block_gi.Transform.Origin
        rot = math.degrees(math.atan2(
            block_gi.Transform.BasisX.Y,
            block_gi.Transform.BasisX.X
        ))
        results.append("{}Origin: ({:.3f}, {:.3f}, {:.3f}) Rot: {:.1f}deg".format(
            prefix, origin.X, origin.Y, origin.Z, rot
        ))
    except Exception as e:
        results.append("{}Transform error: {}".format(prefix, str(e)))

    # GraphicsStyle (layer) - full detail
    gs_info = get_gs_full_info(block_gi)
    results.append("{}GraphicsStyle: name='{}' subcat='{}' cat='{}' id={}".format(
        prefix, gs_info["name"], gs_info["subcategory"],
        gs_info["category"], gs_info["id"]
    ))

    # Internal Id
    try:
        results.append("{}GeomInstance.Id: {}".format(prefix, block_gi.Id))
    except Exception:
        pass

    # Check all accessible attributes for naming clues
    interesting_attrs = []
    for attr_name in dir(block_gi):
        if attr_name.startswith('_'):
            continue
        if attr_name in ('Transform', 'GraphicsStyleId', 'Id',
                         'IsElementGeometry', 'IsReadOnly', 'Visibility',
                         'SymbolGeometry', 'GetHashCode', 'GetType',
                         'ToString', 'Equals', 'Dispose',
                         'GetInstanceGeometry', 'GetSymbolGeometry',
                         'ReferenceEquals', 'MemberwiseClone',
                         'Finalize'):
            continue
        try:
            val = getattr(block_gi, attr_name)
            if callable(val):
                continue
            interesting_attrs.append("{}  .{} = {}".format(prefix, attr_name, val))
        except Exception:
            pass

    if interesting_attrs:
        results.append("{}Other attributes:".format(prefix))
        results.extend(interesting_attrs)

    # Explore GetSymbolGeometry children
    if depth < max_depth:
        try:
            sym_geom = block_gi.GetSymbolGeometry()
            if sym_geom:
                child_types = {}
                nested_gis = []
                for child in sym_geom:
                    t = type(child).__name__
                    child_types[t] = child_types.get(t, 0) + 1
                    if isinstance(child, DB.GeometryInstance):
                        nested_gis.append(child)

                results.append("{}SymbolGeometry children: {}".format(
                    prefix, ", ".join("{}:{}".format(k, v)
                                      for k, v in sorted(child_types.items()))
                ))

                # Check GraphicsStyles of ALL children (not just GIs)
                child_layers = set()
                for child in block_gi.GetSymbolGeometry():
                    child_gs = get_gs_full_info(child)
                    if child_gs["name"] or child_gs["subcategory"]:
                        child_layers.add(
                            "name='{}' subcat='{}'".format(
                                child_gs["name"], child_gs["subcategory"])
                        )

                if child_layers:
                    results.append("{}Child layer names: {}".format(
                        prefix, " | ".join(sorted(child_layers))
                    ))

                # Recurse into nested GIs (first 3 only)
                for i, nested_gi in enumerate(nested_gis[:3]):
                    results.append("{}--- Nested GI [{}] ---".format(prefix, i))
                    results.extend(explore_block_gi(nested_gi, depth + 1, max_depth))

        except Exception as e:
            results.append("{}SymbolGeometry error: {}".format(prefix, str(e)))

    return results


def main():
    output.print_md("# CAD Block Name Discovery Diagnostic")
    output.print_md("---")

    # Collect ImportInstances
    imports = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.ImportInstance)
        .WhereElementIsNotElementType()
    )

    if not imports:
        output.print_md("No ImportInstance elements found.")
        return

    # Let user pick one
    names = []
    for imp in imports:
        try:
            t = doc.GetElement(imp.GetTypeId())
            n = t.Name if t else "Unknown"
        except Exception:
            n = "Unknown"
        link = "Linked" if imp.IsLinked else "Imported"
        names.append("{} [{}] (ID:{})".format(n, link, imp.Id.IntegerValue))

    selected = forms.SelectFromList.show(names, title="Select CAD file to diagnose")
    if not selected:
        return

    idx = names.index(selected)
    imp = imports[idx]

    output.print_md("**Selected**: {}".format(selected))
    output.print_md("")

    # Also check the ImportInstance's category/subcategories
    output.print_md("## ImportInstance Category Info")
    try:
        cat = imp.Category
        if cat:
            output.print_md("Category: {}".format(cat.Name))
            subcats = cat.SubCategories
            if subcats:
                output.print_md("Subcategories ({} total):".format(subcats.Size))
                for sc in subcats:
                    output.print_md("  - '{}' (Id: {})".format(sc.Name, sc.Id.IntegerValue))
            else:
                output.print_md("No subcategories found.")
    except Exception as e:
        output.print_md("Category access error: {}".format(str(e)))

    # Check CadLinkType subcategories (often contains layer->block mappings)
    output.print_md("")
    output.print_md("## CadLinkType / ImportType Subcategories")
    try:
        type_id = imp.GetTypeId()
        type_elem = doc.GetElement(type_id)
        if type_elem:
            output.print_md("Type: {} (Id: {})".format(type_elem.Name, type_id.IntegerValue))
            type_cat = type_elem.Category
            if type_cat and type_cat.SubCategories:
                output.print_md("Type subcategories ({} total):".format(
                    type_cat.SubCategories.Size))
                for sc in type_cat.SubCategories:
                    output.print_md("  - '**{}**' (Id: {})".format(
                        sc.Name, sc.Id.IntegerValue))
            else:
                output.print_md("No type subcategories.")
    except Exception as e:
        output.print_md("Type category error: {}".format(str(e)))

    # Now traverse geometry
    output.print_md("")
    output.print_md("## Block GeometryInstance Analysis")

    options = DB.Options()
    options.ComputeReferences = False
    options.IncludeNonVisibleObjects = False

    import_transform = imp.GetTotalTransform()

    try:
        geom_element = imp.get_Geometry(options)
    except Exception as e:
        output.print_md("**Error**: Could not read geometry: {}".format(str(e)))
        return

    block_count = 0
    for top_obj in geom_element:
        if not isinstance(top_obj, DB.GeometryInstance):
            continue

        dwg_transform = top_obj.Transform
        output.print_md("### Top-level DWG wrapper GI")

        # Check top-level GI's GraphicsStyle
        top_gs = get_gs_full_info(top_obj)
        output.print_md("Top GS: name='{}' subcat='{}' cat='{}'".format(
            top_gs["name"], top_gs["subcategory"], top_gs["category"]
        ))

        try:
            symbol_geom = top_obj.GetSymbolGeometry()
        except Exception as e:
            output.print_md("GetSymbolGeometry error: {}".format(str(e)))
            continue

        if not symbol_geom:
            continue

        for child_obj in symbol_geom:
            if not isinstance(child_obj, DB.GeometryInstance):
                continue

            block_count += 1
            if block_count > MAX_BLOCKS:
                output.print_md(
                    "... ({} blocks shown, stopping for readability)".format(MAX_BLOCKS)
                )
                break

            output.print_md("---")
            output.print_md("### Block #{} (GI Id: {})".format(
                block_count, child_obj.Id
            ))

            lines = explore_block_gi(child_obj)
            for line in lines:
                output.print_md(line)

    output.print_md("")
    output.print_md("---")
    output.print_md("**Total blocks analysed: {}**".format(
        min(block_count, MAX_BLOCKS)))
    output.print_md("Diagnostic complete.")


main()
