# -*- coding: utf-8 -*-
"""CAD Geometry Diagnostic v2 - Safe attribute discovery.

Handles the fact that CAD GeometryInstance objects do NOT have
a .Symbol property (unlike family GeometryInstances).
"""

__title__ = "CAD Geometry\nDiagnostic v2"
__author__ = "Aukett Swanke BIM"

from pyrevit import script, forms, revit, DB
import math

doc = revit.doc
output = script.get_output()

# =============================================================================
# COLLECT IMPORTS
# =============================================================================
collector = DB.FilteredElementCollector(doc)\
    .OfClass(DB.ImportInstance)\
    .WhereElementIsNotElementType()
imports = list(collector)

if not imports:
    forms.alert("No linked or imported CAD files found.", title="No CAD Files")
    script.exit()

# Let user pick
choices = []
for imp in imports:
    try:
        type_elem = doc.GetElement(imp.GetTypeId())
        type_name = type_elem.Name if type_elem else "Unknown"
    except Exception:
        type_name = "Unknown"
    link_status = "Linked" if imp.IsLinked else "Imported"
    choices.append("{} [{}] (ID: {})".format(type_name, link_status, imp.Id.IntegerValue))

selected = forms.SelectFromList.show(
    choices,
    title="Select CAD File to Diagnose",
    button_name="Diagnose"
)
if not selected:
    script.exit()

selected_idx = choices.index(selected)
imp = imports[selected_idx]

# =============================================================================
# REPORT HEADER
# =============================================================================
output.print_md("# CAD Geometry Diagnostic Report v2")
output.print_md("---")

try:
    type_elem = doc.GetElement(imp.GetTypeId())
    type_name = type_elem.Name if type_elem else "Unknown"
except Exception:
    type_name = "Unknown"

output.print_md("**CAD File**: {} (ID: {})".format(type_name, imp.Id.IntegerValue))
output.print_md("**Is Linked**: {}".format(imp.IsLinked))

imp_transform = imp.GetTransform()
output.print_md("**Import Transform Origin**: ({}, {}, {})".format(
    round(imp_transform.Origin.X, 4),
    round(imp_transform.Origin.Y, 4),
    round(imp_transform.Origin.Z, 4)
))
imp_total_transform = imp.GetTotalTransform()
output.print_md("**Import TotalTransform Origin**: ({}, {}, {})".format(
    round(imp_total_transform.Origin.X, 4),
    round(imp_total_transform.Origin.Y, 4),
    round(imp_total_transform.Origin.Z, 4)
))


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================
def get_graphics_style_name(geom_obj):
    """Get GraphicsStyle name (= CAD layer name)."""
    try:
        gs_id = geom_obj.GraphicsStyleId
        if gs_id and gs_id != DB.ElementId.InvalidElementId:
            gs = doc.GetElement(gs_id)
            if gs:
                return gs.Name
    except Exception:
        pass
    return "<no style>"


def describe_transform(transform):
    """Describe a transform's origin and rotation."""
    if not transform:
        return "None"
    try:
        bx = transform.BasisX
        angle_deg = round(math.degrees(math.atan2(bx.Y, bx.X)), 2)
        return "Origin=({},{},{}) Rot={}deg".format(
            round(transform.Origin.X, 3),
            round(transform.Origin.Y, 3),
            round(transform.Origin.Z, 3),
            angle_deg
        )
    except Exception:
        return "Transform (could not read)"


def safe_get_gi_name(geom_instance):
    """Safely extract any identifying name from a GeometryInstance.

    CAD-sourced GeometryInstance objects do NOT have .Symbol.
    We try multiple approaches.
    """
    # Approach 1: Try .Symbol (works for Revit families, NOT for CAD)
    try:
        sym = geom_instance.Symbol
        if sym:
            return "Symbol: '{}'".format(getattr(sym, 'Name', str(sym)))
    except (AttributeError, Exception):
        pass

    # Approach 2: GraphicsStyle (= CAD layer name)
    gs_name = get_graphics_style_name(geom_instance)

    # Approach 3: Try to get any name-like attributes
    name_attrs = []
    for attr in ['Name', 'UniqueId', 'Id']:
        try:
            val = getattr(geom_instance, attr, None)
            if val is not None:
                name_attrs.append("{}={}".format(attr, val))
        except Exception:
            pass

    if name_attrs:
        return "Layer: '{}' | {}".format(gs_name, ", ".join(name_attrs))
    return "Layer: '{}'".format(gs_name)


def discover_attributes(obj, label=""):
    """List all accessible attributes on an object for discovery."""
    attrs = []
    for attr_name in dir(obj):
        if attr_name.startswith('_'):
            continue
        try:
            val = getattr(obj, attr_name)
            if not callable(val):
                attrs.append((attr_name, str(val)[:80]))
        except Exception:
            attrs.append((attr_name, "<access error>"))
    return attrs


# =============================================================================
# TRAVERSE WITH ALL OPTIONS
# =============================================================================
option_configs = [
    ("Default Options", False, False),
    ("IncludeNonVisibleObjects=True", False, True),
]

for config_name, compute_refs, include_nonvis in option_configs:
    output.print_md("---")
    output.print_md("## Geometry Options: {}".format(config_name))

    options = DB.Options()
    options.ComputeReferences = compute_refs
    options.IncludeNonVisibleObjects = include_nonvis

    try:
        geom_element = imp.get_Geometry(options)
    except Exception as e:
        output.print_md("**ERROR**: get_Geometry failed: {}".format(str(e)))
        continue

    if not geom_element:
        output.print_md("**Result**: get_Geometry returned None")
        continue

    top_level = list(geom_element)
    output.print_md("**Top-level geometry objects**: {}".format(len(top_level)))

    # Type tally
    type_counts = {}
    for obj in top_level:
        t = type(obj).__name__
        type_counts[t] = type_counts.get(t, 0) + 1
    for t, c in sorted(type_counts.items()):
        output.print_md("- {}: {}".format(t, c))

    # =====================================================================
    # DETAILED TRAVERSAL
    # =====================================================================
    output.print_md("### Detailed Hierarchy")

    def dump_geom(geom_elem, depth, max_depth, parent_label):
        if depth > max_depth:
            return 0

        indent = "  " * depth
        obj_count = 0

        for i, geom_obj in enumerate(geom_elem):
            obj_count += 1
            obj_type = type(geom_obj).__name__
            gs_name = get_graphics_style_name(geom_obj)
            label = "{}{}{}".format(parent_label, "." if parent_label else "", i)

            if isinstance(geom_obj, DB.GeometryInstance):
                gi_name = safe_get_gi_name(geom_obj)
                t_desc = describe_transform(geom_obj.Transform)

                output.print_md("{}**[{}] GeometryInstance** | {} | {}".format(
                    indent, label, gi_name, t_desc
                ))

                # Dump ALL non-private attributes of this GeometryInstance
                if depth == 0:
                    output.print_md("{}  **Attribute discovery on top-level GeometryInstance:**".format(indent))
                    attrs = discover_attributes(geom_obj)
                    for attr_name, attr_val in attrs:
                        output.print_md("{}    .{} = {}".format(indent, attr_name, attr_val))

                # Explore GetInstanceGeometry
                try:
                    inst_geom = geom_obj.GetInstanceGeometry()
                    if inst_geom:
                        inst_list = list(inst_geom)
                        inst_types = {}
                        for sub in inst_list:
                            st = type(sub).__name__
                            inst_types[st] = inst_types.get(st, 0) + 1
                        type_summary = ", ".join("{}:{}".format(k, v)
                                                 for k, v in sorted(inst_types.items()))
                        output.print_md("{}  > GetInstanceGeometry: {} objects [{}]".format(
                            indent, len(inst_list), type_summary
                        ))

                        # Recurse
                        dump_geom(inst_geom, depth + 1, max_depth, label + "I")
                except Exception as e:
                    output.print_md("{}  > GetInstanceGeometry ERROR: {}".format(indent, str(e)))

                # Explore GetSymbolGeometry
                try:
                    sym_geom = geom_obj.GetSymbolGeometry()
                    if sym_geom:
                        sym_list = list(sym_geom)
                        sym_types = {}
                        for sub in sym_list:
                            st = type(sub).__name__
                            sym_types[st] = sym_types.get(st, 0) + 1
                        type_summary = ", ".join("{}:{}".format(k, v)
                                                 for k, v in sorted(sym_types.items()))
                        output.print_md("{}  > GetSymbolGeometry: {} objects [{}]".format(
                            indent, len(sym_list), type_summary
                        ))

                        # Recurse into SymbolGeometry too (may differ from InstanceGeometry)
                        dump_geom(sym_geom, depth + 1, max_depth, label + "S")
                except Exception as e:
                    output.print_md("{}  > GetSymbolGeometry ERROR: {}".format(indent, str(e)))

            elif isinstance(geom_obj, DB.Solid):
                face_count = geom_obj.Faces.Size if geom_obj.Faces else 0
                edge_count = geom_obj.Edges.Size if geom_obj.Edges else 0
                output.print_md("{}[{}] Solid | Faces:{} Edges:{} | Layer: '{}'".format(
                    indent, label, face_count, edge_count, gs_name
                ))

            elif isinstance(geom_obj, DB.Line):
                try:
                    start = geom_obj.GetEndPoint(0)
                    end = geom_obj.GetEndPoint(1)
                    output.print_md("{}[{}] Line | ({},{}) to ({},{}) | Layer: '{}'".format(
                        indent, label,
                        round(start.X, 3), round(start.Y, 3),
                        round(end.X, 3), round(end.Y, 3),
                        gs_name
                    ))
                except Exception:
                    output.print_md("{}[{}] Line | Layer: '{}'".format(indent, label, gs_name))

            elif isinstance(geom_obj, DB.Arc):
                try:
                    center = geom_obj.Center
                    radius = round(geom_obj.Radius, 4)
                    output.print_md("{}[{}] Arc | Center:({},{}) R:{} | Layer: '{}'".format(
                        indent, label,
                        round(center.X, 3), round(center.Y, 3),
                        radius, gs_name
                    ))
                except Exception:
                    output.print_md("{}[{}] Arc | Layer: '{}'".format(indent, label, gs_name))

            elif isinstance(geom_obj, DB.PolyLine):
                pts = geom_obj.GetCoordinates()
                output.print_md("{}[{}] PolyLine | Points:{} | Layer: '{}'".format(
                    indent, label, len(pts), gs_name
                ))

            elif isinstance(geom_obj, DB.Point):
                try:
                    coord = geom_obj.Coord
                    output.print_md("{}[{}] Point | ({},{},{}) | Layer: '{}'".format(
                        indent, label,
                        round(coord.X, 3), round(coord.Y, 3), round(coord.Z, 3),
                        gs_name
                    ))
                except Exception:
                    output.print_md("{}[{}] Point | Layer: '{}'".format(indent, label, gs_name))

            elif isinstance(geom_obj, DB.Mesh):
                output.print_md("{}[{}] Mesh | Triangles:{} | Layer: '{}'".format(
                    indent, label, geom_obj.NumTriangles, gs_name
                ))

            else:
                output.print_md("{}[{}] {} | Layer: '{}'".format(
                    indent, label, obj_type, gs_name
                ))

            # Truncation safety
            if obj_count > 500 and depth == 0:
                output.print_md("{}... truncated at 500 top-level objects".format(indent))
                break
            if obj_count > 100 and depth > 0:
                output.print_md("{}... truncated at 100 objects at depth {}".format(indent, depth))
                break

        return obj_count

    total = dump_geom(geom_element, depth=0, max_depth=3, parent_label="")
    output.print_md("**Total objects traversed at top level**: {}".format(total))

# =============================================================================
# LAYER NAMES SUMMARY
# =============================================================================
output.print_md("---")
output.print_md("## All CAD Layer Names (GraphicsStyles)")

options = DB.Options()
options.IncludeNonVisibleObjects = True
geom_element = imp.get_Geometry(options)

layer_names = set()

def collect_layers(geom_elem, depth):
    if depth > 4:
        return
    for geom_obj in geom_elem:
        try:
            gs_id = geom_obj.GraphicsStyleId
            if gs_id and gs_id != DB.ElementId.InvalidElementId:
                gs = doc.GetElement(gs_id)
                if gs:
                    layer_names.add(gs.Name)
        except Exception:
            pass
        if isinstance(geom_obj, DB.GeometryInstance):
            try:
                inst_geom = geom_obj.GetInstanceGeometry()
                if inst_geom:
                    collect_layers(inst_geom, depth + 1)
            except Exception:
                pass

collect_layers(geom_element, 0)

for name in sorted(layer_names):
    output.print_md("- `{}`".format(name))

output.print_md("---")
output.print_md("**Diagnostic complete.** Share this output for troubleshooting.")
