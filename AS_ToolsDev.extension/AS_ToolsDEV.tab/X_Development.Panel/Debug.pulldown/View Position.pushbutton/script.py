# -*- coding: utf-8 -*-
"""Read viewport centre position from active sheet."""
from pyrevit import revit, DB, script

doc = revit.doc
output = script.get_output()
view = doc.ActiveView

if view.ViewType != DB.ViewType.DrawingSheet:
    output.print_md("**Open a sheet with a placed viewport first**")
else:
    vp_ids = view.GetAllViewports()
    for vp_id in vp_ids:
        vp = doc.GetElement(vp_id)
        center = vp.GetBoxCenter()
        output.print_md("**Viewport:** {}".format(vp.Name))
        output.print_md("- Center X: {} ft ({} mm)".format(
            round(center.X, 6), round(center.X * 304.8, 1)))
        output.print_md("- Center Y: {} ft ({} mm)".format(
            round(center.Y, 6), round(center.Y * 304.8, 1)))