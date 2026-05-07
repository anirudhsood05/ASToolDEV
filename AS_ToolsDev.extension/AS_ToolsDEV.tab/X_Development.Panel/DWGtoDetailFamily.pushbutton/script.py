# -*- coding: utf-8 -*-
__title__ = 'DWG to\nDetail Family'
__doc__ = """Select DWG files and a family template to batch-convert DWGs into
Revit detail item families (.rfa), saved to a chosen output folder."""
__author__ = "AS Tools"

import os
import clr

clr.AddReference('System.Windows.Forms')

from System.Windows.Forms import OpenFileDialog, FolderBrowserDialog, DialogResult
from Autodesk.Revit.DB import (
    Transaction, SaveAsOptions, DWGImportOptions, ImportPlacement,
    FilteredElementCollector, BuiltInCategory, ViewPlan, ViewType,
    IFailuresPreprocessor, FailureProcessingResult, FailureSeverity,
)
from pyrevit import revit, forms, script

app    = revit.doc.Application
output = script.get_output()


# ---------------------------------------------------------------------------
# Failure handler – swallows DWG import warnings silently
# ---------------------------------------------------------------------------
class WarningSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        for msg in list(failuresAccessor.GetFailureMessages()):
            if msg.GetSeverity() == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(msg)
        return FailureProcessingResult.Continue


# ---------------------------------------------------------------------------
# File / folder pickers
# ---------------------------------------------------------------------------
def pick_file(title, file_filter, multi=False):
    dlg = OpenFileDialog()
    dlg.Title      = title
    dlg.Filter     = file_filter
    dlg.Multiselect = multi
    if dlg.ShowDialog() == DialogResult.OK:
        return list(dlg.FileNames) if multi else dlg.FileName
    return None


def pick_folder(description):
    dlg = FolderBrowserDialog()
    dlg.Description = description
    if dlg.ShowDialog() == DialogResult.OK:
        return dlg.SelectedPath
    return None


# ---------------------------------------------------------------------------
# Revit helpers
# ---------------------------------------------------------------------------
def get_ref_level_floor_plan(fam_doc):
    """Return the 'Ref. Level' floor plan view from a family document."""
    collector = FilteredElementCollector(fam_doc).OfCategory(BuiltInCategory.OST_Views)
    for view in collector:
        if (isinstance(view, ViewPlan)
                and not view.IsTemplate
                and view.ViewType == ViewType.FloorPlan
                and view.Name == "Ref. Level"):
            return view
    return None


# ---------------------------------------------------------------------------
# Core conversion
# ---------------------------------------------------------------------------
def create_family_from_dwg(dwg_path, template_path, save_folder):
    """
    Open a new family from *template_path*, import *dwg_path*, save as .rfa.
    Returns (save_path, None) on success or (None, error_message) on failure.
    """
    # --- open family document -------------------------------------------------
    try:
        fam_doc = app.NewFamilyDocument(template_path)
    except Exception as ex:
        return None, "NewFamilyDocument failed: " + str(ex)

    if fam_doc is None:
        return None, "NewFamilyDocument returned None."

    # --- import DWG inside a transaction -------------------------------------
    t = Transaction(fam_doc, "Import DWG")
    try:
        fail_opts = t.GetFailureHandlingOptions()
        fail_opts.SetFailuresPreprocessor(WarningSwallower())
        t.SetFailureHandlingOptions(fail_opts)
        t.Start()

        floor_plan = get_ref_level_floor_plan(fam_doc)
        if floor_plan is None:
            t.RollBack()
            fam_doc.Close(False)
            return None, "No 'Ref. Level' floor plan found in template."

        dwg_opts           = DWGImportOptions()
        dwg_opts.Placement = ImportPlacement.Origin
        # out parameter handled automatically by IronPython – return value ignored
        fam_doc.Import(dwg_path, dwg_opts, floor_plan)
        t.Commit()

    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        try:
            fam_doc.Close(False)
        except Exception:
            pass
        return None, "Import transaction failed: " + str(ex)

    # --- save as .rfa --------------------------------------------------------
    dwg_filename = os.path.basename(dwg_path)
    rfa_name     = os.path.splitext(dwg_filename)[0] + ".rfa"
    save_path    = os.path.join(save_folder, rfa_name)

    try:
        save_opts = SaveAsOptions()
        save_opts.OverwriteExistingFile = True
        fam_doc.SaveAs(save_path, save_opts)
        fam_doc.Close(False)
        return save_path, None
    except Exception as ex:
        try:
            fam_doc.Close(False)
        except Exception:
            pass
        return None, "Save failed: " + str(ex)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    # 1 – pick template
    template_path = pick_file(
        "Select Family Template (.rft)",
        "Family Templates (*.rft)|*.rft"
    )
    if not template_path:
        forms.alert("No template selected. Operation cancelled.", exitscript=True)
        return
    if not os.path.isfile(template_path):
        forms.alert("Template file not found:\n" + template_path, exitscript=True)
        return

    # 2 – pick DWG files
    dwg_files = pick_file(
        "Select DWG Files to Convert",
        "DWG Files (*.dwg)|*.dwg",
        multi=True
    )
    if not dwg_files:
        forms.alert("No DWG files selected. Operation cancelled.", exitscript=True)
        return

    # 3 – pick output folder
    save_folder = pick_folder("Select Folder to Save RFA Families")
    if not save_folder:
        forms.alert("No save folder selected. Operation cancelled.", exitscript=True)
        return

    # 4 – process each DWG
    created = 0
    failed  = []

    for dwg_path in dwg_files:
        save_path, err = create_family_from_dwg(dwg_path, template_path, save_folder)
        name = os.path.basename(dwg_path)
        if save_path:
            created += 1
            output.print_md("**Created**: {}".format(os.path.basename(save_path)))
        else:
            failed.append("{} ({})".format(name, err or "Unknown error"))
            output.print_md("**Failed**: {} - {}".format(name, err or "Unknown error"))

    # 5 – summary alert
    msg = "Created {} famil{} saved to:\n{}".format(
        created,
        "ies" if created != 1 else "y",
        save_folder
    )
    if failed:
        msg += "\n\nFailed ({}):\n- {}".format(len(failed), "\n- ".join(failed))

    forms.alert(msg)


main()
