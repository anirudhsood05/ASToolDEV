# -*- coding: utf-8 -*-
__title__ = 'DWG\nConverter'
__doc__ = """Guided DWG-to-native-lines converter.

Project mode (two steps):
  Step 1 — select a family template, DWG file(s), and output folder.
  Step 2 — choose line type and linestyle assignment.
  The tool creates each .rfa family, converts the DWG geometry to native
  Revit lines, removes the original CAD import, and saves a clean family.

Family Editor mode:
  Detects the imported DWG in the open family, then guides you through
  choosing line type and linestyle to convert it in place."""
__author__ = "AS Tools"

import os
import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows.Forms import OpenFileDialog, FolderBrowserDialog, DialogResult
import System.Windows.Controls as WpfControls
import System.Windows as Wpf

from Autodesk.Revit.DB import (
    Transaction, SaveAsOptions, DWGImportOptions, ImportPlacement,
    FilteredElementCollector, BuiltInCategory, ViewPlan, ViewType,
    IFailuresPreprocessor, FailureProcessingResult, FailureSeverity,
    Options, GeometryInstance, PolyLine, Curve, Line,
    GraphicsStyleType, ViewSchedule, ImportInstance, CurveArray,
    SketchPlane, Plane, XYZ,
)
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, forms, script

app    = __revit__.Application
uidoc  = revit.uidoc
doc    = revit.doc
output = script.get_output()


# ---------------------------------------------------------------------------
# SHARED UTILITIES
# ---------------------------------------------------------------------------

class WarningSwallower(IFailuresPreprocessor):
    """Suppresses DWG import warnings."""
    def PreprocessFailures(self, failuresAccessor):
        for msg in list(failuresAccessor.GetFailureMessages()):
            if msg.GetSeverity() == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(msg)
        return FailureProcessingResult.Continue


def pick_file(title, file_filter, multi=False):
    dlg = OpenFileDialog()
    dlg.Title       = title
    dlg.Filter      = file_filter
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


# ── Linestyle helpers ────────────────────────────────────────────────────────

def get_all_linestyles(document):
    """Return list of Category from OST_Lines SubCategories. No transaction needed."""
    try:
        cat = document.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
        return list(cat.SubCategories)
    except Exception:
        return []


def get_linestyle_by_name(document, name):
    """Return GraphicsStyle for the named linestyle, or None."""
    for cat in get_all_linestyles(document):
        if cat.Name == name:
            return cat.GetGraphicsStyle(GraphicsStyleType.Projection)
    return None


# ── View helpers ─────────────────────────────────────────────────────────────

def get_ref_level_floor_plan(fam_doc):
    """Return the 'Ref. Level' floor plan view from a family document."""
    for view in FilteredElementCollector(fam_doc).OfCategory(BuiltInCategory.OST_Views):
        if (isinstance(view, ViewPlan)
                and not view.IsTemplate
                and view.ViewType == ViewType.FloorPlan
                and view.Name == "Ref. Level"):
            return view
    return None


# ── Geometry extraction ──────────────────────────────────────────────────────

def get_layer_name(geom_obj, document):
    try:
        gs = document.GetElement(geom_obj.GraphicsStyleId)
        if gs is not None:
            return gs.GraphicsStyleCategory.Name
    except Exception:
        pass
    return "<Unknown>"


def _iter_geom_curves(import_instance, document):
    """Yield (layer_name, Curve) pairs; PolyLines split into segments."""
    opts = Options()
    opts.ComputeReferences          = False
    opts.IncludeNonVisibleObjects   = False
    geom_elem = import_instance.get_Geometry(opts)
    if geom_elem is None:
        return
    for top_obj in geom_elem:
        if not isinstance(top_obj, GeometryInstance):
            continue
        for child in top_obj.GetInstanceGeometry():
            layer = get_layer_name(child, document)
            if isinstance(child, PolyLine):
                pts = child.GetCoordinates()
                for i in range(len(pts) - 1):
                    try:
                        yield layer, Line.CreateBound(pts[i], pts[i + 1])
                    except Exception:
                        pass  # degenerate segment — skip
            elif isinstance(child, Curve):
                try:
                    if child.Length > 1e-9:
                        yield layer, child
                except Exception:
                    pass


def extract_curves_by_layer(import_instance, document):
    """Return {layer_name: [Curve]} from an ImportInstance."""
    result = {}
    for layer, curve in _iter_geom_curves(import_instance, document):
        result.setdefault(layer, []).append(curve)
    return result


def extract_all_curves(import_instance, document):
    """Return flat list of all Curves, ignoring layers."""
    return [c for _, c in _iter_geom_curves(import_instance, document)]


# ── Line creation ─────────────────────────────────────────────────────────────

def create_line(document, view, curve, line_type):
    """
    Create one native line from a Curve. line_type: 'detail' | 'model'.
    Called inside an open Transaction.
    """
    is_family = document.IsFamilyDocument
    if line_type == 'detail':
        if is_family:
            return document.FamilyCreate.NewDetailCurve(view, curve)
        return document.Create.NewDetailCurve(view, curve)
    # model
    sp = view.SketchPlane
    if sp is None:
        # Family views may not have a sketch plane set — create a fallback
        plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero)
        sp    = SketchPlane.Create(document, plane)
    if is_family:
        return document.FamilyCreate.NewModelCurve(curve, sp)
    return document.Create.NewModelCurve(curve, sp)


def apply_linestyle(element, document, linestyle_name):
    """Assign a linestyle to a line element. Silent no-op if not found."""
    if element is None or not linestyle_name:
        return
    gs = get_linestyle_by_name(document, linestyle_name)
    if gs is None:
        return
    try:
        element.LineStyle = gs
    except Exception:
        pass


# ── Conversion core (runs inside an open Transaction) ────────────────────────

def convert_import_to_lines(document, import_instance, view, line_type, style_config):
    """
    Convert ImportInstance geometry to native lines.
    style_config: {'mode': 'simple', 'style': name}
               or {'mode': 'by_layer', 'mapping': {layer: style}}
    Returns (created_count, [failure_strings]).
    """
    created  = 0
    failures = []
    mode     = style_config.get('mode', 'simple')

    if mode == 'simple':
        style = style_config.get('style', '')
        for curve in extract_all_curves(import_instance, document):
            try:
                elem = create_line(document, view, curve, line_type)
                apply_linestyle(elem, document, style)
                created += 1
            except Exception as ex:
                failures.append(str(ex))
    else:
        mapping        = style_config.get('mapping', {})
        layer_curve_map = extract_curves_by_layer(import_instance, document)
        for layer, curves in layer_curve_map.items():
            style = mapping.get(layer, '')
            if not style:
                output.print_md("  Skipped layer (no style assigned): **{}**".format(layer))
                continue
            for curve in curves:
                try:
                    elem = create_line(document, view, curve, line_type)
                    apply_linestyle(elem, document, style)
                    created += 1
                except Exception as ex:
                    failures.append("Layer '{}': {}".format(layer, ex))

    return created, failures


# ── WPF By-Layer mapping dialog ───────────────────────────────────────────────

BY_LAYER_XAML = u"""<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="Map DWG Layers to Linestyles"
    Width="540" SizeToContent="Height"
    MinHeight="150" MaxHeight="680"
    ResizeMode="CanResize"
    WindowStartupLocation="CenterScreen"
    ShowInTaskbar="False"
    FontFamily="Arial" FontSize="12" Background="#F5F5F5">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0"
               Text="Assign a linestyle to each DWG layer (leave blank to skip):"
               FontWeight="Bold" Margin="0,0,0,8"/>
    <ScrollViewer Grid.Row="1" MaxHeight="500"
                  VerticalScrollBarVisibility="Auto">
      <StackPanel x:Name="layerStack" Margin="0,0,4,0"/>
    </ScrollViewer>
    <Separator Grid.Row="2" Margin="0,10"/>
    <StackPanel Grid.Row="3" Orientation="Horizontal"
                HorizontalAlignment="Right">
      <Button x:Name="btnCancel" Content="Cancel"
              Width="80" Height="26" Margin="0,0,8,0"/>
      <Button x:Name="btnOk" Content="OK"
              Width="80" Height="26"
              Background="#143D5C" Foreground="White" FontWeight="Bold"/>
    </StackPanel>
  </Grid>
</Window>"""


class ByLayerDialog(forms.WPFWindow):
    """Dialog that maps each DWG layer to a linestyle via a ComboBox per row."""

    def __init__(self, layer_names, linestyle_names):
        forms.WPFWindow.__init__(self, BY_LAYER_XAML, literal_string=True)
        self._ok_clicked = False
        self._combos     = {}
        all_choices      = [''] + list(linestyle_names)

        for layer in layer_names:
            row  = WpfControls.Grid()
            col0 = WpfControls.ColumnDefinition()
            col0.Width = Wpf.GridLength(200)
            col1 = WpfControls.ColumnDefinition()
            col1.Width = Wpf.GridLength(1, Wpf.GridUnitType.Star)
            row.ColumnDefinitions.Add(col0)
            row.ColumnDefinitions.Add(col1)
            row.Margin = Wpf.Thickness(0, 3, 0, 3)

            lbl = WpfControls.TextBlock()
            lbl.Text              = layer
            lbl.VerticalAlignment = Wpf.VerticalAlignment.Center
            lbl.TextTrimming      = Wpf.TextTrimming.CharacterEllipsis
            lbl.ToolTip           = layer
            WpfControls.Grid.SetColumn(lbl, 0)
            row.Children.Add(lbl)

            combo = WpfControls.ComboBox()
            for choice in all_choices:
                combo.Items.Add(choice)
            combo.SelectedIndex = 0
            combo.Height        = 22
            combo.Margin        = Wpf.Thickness(8, 0, 0, 0)
            WpfControls.Grid.SetColumn(combo, 1)
            row.Children.Add(combo)

            self.layerStack.Children.Add(row)
            self._combos[layer] = combo

        self.btnOk.Click     += self.on_ok
        self.btnCancel.Click += self.on_cancel

    def on_ok(self, sender, args):
        self._ok_clicked = True
        self.Close()

    def on_cancel(self, sender, args):
        self.Close()

    def get_mapping(self):
        """Return {layer: linestyle} for assigned layers, or {} if cancelled."""
        if not self._ok_clicked:
            return {}
        return {layer: combo.SelectedItem
                for layer, combo in self._combos.items()
                if combo.SelectedItem}


# ── Style config helpers ─────────────────────────────────────────────────────

def ask_style_config(document, import_instance, linestyle_names, allow_by_layer,
                     step_prefix=''):
    """
    Prompt for linestyle assignment mode and return a style_config dict, or None if cancelled.
    If allow_by_layer=True and user picks By Layer, the By Layer dialog is shown immediately
    (requires import_instance to be available for layer extraction).
    """
    if allow_by_layer:
        assign_mode = forms.ask_for_one_item(
            ['Simple (one style for all layers)',
             'By Layer (assign per DWG layer)'],
            prompt=step_prefix + "How should linestyles be assigned?",
            title="Linestyle Assignment"
        )
        if not assign_mode:
            return None
    else:
        assign_mode = 'Simple (one style for all layers)'

    if assign_mode.startswith('Simple'):
        chosen = forms.SelectFromList.show(
            linestyle_names,
            title=step_prefix + "Select Linestyle",
            prompt="Choose a linestyle for all converted lines:",
            multiselect=False
        )
        if not chosen:
            return None
        return {'mode': 'simple', 'style': chosen}

    # By Layer — needs layers from the ImportInstance
    layer_map = extract_curves_by_layer(import_instance, document)
    if not layer_map:
        forms.alert("No geometry layers found in the selected DWG.")
        return None
    dlg     = ByLayerDialog(sorted(layer_map.keys()), linestyle_names)
    dlg.show_dialog()
    mapping = dlg.get_mapping()
    if not mapping:
        return None
    return {'mode': 'by_layer', 'mapping': mapping}


# ---------------------------------------------------------------------------
# CORE PIPELINE — single DWG processed inside its own family document
# ---------------------------------------------------------------------------

def process_dwg_to_family(dwg_path, template_path, save_folder,
                           line_type, style_config, linestyle_names=None):
    """
    Full pipeline for one DWG:
      1. Open family from template
      2. Import DWG (Transaction 1)
      3. Resolve by-layer style config if mapping was deferred (shows dialog now)
      4. Convert to native lines + delete import (Transaction 2)
      5. Save .rfa and close

    style_config may have mapping=None for by-layer when the dialog should be
    shown after the import (layers are unknown before import).
    linestyle_names is required in that case.

    Returns (save_path, None) on success or (None, error_string) on failure.
    """
    # Open family doc
    try:
        fam_doc = app.NewFamilyDocument(template_path)
    except Exception as ex:
        return None, "Could not open template: " + str(ex)
    if fam_doc is None:
        return None, "NewFamilyDocument returned None."

    floor_plan = get_ref_level_floor_plan(fam_doc)
    if floor_plan is None:
        fam_doc.Close(False)
        return None, "No 'Ref. Level' floor plan found in template."

    # Transaction 1: Import DWG
    import_id = None
    t1 = Transaction(fam_doc, "Import DWG")
    try:
        fail_opts = t1.GetFailureHandlingOptions()
        fail_opts.SetFailuresPreprocessor(WarningSwallower())
        t1.SetFailureHandlingOptions(fail_opts)
        t1.Start()
        dwg_opts           = DWGImportOptions()
        dwg_opts.Placement = ImportPlacement.Origin
        import_result      = fam_doc.Import(dwg_path, dwg_opts, floor_plan)
        # IronPython returns (bool, ElementId) for out-param methods
        import_id = import_result[1] if isinstance(import_result, tuple) else None
        t1.Commit()
    except Exception as ex:
        try:    t1.RollBack()
        except Exception: pass
        try:    fam_doc.Close(False)
        except Exception: pass
        return None, "DWG import failed: " + str(ex)

    import_instance = fam_doc.GetElement(import_id) if import_id else None
    if import_instance is None:
        fam_doc.Close(False)
        return None, "ImportInstance not found after import."

    # Resolve deferred by-layer config (dialog shown after import so layers are known)
    if style_config.get('mode') == 'by_layer' and style_config.get('mapping') is None:
        layer_map = extract_curves_by_layer(import_instance, fam_doc)
        if not layer_map:
            fam_doc.Close(False)
            return None, "No geometry layers found in DWG."
        dlg     = ByLayerDialog(sorted(layer_map.keys()), linestyle_names or [])
        dlg.show_dialog()
        mapping = dlg.get_mapping()
        if not mapping:
            fam_doc.Close(False)
            return None, "Cancelled by user."
        effective_config = {'mode': 'by_layer', 'mapping': mapping}
    else:
        effective_config = style_config

    # Transaction 2: Convert geometry to lines + delete import
    t2 = Transaction(fam_doc, "Convert DWG to Lines")
    try:
        t2.Start()
        created, failures = convert_import_to_lines(
            fam_doc, import_instance, floor_plan, line_type, effective_config)
        try:
            fam_doc.Delete(import_id)
        except Exception:
            pass
        t2.Commit()
    except Exception as ex:
        try:    t2.RollBack()
        except Exception: pass
        try:    fam_doc.Close(False)
        except Exception: pass
        return None, "Conversion failed: " + str(ex)

    for fail in failures:
        output.print_md("  **Warning**: {}".format(fail))

    # Save .rfa
    rfa_name  = os.path.splitext(os.path.basename(dwg_path))[0] + ".rfa"
    save_path = os.path.join(save_folder, rfa_name)
    try:
        save_opts = SaveAsOptions()
        save_opts.OverwriteExistingFile = True
        fam_doc.SaveAs(save_path, save_opts)
        fam_doc.Close(False)
    except Exception as ex:
        try: fam_doc.Close(False)
        except Exception: pass
        return None, "Save failed: " + str(ex)

    return save_path, None


# ---------------------------------------------------------------------------
# WORKFLOW A — Project document: guided two-step creation + conversion
# ---------------------------------------------------------------------------

def run_project_workflow():
    output.print_md("## DWG Converter — Project Mode")

    # ── STEP 1: Files ────────────────────────────────────────────────────────
    output.print_md("### Step 1: Select Files and Output Folder")

    template_path = pick_file(
        "Step 1 of 3 — Select Family Template (.rft)",
        "Family Templates (*.rft)|*.rft"
    )
    if not template_path:
        forms.alert("No template selected. Operation cancelled.", exitscript=True)
        return
    if not os.path.isfile(template_path):
        forms.alert("Template file not found:\n" + template_path, exitscript=True)
        return

    dwg_files = pick_file(
        "Step 2 of 3 — Select DWG File(s) to Convert",
        "DWG Files (*.dwg)|*.dwg",
        multi=True
    )
    if not dwg_files:
        forms.alert("No DWG files selected. Operation cancelled.", exitscript=True)
        return

    save_folder = pick_folder("Step 3 of 3 — Select Output Folder for RFA Families")
    if not save_folder:
        forms.alert("No output folder selected. Operation cancelled.", exitscript=True)
        return

    output.print_md("Template : `{}`".format(os.path.basename(template_path)))
    output.print_md("DWG count: **{}**".format(len(dwg_files)))
    output.print_md("Output   : `{}`".format(save_folder))

    # ── STEP 2: Line conversion options ──────────────────────────────────────
    output.print_md("### Step 2: Configure Line Conversion")

    line_type_label = forms.ask_for_one_item(
        ['Detail Line', 'Model Line'],
        prompt="Step 2a: What type of Revit lines should the DWG geometry become?",
        title="Line Type"
    )
    if not line_type_label:
        script.exit()
        return
    line_type = 'detail' if line_type_label == 'Detail Line' else 'model'

    linestyle_names = sorted([c.Name for c in get_all_linestyles(doc)])
    if not linestyle_names:
        forms.alert("No linestyles found in the current document.", exitscript=True)
        return

    is_batch = len(dwg_files) > 1

    if is_batch:
        # Batch: simple mode only — one dialog for all DWGs
        output.print_md("*Batch mode: one linestyle applied to all {} files.*".format(
            len(dwg_files)))
        chosen = forms.SelectFromList.show(
            linestyle_names,
            title="Step 2b: Select Linestyle (applied to all DWGs)",
            prompt="Choose a linestyle for all converted lines:",
            multiselect=False
        )
        if not chosen:
            script.exit()
            return
        style_config   = {'mode': 'simple', 'style': chosen}
        linestyle_names_for_defer = None
    else:
        # Single DWG: allow by-layer; defer the dialog until after import
        assign_mode = forms.ask_for_one_item(
            ['Simple (one style for all layers)',
             'By Layer (assign per DWG layer)'],
            prompt="Step 2b: How should linestyles be assigned?",
            title="Linestyle Assignment"
        )
        if not assign_mode:
            script.exit()
            return

        if assign_mode.startswith('Simple'):
            chosen = forms.SelectFromList.show(
                linestyle_names,
                title="Step 2c: Select Linestyle",
                prompt="Choose a linestyle for all converted lines:",
                multiselect=False
            )
            if not chosen:
                script.exit()
                return
            style_config              = {'mode': 'simple', 'style': chosen}
            linestyle_names_for_defer = None
        else:
            # Layers are unknown until after the DWG is imported — defer the dialog
            style_config              = {'mode': 'by_layer', 'mapping': None}
            linestyle_names_for_defer = linestyle_names

    # ── STEP 3: Process ───────────────────────────────────────────────────────
    output.print_md("### Step 3: Processing")

    created_count = 0
    failed_list   = []

    for dwg_path in dwg_files:
        dwg_name = os.path.basename(dwg_path)
        output.print_md("Processing: **{}**".format(dwg_name))

        save_path, err = process_dwg_to_family(
            dwg_path, template_path, save_folder,
            line_type, style_config, linestyle_names_for_defer
        )
        if save_path:
            created_count += 1
            output.print_md("  Saved: `{}`".format(os.path.basename(save_path)))
        else:
            failed_list.append("{} — {}".format(dwg_name, err or "Unknown error"))
            output.print_md("  **Failed**: {}".format(err or "Unknown error"))

    # ── Summary ───────────────────────────────────────────────────────────────
    msg = "Converted {} DWG{} to native {} lines.\nFamilies saved to:\n{}".format(
        created_count,
        "s" if created_count != 1 else "",
        line_type_label,
        save_folder
    )
    if failed_list:
        msg += "\n\nFailed ({}):\n- {}".format(len(failed_list), "\n- ".join(failed_list))
    forms.alert(msg)


# ---------------------------------------------------------------------------
# WORKFLOW B — Family Editor: convert existing ImportInstance in open family
# ---------------------------------------------------------------------------

def run_family_editor_workflow():
    output.print_md("## DWG Converter — Family Editor Mode")

    # Find all ImportInstances in the open family
    imports = list(
        FilteredElementCollector(doc)
        .OfClass(ImportInstance)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    if not imports:
        forms.alert(
            "No imported DWG files found in this family.\n\n"
            "Use Insert → Import CAD to import a DWG first, "
            "then run this tool again.",
            exitscript=True
        )
        return

    # Select import — auto if only one, otherwise let user choose
    if len(imports) == 1:
        import_instance = imports[0]
    else:
        names   = [imp.Category.Name if imp.Category else "Unnamed ({})".format(imp.Id)
                   for imp in imports]
        chosen_name = forms.ask_for_one_item(
            names,
            prompt="Multiple imported DWGs found. Select one to convert:",
            title="Select DWG Import"
        )
        if not chosen_name:
            script.exit()
            return
        import_instance = imports[names.index(chosen_name)]

    active_view = doc.ActiveView
    if isinstance(active_view, ViewSchedule):
        forms.alert("Please switch to a floor plan view first.", exitscript=True)
        return

    output.print_md("Selected import: **{}**".format(
        import_instance.Category.Name if import_instance.Category else "Unnamed"))

    # ── Step 1: Line type ─────────────────────────────────────────────────────
    line_type_label = forms.ask_for_one_item(
        ['Detail Line', 'Model Line'],
        prompt="Step 1: What type of Revit lines should the DWG geometry become?",
        title="Line Type"
    )
    if not line_type_label:
        script.exit()
        return
    line_type = 'detail' if line_type_label == 'Detail Line' else 'model'

    # ── Step 2: Linestyle assignment ──────────────────────────────────────────
    linestyle_names = sorted([c.Name for c in get_all_linestyles(doc)])
    if not linestyle_names:
        forms.alert("No linestyles found in this family document.", exitscript=True)
        return

    style_config = ask_style_config(
        doc, import_instance, linestyle_names,
        allow_by_layer=True,
        step_prefix="Step 2: "
    )
    if style_config is None:
        script.exit()
        return

    # ── Convert ───────────────────────────────────────────────────────────────
    output.print_md("### Converting...")

    t = Transaction(doc, "DWG to Native Lines")
    t.Start()
    try:
        created, failures = convert_import_to_lines(
            doc, import_instance, active_view, line_type, style_config)
        t.Commit()
    except Exception as ex:
        try:    t.RollBack()
        except Exception: pass
        forms.alert("Conversion failed: " + str(ex), exitscript=True)
        return

    for fail in failures:
        output.print_md("**Warning**: {}".format(fail))

    output.print_md("Created **{}** {} line{}.".format(
        created, line_type_label, "s" if created != 1 else ""))

    # ── Offer to delete the original import ───────────────────────────────────
    msg = "Created {} {} line{}.\n\nDelete the original DWG import from this family?".format(
        created, line_type_label, "s" if created != 1 else "")
    if failures:
        msg = "{} curve(s) could not be converted (see output panel).\n\n".format(
            len(failures)) + msg

    if forms.alert(msg, yes=True, no=True):
        t_del = Transaction(doc, "Delete DWG Import")
        t_del.Start()
        try:
            doc.Delete(import_instance.Id)
            t_del.Commit()
            output.print_md("Original DWG import deleted.")
        except Exception as ex:
            try:    t_del.RollBack()
            except Exception: pass
            output.print_md("**Could not delete import**: {}".format(str(ex)))


# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------

def main():
    if doc is None:
        forms.alert("No document is open. Please open a project or family first.",
                    exitscript=True)
        return

    if doc.IsFamilyDocument:
        run_family_editor_workflow()
    else:
        if isinstance(doc.ActiveView, ViewSchedule):
            forms.alert(
                "Please switch to a non-schedule view before running this tool.",
                exitscript=True)
            return
        run_project_workflow()


main()
