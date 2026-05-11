# -*- coding: utf-8 -*-
# script.py - Purge Families tool (workshared-aware, with modeless cancel window)
# Compatible: Revit 2023-2026, pyRevit v4.5+ / v6, IronPython 2.7

__title__ = 'Purge\nfamilies'
__doc__ = """Reduce model size by purging loaded families recursively.
Supports workshared models - attempts to check out each family before editing.
Purged families are reloaded back into the active project so the project keeps
the same set of families but at a reduced internal size.
A cancel window appears on run - click Cancel to stop gracefully at the next safe point."""
__helpurl__ = "https://apex-project.github.io/pyApex/help#purge-families"

# Required for modeless WPF window to survive garbage collection
__persistentengine__ = True

import os
import sys
import time
import math
import locale
import tempfile
import traceback
from datetime import datetime

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

from System.IO import StringReader
from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.Windows.Threading import DispatcherFrame, Dispatcher, DispatcherPriority
from System.Collections.Generic import List
from System import Action

from pyrevit import script, revit
from pyrevit.revit import doc as ACTIVE_DOC

from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter, Document, ElementId,
    FamilyInstanceFilter, FamilySymbol, FilteredElementCollector,
    IFamilyLoadOptions, SaveAsOptions, StorageType, Transaction,
    TransactWithCentralOptions, SynchronizeWithCentralOptions,
    RelinquishOptions, WorksharingUtils
)

# Optional API surfaces - guard for version differences
try:
    from Autodesk.Revit.DB import CheckoutStatus
    HAS_CHECKOUT_STATUS = True
except ImportError:
    CheckoutStatus = None
    HAS_CHECKOUT_STATUS = False

try:
    from Autodesk.Revit.DB import AnnotationSymbolType
except ImportError:
    AnnotationSymbolType = None

try:
    from Autodesk.Revit.DB import ParameterType
    HAS_PARAMETER_TYPE = True
except ImportError:
    ParameterType = None
    HAS_PARAMETER_TYPE = False


output = script.get_output()
logger = script.get_logger()
my_config = script.get_config()

try:
    locale.setlocale(locale.LC_ALL, '')
except Exception:
    pass

window_title = __title__.replace("\n", " ")
try:
    output.set_title(window_title)
except Exception:
    pass


def dbg(msg):
    try:
        print("[PurgeFamilies] %s" % msg)
    except Exception:
        pass


# -------------------- Runtime state --------------------
IS_WORKSHARED = False
CHECKED_OUT_IDS = []     # Track what we checked out so we can relinquish at end
SYNC_AT_END = False
CANCEL_REQUESTED = False
_CANCEL_WIN = None

PURGE_RESULTS_CSV = []
START_TIME = None
PURGE_RESULTS = {}
PURGE_SIZES_SUM = 0.0
SKIPPED_NOT_OWNED = []   # Families skipped because owned by others
RELOADED_COUNT = 0       # Families successfully reloaded back to host
RELOAD_FAILED = []       # Families that could not be reloaded back


# -------------------- Cancel window (modeless WPF) --------------------
CANCEL_XAML = u"""<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Purge Families - Running"
        Width="360" Height="170"
        Topmost="True"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5"
        ShowInTaskbar="True">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0"
                   Text="Purge Families is running..."
                   FontWeight="Bold" FontSize="13"
                   Foreground="#2C2C2C"
                   Margin="0,0,0,6"/>
        <TextBlock x:Name="statusText" Grid.Row="1"
                   Text="Starting..."
                   TextWrapping="Wrap"
                   Foreground="#2C2C2C"
                   FontSize="11"/>
        <Button x:Name="cancelButton" Grid.Row="2"
                Content="Cancel"
                Width="100" Height="28"
                HorizontalAlignment="Right"
                Background="#E8A735"
                Foreground="White"
                FontWeight="Bold"
                BorderThickness="0"/>
    </Grid>
</Window>"""


def _on_cancel_click(sender, args):
    global CANCEL_REQUESTED
    CANCEL_REQUESTED = True
    try:
        sender.IsEnabled = False
        sender.Content = "Cancelling..."
    except Exception:
        pass
    try:
        if _CANCEL_WIN is not None:
            txt = _CANCEL_WIN.FindName("statusText")
            if txt is not None:
                txt.Text = "Cancelling - finishing current family safely..."
    except Exception:
        pass


def _on_cancel_closing(sender, args):
    global CANCEL_REQUESTED
    CANCEL_REQUESTED = True


def _pump_dispatcher():
    try:
        frame = DispatcherFrame()

        def exit_frame():
            frame.Continue = False

        Dispatcher.CurrentDispatcher.BeginInvoke(
            DispatcherPriority.Background, Action(exit_frame))
        Dispatcher.PushFrame(frame)
    except Exception:
        pass


def show_cancel_window():
    global _CANCEL_WIN
    try:
        win = XamlReader.Parse(CANCEL_XAML)
        btn = win.FindName("cancelButton")
        if btn is not None:
            btn.Click += _on_cancel_click
        win.Closing += _on_cancel_closing
        win.Show()
        _CANCEL_WIN = win
        _pump_dispatcher()
        return True
    except Exception as e:
        logger.debug("Cancel window failed to show: %s" % str(e))
        _CANCEL_WIN = None
        return False


def update_cancel_status(msg):
    if _CANCEL_WIN is None:
        return
    try:
        txt = _CANCEL_WIN.FindName("statusText")
        if txt is not None:
            txt.Text = msg
        _pump_dispatcher()
    except Exception:
        pass


def close_cancel_window():
    global _CANCEL_WIN
    if _CANCEL_WIN is None:
        return
    try:
        try:
            _CANCEL_WIN.Closing -= _on_cancel_closing
        except Exception:
            pass
        _CANCEL_WIN.Close()
    except Exception:
        pass
    _CANCEL_WIN = None


def cancelled():
    _pump_dispatcher()
    return CANCEL_REQUESTED


# -------------------- Compatibility helpers --------------------
def eid_int(eid):
    if eid is None:
        return -1
    # Revit 2024+ uses .Value (Int64); earlier uses .IntegerValue
    try:
        return int(eid.Value)
    except Exception:
        pass
    try:
        return int(eid.IntegerValue)
    except Exception:
        pass
    return -1


def safe_name(element):
    if element is None:
        return "<none>"
    try:
        from Autodesk.Revit.DB import Element as _Element
        return _Element.Name.GetValue(element)
    except Exception:
        pass
    try:
        return element.Name
    except Exception:
        return "<unnamed>"


def is_string(v):
    try:
        return isinstance(v, (str, unicode))
    except NameError:
        return isinstance(v, str)


def doc_is_valid(d):
    if d is None:
        return False
    try:
        return bool(d.IsValidObject)
    except Exception:
        return False


def doc_is_workshared(d):
    try:
        return bool(d.IsWorkshared)
    except Exception:
        return False


# -------------------- Transaction helper --------------------
def run_in_transaction(document, name, func):
    """Run func() inside a transaction. Rolls back on error. Returns (ok, result)."""
    if not doc_is_valid(document):
        return False, None
    t = Transaction(document, name)
    started = False
    try:
        t.Start()
        started = True
        result = func()
        t.Commit()
        return True, result
    except Exception as e:
        logger.debug("Transaction '%s' failed: %s" % (name, str(e)))
        if started:
            try:
                t.RollBack()
            except Exception:
                pass
        return False, None


# -------------------- Workshared helpers --------------------
def get_checkout_status(document, element_id):
    """Returns: 'owned' (me), 'other' (someone else), 'free' (nobody), 'unknown'"""
    if not IS_WORKSHARED:
        return 'owned'
    if not HAS_CHECKOUT_STATUS:
        return 'unknown'
    try:
        status = WorksharingUtils.GetCheckoutStatus(document, element_id)
        if status == CheckoutStatus.OwnedByCurrentUser:
            return 'owned'
        elif status == CheckoutStatus.OwnedByOtherUser:
            return 'other'
        elif status == CheckoutStatus.NotOwned:
            return 'free'
    except Exception as e:
        logger.debug("GetCheckoutStatus error: %s" % str(e))
    return 'unknown'


def get_element_owner(document, element_id):
    if not IS_WORKSHARED:
        return ""
    try:
        info = WorksharingUtils.GetWorksharingTooltipInfo(document, element_id)
        if info is not None:
            return info.Owner or ""
    except Exception:
        pass
    return ""


def try_checkout(document, element_id):
    """Attempt to check out a single element. Returns True if we own it after the call."""
    global CHECKED_OUT_IDS
    if not IS_WORKSHARED:
        return True

    status = get_checkout_status(document, element_id)
    if status == 'owned':
        return True
    if status == 'other':
        return False

    try:
        id_list = List[ElementId]()
        id_list.Add(element_id)
        WorksharingUtils.CheckoutElements(document, id_list)
        # Verify we actually got it (CheckoutElements doesn't always raise on failure)
        if get_checkout_status(document, element_id) == 'owned':
            CHECKED_OUT_IDS.append(eid_int(element_id))
            return True
    except Exception as e:
        logger.debug("Checkout failed for id %d: %s" % (eid_int(element_id), str(e)))
    return False


def relinquish_all_checked_out(document):
    if not IS_WORKSHARED:
        return
    try:
        opts = RelinquishOptions(False)
        opts.CheckedOutElements = True
        opts.FamilyWorksets = False
        opts.UserWorksets = False
        opts.UserCreatedWorksets = False
        try:
            opts.StandardWorksets = False
        except Exception:
            pass
        WorksharingUtils.RelinquishOwnership(document, opts, None)
        dbg("Released checked-out elements")
    except Exception as e:
        logger.debug("Relinquish failed: %s" % str(e))


def sync_with_central(document):
    if not IS_WORKSHARED:
        return True
    try:
        dbg("Syncing with central...")
        update_cancel_status("Syncing with central...")

        twc_opts = TransactWithCentralOptions()
        swc_opts = SynchronizeWithCentralOptions()
        swc_opts.Comment = "Purge Families tool - automated purge"
        swc_opts.SaveLocalBefore = False
        swc_opts.SaveLocalAfter = False

        relinquish = RelinquishOptions(True)
        relinquish.CheckedOutElements = True
        swc_opts.SetRelinquishOptions(relinquish)

        document.SynchronizeWithCentral(twc_opts, swc_opts)
        dbg("Sync with central complete")
        return True
    except Exception as e:
        logger.error("Sync with central failed: %s" % str(e))
        print("WARNING: Sync with central failed - %s" % str(e))
        print("         You will need to sync manually to push purge results to central.")
        return False


# -------------------- Config --------------------
def config_temp_dir():
    default_dir = os.path.join(tempfile.gettempdir(), "pyApex_purgeFamilies")
    v = None
    try:
        v = my_config.temp_dir
    except Exception:
        v = None
    if v is None:
        v = default_dir
        try:
            my_config.temp_dir = v
            script.save_config()
        except Exception:
            pass
    if isinstance(v, list):
        v = v[0] if v else default_dir
    if not is_string(v):
        v = default_dir
    return v


# -------------------- Globals --------------------
BUILTINCATEGORIES_DICT = {}
try:
    for _c in BuiltInCategory.GetValues(BuiltInCategory):
        try:
            BUILTINCATEGORIES_DICT[int(_c)] = _c
        except Exception:
            continue
except Exception:
    pass


# -------------------- Utils --------------------
def invert_dict_of_lists(d):
    result = {}
    if not d:
        return result
    for k in d.keys():
        for e in (d.get(k) or []):
            result.setdefault(e, []).append(k)
    return result


def file_size_mb(f):
    if not f or not is_string(f) or not os.path.exists(f):
        return 0.0
    try:
        return float(os.stat(f).st_size) / (1024.0 * 1024.0)
    except Exception:
        return 0.0


def time_elapsed():
    if START_TIME is None:
        return 0
    return time.time() - START_TIME


def roundup(x):
    try:
        return int(math.ceil(x / 10.0)) * 10
    except Exception:
        return 0


def time_format(t):
    try:
        t = float(t)
    except Exception:
        return "0 sec"
    if t < 60:
        return "%d sec" % t
    elif t < 120:
        return "%d m %d sec" % (t / 60, roundup(t % 60))
    elif t < 6000:
        return "%d min" % (t / 60.0)
    else:
        h = t / 3600.0
        m = (t / 60.0) % 60
        return "%d h %d min" % (h, roundup(m))


def safe_mkdir(path):
    if not path or not is_string(path):
        return False
    try:
        if not os.path.exists(path):
            os.makedirs(path)
        return os.path.isdir(path)
    except Exception as e:
        logger.debug("safe_mkdir failed: %s" % str(e))
        return False


def safe_remove(path, retries=3, delay=0.3):
    if not path or not is_string(path) or not os.path.exists(path):
        return False
    for _ in range(retries):
        try:
            os.remove(path)
            return True
        except Exception:
            time.sleep(delay)
    return False


def sanitise_filename(s):
    if not s:
        return "unnamed"
    for ch in '<>:"/\\|?*\r\n\t':
        s = s.replace(ch, "_")
    return s.strip() or "unnamed"


# -------------------- Dependency analysis --------------------
def dependencies_find(document, element, result=None):
    if result is None:
        result = []
    if element is None:
        return result
    try:
        if not element.IsValidObject:
            return result
    except Exception:
        return result

    try:
        params = element.Parameters
    except Exception:
        params = None

    if params:
        for p in params:
            try:
                if p is None or not p.HasValue:
                    continue
                try:
                    if p.StorageType != StorageType.ElementId:
                        continue
                except Exception:
                    continue
                eid = p.AsElementId()
                e_id_child = eid_int(eid)
                if e_id_child > 0:
                    result.append(e_id_child)
                    try:
                        e_child = document.GetElement(eid)
                    except Exception:
                        e_child = None
                    if e_child is not None:
                        ct = type(e_child)
                        is_fs = (ct == FamilySymbol) or (
                            AnnotationSymbolType is not None and ct == AnnotationSymbolType)
                        if is_fs:
                            try:
                                fam = e_child.Family
                                if fam is not None:
                                    result.append(eid_int(fam.Id))
                            except Exception:
                                pass
            except Exception:
                continue

    try:
        tid = eid_int(element.GetTypeId())
        if tid > 0:
            result.append(tid)
    except Exception:
        pass

    try:
        ls = element.LineStyle
        if ls is not None:
            ls_int = eid_int(ls.Id)
            if ls_int > 0:
                result.append(ls_int)
    except Exception:
        pass

    return result


def _is_family_type_parameter(param):
    if param is None:
        return False
    try:
        if HAS_PARAMETER_TYPE and ParameterType is not None:
            if param.Definition.ParameterType == ParameterType.FamilyType:
                return True
    except Exception:
        pass
    try:
        defn = param.Definition
        get_data_type = getattr(defn, 'GetDataType', None)
        if get_data_type is not None:
            dt = get_data_type()
            try:
                if 'familyType' in str(dt.TypeId).lower():
                    return True
            except Exception:
                pass
    except Exception:
        pass
    return False


def dependencies_structure(document):
    result = {}
    elements = []
    try:
        elements = list(FilteredElementCollector(document)
                        .WhereElementIsNotElementType().ToElements())
    except Exception as e:
        logger.error("Collect instances failed: %s" % str(e))
    try:
        elements += list(FilteredElementCollector(document)
                         .WhereElementIsElementType().ToElements())
    except Exception as e:
        logger.error("Collect types failed: %s" % str(e))

    for e in elements:
        try:
            if e is None or not e.IsValidObject:
                continue
            e_id = eid_int(e.Id)
            if e_id <= 0:
                continue
            childs = dependencies_find(document, e)
            if childs:
                result.setdefault(e_id, []).extend(childs)
        except Exception:
            continue

    inv = invert_dict_of_lists(result)
    inv.pop(-1, None)
    inv_set = set(inv.keys())

    try:
        if document.IsFamilyDocument:
            mgr = document.FamilyManager
            if mgr is not None:
                try:
                    params = mgr.Parameters
                except Exception:
                    params = None
                if params:
                    fam_params = [p for p in params if _is_family_type_parameter(p)]
                    if fam_params:
                        try:
                            doc_types = mgr.Types
                        except Exception:
                            doc_types = []
                        found_params = []
                        categories = []
                        for t in doc_types:
                            for p in fam_params:
                                try:
                                    if p in found_params or not t.HasValue(p):
                                        continue
                                    pft = document.GetElement(t.AsElementId(p))
                                    if pft is not None and pft.Category is not None:
                                        cat = BUILTINCATEGORIES_DICT.get(eid_int(pft.Category.Id))
                                        if cat is not None:
                                            categories.append(cat)
                                        found_params.append(p)
                                except Exception:
                                    continue
                        for c in categories:
                            try:
                                cat_elems = FilteredElementCollector(document).OfCategory(c) \
                                    .WhereElementIsElementType().ToElements()
                                for x in cat_elems:
                                    try:
                                        is_fs = (type(x) == FamilySymbol) or (
                                            AnnotationSymbolType is not None
                                            and type(x) == AnnotationSymbolType)
                                        if is_fs and x.Family is not None:
                                            inv_set.add(eid_int(x.Family.Id))
                                    except Exception:
                                        continue
                            except Exception:
                                continue
    except Exception as e:
        logger.debug("Family param processing error: %s" % str(e))

    return inv_set


# -------------------- Family collection --------------------
class FamilyLoadOption(IFamilyLoadOptions):
    """Always overwrite existing family + parameter values when reloading."""
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        return True


def get_familysymbol_instances(document, fs):
    instances = set()
    try:
        fifilter = FamilyInstanceFilter(document, fs.Id)
        cl = FilteredElementCollector(document)
        for e in cl.WhereElementIsNotElementType().WherePasses(fifilter).ToElements():
            instances.add(e)
    except Exception as e:
        logger.debug("FamilyInstanceFilter error: %s" % str(e))

    try:
        fam = getattr(fs, 'Family', None)
        target_fam_int = eid_int(fam.Id) if fam is not None else -1
        if target_fam_int > 0:
            cl2 = FilteredElementCollector(document)
            legends = cl2.WhereElementIsNotElementType() \
                .OfCategory(BuiltInCategory.OST_LegendComponents).ToElements()
            for e in legends:
                try:
                    p = e.get_Parameter(BuiltInParameter.LEGEND_COMPONENT)
                    if p is None:
                        continue
                    legend_fs = document.GetElement(p.AsElementId())
                    if legend_fs is None:
                        continue
                    lfam = getattr(legend_fs, 'Family', None)
                    if lfam is not None and eid_int(lfam.Id) == target_fam_int:
                        instances.add(e)
                except Exception:
                    continue
    except Exception as e:
        logger.debug("Legend scan error: %s" % str(e))

    return instances


def get_families(document):
    used, not_used = [], []
    try:
        from Autodesk.Revit.DB import Family
        all_families = list(FilteredElementCollector(document)
                            .OfClass(Family).ToElements())
        dbg("get_families: found %d Family elements" % len(all_families))
    except Exception as e:
        logger.error("get_families: Family collector failed: %s" % str(e))
        return used, not_used

    for idx, fam in enumerate(all_families):
        try:
            if fam is None:
                continue
            try:
                if not fam.IsValidObject:
                    continue
            except Exception:
                continue
            try:
                if fam.IsInPlace:
                    continue
            except Exception:
                pass

            total_instances = 0
            try:
                symbol_ids = fam.GetFamilySymbolIds()
            except Exception:
                symbol_ids = None

            if symbol_ids is not None:
                for sym_id in symbol_ids:
                    try:
                        sym = document.GetElement(sym_id)
                        if sym is None:
                            continue
                        total_instances += len(get_familysymbol_instances(document, sym))
                    except Exception:
                        continue

            if total_instances > 0:
                used.append(fam)
            else:
                not_used.append(fam)
        except Exception as e:
            logger.debug("get_families: error on family %d: %s" % (idx, str(e)))
            continue

    dbg("get_families: %d used, %d not used" % (len(used), len(not_used)))
    return used, not_used


def filter_elements_by_dependencies(document, els, deps):
    in_use, not_in_use = [], []
    for e in els:
        try:
            if e is None or not e.IsValidObject:
                continue
            if eid_int(e.Id) in deps:
                in_use.append(e)
            else:
                not_in_use.append(e)
        except Exception:
            continue
    return in_use, not_in_use


# -------------------- Save / purge / reload --------------------
def save_family(fam_doc, directory, after=False):
    if not doc_is_valid(fam_doc) or not is_string(directory):
        return None, None

    if after:
        directory = os.path.join(directory, "_AFTER")
    if not safe_mkdir(directory):
        return None, None

    fam_file = fam_doc.Title or "unnamed.rfa"
    if not fam_file.lower().endswith(".rfa"):
        fam_file += ".rfa"

    fam_save_path = os.path.join(directory, fam_file)

    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
    opts.MaximumBackups = 1

    try:
        fam_doc.SaveAs(fam_save_path, opts)
        return fam_save_path, fam_file
    except Exception as e:
        logger.error("SaveAs failed for '%s': %s" % (fam_file, str(e)))
        return None, None


def purge_families_in_doc(document, families_not_used):
    """Delete unused families. In workshared docs, only deletes ones we own."""
    not_purged = []
    count = 0
    for f in families_not_used:
        try:
            if f is None or not f.IsValidObject:
                continue
            try:
                if f.IsInPlace:
                    continue
            except Exception:
                pass

            if IS_WORKSHARED and not try_checkout(document, f.Id):
                not_purged.append(f)
                continue

            document.Delete(f.Id)
            count += 1
        except Exception:
            not_purged.append(f)
    return count, len(families_not_used), not_purged


def load_family_back(after_save_path, target_doc, fam_doc=None, fam_name=None):
    """Reload the purged family back into target_doc.

    Preferred path: load from the saved _AFTER file on disk - this avoids any
    issues with the in-memory family doc having unsaved modifications.
    Fallback: call LoadFamily on the open family doc directly.
    """
    global RELOADED_COUNT, RELOAD_FAILED

    if not doc_is_valid(target_doc):
        if fam_name:
            RELOAD_FAILED.append(fam_name)
        return False

    opts = FamilyLoadOption()

    # Primary: load from saved file path (clean, post-purge state)
    if after_save_path and is_string(after_save_path) and os.path.exists(after_save_path):
        try:
            from Autodesk.Revit.DB import Family
            # IronPython returns (bool, Family) for methods with out-params
            ret = target_doc.LoadFamily(after_save_path, opts)
            # ret may be a bool, or a tuple (bool, Family)
            success = ret[0] if isinstance(ret, tuple) else bool(ret)
            if success:
                RELOADED_COUNT += 1
                return True
        except Exception as e:
            print("  reload from path failed (%s): %s" % (fam_name or "?", str(e)))

        # Retry without options (older Revit API)
        try:
            from Autodesk.Revit.DB import Family
            ret = target_doc.LoadFamily(after_save_path)
            success = ret[0] if isinstance(ret, tuple) else bool(ret)
            if success:
                RELOADED_COUNT += 1
                return True
        except Exception as e:
            print("  reload from path (no opts) failed (%s): %s" % (fam_name or "?", str(e)))

    # Fallback: load from the open (in-memory) family document
    if doc_is_valid(fam_doc):
        try:
            ret = target_doc.LoadFamily(fam_doc, opts)
            success = ret[0] if isinstance(ret, tuple) else bool(ret)
            if success:
                RELOADED_COUNT += 1
                return True
        except Exception as e:
            print("  reload from doc failed (%s): %s" % (fam_name or "?", str(e)))

        try:
            ret = target_doc.LoadFamily(fam_doc)
            success = ret[0] if isinstance(ret, tuple) else bool(ret)
            if success:
                RELOADED_COUNT += 1
                return True
        except Exception as e:
            print("  reload fallback failed (%s): %s" % (fam_name or "?", str(e)))

    if fam_name:
        RELOAD_FAILED.append(fam_name)
    return False


# -------------------- CSV --------------------
def write_csv(path, data=None, separator=";"):
    if data is None:
        data = PURGE_RESULTS_CSV
    if not path or not is_string(path):
        return

    lines = []
    for row in data:
        try:
            lines.append(separator.join([str(x) for x in row]))
        except Exception:
            continue

    try:
        with open(path + ".csv", "w") as f:
            f.write("\n".join(lines))
    except Exception as e:
        logger.error("CSV write error: %s" % str(e))


# -------------------- Directory --------------------
def create_directory(filename, top_dir, date=True):
    if not is_string(top_dir):
        top_dir = os.path.join(tempfile.gettempdir(), "pyApex_purgeFamilies")

    ts = "_" + datetime.now().strftime("%y%m%d_%H-%M-%S") if date else ""
    base = sanitise_filename((filename or "doc").replace(".", "_")) + ts
    full = os.path.join(top_dir, base)
    safe_mkdir(full)
    return full


# -------------------- Recursive purge --------------------
def _purge_one_family(document, f, level, max_level, directory, skipped_title, tbs, host_is_workshared):
    """Process a single family: EditFamily -> recursive purge -> reload -> close."""
    global PURGE_SIZES_SUM

    try:
        if f is None or not f.IsValidObject or f.IsInPlace or not f.IsEditable:
            return 0, {}
    except Exception:
        return 0, {}

    fam_name = safe_name(f)

    # Workshared: check ownership before attempting to edit
    if host_is_workshared:
        status = get_checkout_status(document, f.Id)
        if status == 'other':
            owner = get_element_owner(document, f.Id)
            print(tbs + "SKIP (owned by %s): %s" % (owner or "other user", fam_name))
            SKIPPED_NOT_OWNED.append((fam_name, owner))
            return 0, {}
        elif status == 'free':
            if not try_checkout(document, f.Id):
                print(tbs + "SKIP (checkout failed): %s" % fam_name)
                SKIPPED_NOT_OWNED.append((fam_name, "checkout failed"))
                return 0, {}

    # Rename if clashes with host doc name (avoid duplicate name on reload)
    try:
        doc_base = os.path.splitext(document.Title or "")[0]
        if fam_name == doc_base:
            def _rename():
                f.Name = fam_name + "_"
                return True
            ok, _ = run_in_transaction(document, "Rename Family", _rename)
            if ok:
                print(tbs + fam_name + " renamed")
                fam_name = safe_name(f)
            else:
                return 0, {}
    except Exception:
        pass

    if cancelled():
        return 0, {}

    fam_doc = None
    try:
        fam_doc = document.EditFamily(f)
    except Exception as e:
        logger.debug("EditFamily failed '%s': %s" % (fam_name, str(e)))
        return 0, {}

    if not doc_is_valid(fam_doc):
        return 0, {}
    try:
        if not fam_doc.IsFamilyDocument:
            fam_doc.Close(False)
            return 0, {}
    except Exception:
        return 0, {}

    sub_dir = None
    fam_save_path = None
    fam_file = None
    if directory and is_string(directory):
        try:
            fam_save_path, fam_file = save_family(fam_doc, directory)
            if fam_save_path is None:
                try:
                    fam_doc.Close(False)
                except Exception:
                    pass
                return 0, {}
            base = fam_file[:-4] if fam_file.lower().endswith(".rfa") else fam_file
            sub_dir = os.path.join(str(directory), sanitise_filename(base[:32]))
        except Exception as e:
            print(tbs + "Save error: %s - %s" % (fam_doc.Title, str(e)))
            try:
                fam_doc.Close(False)
            except Exception:
                pass
            return 0, {}

    child_purged = 0
    child_by_func = {}
    try:
        fam_doc, child_purged, child_by_func = process_purge(
            fam_doc, parent=document,
            level=level + 1, max_level=max_level,
            directory=sub_dir, skipped_title=skipped_title
        )
    except Exception as e:
        print(tbs + "Recursive purge error '%s': %s" % (fam_name, str(e)))
        try:
            fam_doc.Close(False)
        except Exception:
            pass
        return 0, {}

    fam_by_func = {}
    for pk in child_by_func:
        fam_by_func[pk] = fam_by_func.get(pk, 0) + child_by_func[pk]

    if sub_dir and is_string(sub_dir) and os.path.exists(sub_dir):
        try:
            if not os.listdir(sub_dir):
                os.rmdir(sub_dir)
        except Exception:
            pass

    size_before, size_after, size_diff = 0.0, 0.0, 0.0
    after_save_path = None
    if directory and fam_save_path and os.path.exists(fam_save_path):
        size_before = file_size_mb(fam_save_path)
        if child_purged > 0:
            _p2, _f2 = save_family(fam_doc, directory, after=True)
            after_save_path = _p2  # capture for reload below
            if _p2 and os.path.exists(_p2):
                size_after = file_size_mb(_p2)
                size_diff = size_after - size_before
                if size_diff != 0:
                    try:
                        pct = -(1 - size_after / size_before) * 100 if size_before > 0 else 0
                        print(tbs + "\t%.3f Mb (%.3f -> %.3f, %d%%)" %
                              (size_diff, size_before, size_after, pct))
                        PURGE_SIZES_SUM += size_diff
                    except Exception:
                        pass
        else:
            safe_remove(fam_save_path)

    try:
        fam_type = str(fam_doc.OwnerFamily.FamilyCategory.Name)
    except Exception:
        fam_type = "Not family"

    try:
        csv_row = [
            str(level), fam_type,
            fam_doc.Title or "<untitled>",
            locale.str(size_before),
            locale.str(size_after),
            locale.str(size_diff)
        ]
        for pk in fam_by_func:
            csv_row.append(str(fam_by_func[pk]))
        PURGE_RESULTS_CSV.append(csv_row)
    except Exception:
        pass

    # Reload if the family got smaller (nested families were purged) and not cancelled.
    # Always attempt reload when child_purged > 0, even if size comparison unavailable.
    should_reload = (child_purged > 0) and not cancelled()
    if should_reload:
        if not load_family_back(after_save_path, document, fam_doc=fam_doc, fam_name=fam_name):
            print(tbs + "WARNING: failed to reload '%s' back to host" % fam_name)

    try:
        fam_doc.Close(False)
    except Exception:
        pass

    return child_purged, fam_by_func


def process_purge(document, parent=None, level=0, max_level=1, directory=None, skipped_title=None):
    global PURGE_RESULTS, PURGE_RESULTS_CSV, START_TIME

    if not isinstance(document, Document) or not doc_is_valid(document):
        return document, 0, {}
    if directory is not None and not is_string(directory):
        return document, 0, {}

    if cancelled():
        return document, 0, {}

    dbg("process_purge L%d: entering '%s'" % (level, document.Title or "<untitled>"))

    try:
        families_in_use, families_not_in_use = get_families(document)
    except Exception as e:
        logger.error("get_families failed: %s" % str(e))
        return document, 0, {}

    purged_count = 0
    by_func = {}
    title_printed = False
    tbs = "\t" * level
    host_is_workshared = (level == 0) and IS_WORKSHARED

    if not skipped_title:
        skipped_title = ""
    skipped_title += "\n" + tbs + (document.Title or "<untitled>")

    if parent is None:
        print(skipped_title)
        title_printed = True
        START_TIME = time.time()

    # Nested family docs: prune unused families via dep analysis + delete
    if level > 0:
        if cancelled():
            return document, 0, {}

        try:
            deps = dependencies_structure(document)
        except Exception as e:
            logger.error("dependencies_structure failed: %s" % str(e))
            deps = set()

        in_use, not_in_use = filter_elements_by_dependencies(document, families_not_in_use, deps)
        families_not_in_use = list(set(families_not_in_use) - set(in_use))
        families_in_use = list(set(families_in_use + in_use))

        if families_not_in_use and not cancelled():
            def _do_purge():
                return purge_families_in_doc(document, families_not_in_use)
            ok, res = run_in_transaction(document, "Purge Families", _do_purge)
            if ok and res is not None:
                cnt, found, failed = res
                families_in_use += failed
                purged_count += cnt
                if found > 0 or cnt > 0:
                    PURGE_RESULTS.setdefault("Families", 0)
                    PURGE_RESULTS["Families"] += cnt
                    by_func["Families"] = cnt
                    if not title_printed:
                        print(skipped_title)
                        title_printed = True
                    if cnt != found:
                        print(tbs + "\tFamilies -%d of %d" % (cnt, found))
                    else:
                        print(tbs + "\tFamilies -%d" % cnt)

    families_in_use = list(set(families_in_use))
    total_families = len(families_in_use)

    if level == 0:
        families_not_in_use = list(set(families_not_in_use))
        not_in_use_len = len(families_not_in_use)
        families_in_use += families_not_in_use
        total_families = len(families_in_use)

        _csv = []
        for idx in range(len(families_in_use)):
            try:
                _csv.append([idx, safe_name(families_in_use[idx])])
            except Exception:
                continue

        if directory and is_string(directory):
            try:
                write_csv(os.path.join(directory, (document.Title or "doc") + "_list"), _csv)
            except Exception:
                pass

        print("Families found: %d\n(%d in use, %d not in use)" %
              (total_families, total_families - not_in_use_len, not_in_use_len))
        print("Search took %s seconds\n\nPurge process started" % time_elapsed())
        START_TIME = time.time()

    if level < max_level or max_level is None:
        if title_printed:
            skipped_title = None

        for idx in range(len(families_in_use)):
            if cancelled():
                if parent is None:
                    print("\n*** Cancelled by user at family %d of %d ***" % (idx, total_families))
                return document, purged_count, by_func

            f = families_in_use[idx]
            fam_name = safe_name(f)

            if parent is None:
                try:
                    pct = float(idx) / total_families if total_families > 0 else 0
                    output.update_progress(int(100 * pct), 100)
                    if pct > 0.01:
                        left = time_elapsed() * ((1.0 / pct) - 1)
                        left_txt = " - %s left - %.3f Mb purged" % (time_format(left), PURGE_SIZES_SUM)
                    else:
                        left_txt = ""
                    output.set_title("%s - %d of %d finished%s" %
                                     (window_title, idx, total_families, left_txt))
                except Exception:
                    pass

                update_cancel_status("Family %d of %d:\n%s" % (idx + 1, total_families, fam_name))

            child_purged, fam_by_func = _purge_one_family(
                document, f, level, max_level, directory, skipped_title, tbs, host_is_workshared
            )
            purged_count += child_purged

    return document, purged_count, by_func


# -------------------- Workshared prompts --------------------
def prompt_workshared_options():
    """Returns (proceed, sync_at_end)."""
    try:
        from pyrevit import forms
        res = forms.alert(
            "This model is WORKSHARED.\n\n"
            "The script will:\n"
            "  - Check out each family before editing it\n"
            "  - Skip families owned by other users\n"
            "  - Reload each purged family back into the project\n"
            "  - Operate only on your local copy\n\n"
            "Recommended: run on a detached copy to avoid affecting the central model.\n\n"
            "How do you want to proceed?",
            title="Purge Families - Workshared Model",
            options=["Proceed (no sync)",
                     "Proceed + Sync with Central at end",
                     "Cancel"],
            warn_icon=True
        )
        if res == "Cancel" or res is None:
            return False, False
        if res == "Proceed + Sync with Central at end":
            return True, True
        return True, False
    except Exception as e:
        logger.debug("Workshared prompt failed, defaulting to cancel: %s" % str(e))
        print("ERROR: Could not show workshared prompt - aborting for safety.")
        return False, False


# -------------------- Main --------------------
def validate_active_doc():
    if not doc_is_valid(ACTIVE_DOC):
        print("ERROR: No valid active document.")
        return False
    try:
        if ACTIVE_DOC.IsReadOnly:
            print("ERROR: Active document is read-only.")
            return False
    except Exception:
        pass
    return True


def prepare_output_dir():
    """Validate / create the staging directory used to SaveAs family snapshots."""
    purge_dir = config_temp_dir()
    dbg("Temp dir: %s" % purge_dir)

    purge_dir_lower = purge_dir.lower()
    risky_markers = ["\\shellfolders\\", "\\onedrive", "\\sharepoint"]
    for marker in risky_markers:
        if marker in purge_dir_lower:
            print("ERROR: Temp directory is on a OneDrive / redirected folder: %s" % purge_dir)
            print("       These paths can cause Revit to crash during family SaveAs operations.")
            print("       Please use the Config button to pick a local drive path, e.g.:")
            print("         C:\\Temp\\PurgeFamilies")
            return None

    if not safe_mkdir(purge_dir):
        print("ERROR: Cannot create/access temp directory: %s" % purge_dir)
        return None

    try:
        test_file = os.path.join(purge_dir, ".write_test")
        with open(test_file, "w") as f:
            f.write("ok")
        os.remove(test_file)
    except Exception as e:
        print("ERROR: Temp directory is not writable: %s" % str(e))
        return None

    directory = create_directory(ACTIVE_DOC.Title or "document", purge_dir)
    if not os.path.isdir(directory):
        print("ERROR: Output directory was not created: %s" % directory)
        return None
    return directory


def main():
    global PURGE_RESULTS_CSV, START_TIME, CANCEL_REQUESTED
    global IS_WORKSHARED, SYNC_AT_END, CHECKED_OUT_IDS, SKIPPED_NOT_OWNED
    global PURGE_RESULTS, PURGE_SIZES_SUM, RELOADED_COUNT, RELOAD_FAILED

    # Reset per-run state
    CANCEL_REQUESTED = False
    CHECKED_OUT_IDS = []
    SKIPPED_NOT_OWNED = []
    PURGE_RESULTS = {}
    PURGE_SIZES_SUM = 0.0
    RELOADED_COUNT = 0
    RELOAD_FAILED = []

    dbg("Startup")
    if not validate_active_doc():
        return
    dbg("Document: %s" % (ACTIVE_DOC.Title or "<untitled>"))

    IS_WORKSHARED = doc_is_workshared(ACTIVE_DOC)
    SYNC_AT_END = False

    if IS_WORKSHARED:
        dbg("Document is workshared - prompting user")
        proceed, SYNC_AT_END = prompt_workshared_options()
        if not proceed:
            print("Cancelled by user at workshared prompt.")
            return
        if not HAS_CHECKOUT_STATUS:
            print("WARNING: CheckoutStatus API unavailable - checkout verification disabled.")
        dbg("Workshared mode: proceeding (sync_at_end=%s)" % SYNC_AT_END)

    directory = prepare_output_dir()
    if directory is None:
        return
    dbg("Output dir: %s" % directory)

    # Determine starting level: if this is an RFA itself, skip the host pass
    level = 0
    max_level = 99
    if (ACTIVE_DOC.Title or "")[-4:].lower() == ".rfa":
        level = 1
    try:
        if __forceddebugmode__:
            max_level = 1
            dbg("Debug mode - max_level=1")
    except Exception:
        pass

    START_TIME = time.time()
    PURGE_RESULTS_CSV = [["Level", "Category", "Family name", "Size before",
                          "Size after", "Size diff", "Purged: Families"]]

    if show_cancel_window():
        dbg("Cancel window shown")

    dbg("Starting family collection...")
    try:
        process_purge(ACTIVE_DOC, level=level, max_level=max_level, directory=directory)
    except Exception as e:
        logger.error("Purge process error: %s" % str(e))
        print("ERROR: Purge process failed - %s" % str(e))
        try:
            print(traceback.format_exc())
        except Exception:
            pass
    finally:
        close_cancel_window()

    # Workshared cleanup
    if IS_WORKSHARED:
        try:
            if SYNC_AT_END and not CANCEL_REQUESTED:
                sync_with_central(ACTIVE_DOC)
            else:
                if CANCEL_REQUESTED:
                    print("\nReleasing checked-out elements (changes in local only, not pushed to central)...")
                    relinquish_all_checked_out(ACTIVE_DOC)
                else:
                    print("\nNOTE: Purge completed in local session only.")
                    print("      Sync with Central manually to push changes and release checkouts.")
        except Exception as e:
            logger.error("Post-run workshared cleanup error: %s" % str(e))

    try:
        output.set_title("%s - Write CSV" % window_title)
    except Exception:
        pass
    write_csv(directory)
    try:
        output.set_title("%s - Done" % window_title)
    except Exception:
        pass

    if CANCEL_REQUESTED:
        print("\n\n=== CANCELLED BY USER ===")
        print("(Partial results below)")
    else:
        print("\n\nFinished")

    for r in PURGE_RESULTS.keys():
        print("%s: %d" % (r, PURGE_RESULTS[r]))
    total = sum(PURGE_RESULTS.values()) if PURGE_RESULTS else 0
    print('\nTOTAL purged: %d' % total)
    print('SIZE DIFFERENCE: %.3f Mb' % PURGE_SIZES_SUM)
    print('Families reloaded back into project: %d' % RELOADED_COUNT)

    if RELOAD_FAILED:
        print('\n--- Families that FAILED to reload ---')
        for nm in RELOAD_FAILED:
            print("  %s" % nm)

    if IS_WORKSHARED and SKIPPED_NOT_OWNED:
        print('\n--- Skipped (workshared ownership) ---')
        seen = set()
        for name, reason in SKIPPED_NOT_OWNED:
            key = (name, reason)
            if key in seen:
                continue
            seen.add(key)
            print("  %s  [%s]" % (name, reason or "unknown"))
        print("Total skipped: %d" % len(seen))

    print("\n--- %s ---" % time_format(time_elapsed()))

    try:
        output.update_progress(100, 100)
    except Exception:
        pass


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        try:
            print("FATAL: %s" % str(e))
            print(traceback.format_exc())
        except Exception:
            pass
        try:
            close_cancel_window()
        except Exception:
            pass
