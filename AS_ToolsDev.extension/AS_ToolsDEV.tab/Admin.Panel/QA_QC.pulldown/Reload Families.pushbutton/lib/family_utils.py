# -*- coding: utf-8 -*-
""" Enhanced Family Loader with Reload Capability """
import os
import re
from pyrevit.framework import clr
from pyrevit import forms, revit, DB, script

logger = script.get_logger()

class FamilyReloadOptions(object):
    """Custom family load options for forced reloading"""
    
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        """
        Called when family already exists in project
        Returns True to overwrite, sets overwriteParameterValues to True
        """
        overwriteParameterValues = True
        return True  # Always overwrite existing families
    
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        """
        Called when shared family is found
        Returns True to continue loading
        """
        overwriteParameterValues = True
        return True

class SmartSortableFamilySymbol(object):
    """
    Enables smart sorting of family symbols.
    (Copied from original for compatibility)
    """
    def __init__(self, symbol_name):
        self.symbol_name = symbol_name
        self.sort_alphabetically = False
        self.number_list = [
            int(x)
            for x in re.findall(r'\d+', self.symbol_name)]
        if not self.number_list:
            self.sort_alphabetically = True

    def __str__(self):
        return self.symbol_name

    def __repr__(self):
        return '<SmartSortableFamilySymbol Name:{} Values:{} StringSort:{}>'\
               .format(self.symbol_name,
                       self.number_list,
                       self.sort_alphabetically)

    def __eq__(self, other):
        return self.symbol_name == other.symbol_name

    def __hash__(self):
        return hash(self.symbol_name)

    def __lt__(self, other):
        if self.sort_alphabetically or other.sort_alphabetically:
            return self.symbol_name < other.symbol_name
        else:
            return self.number_list < other.number_list

class FamilyLoaderEnhanced(object):
    """
    Enhanced family loader with reload capability
    
    Attributes
    ----------
    path : str
        Absolute path to family .rfa file
    name : str
        File name without extension
    is_loaded : bool
        Checks if family name already exists in project
    existing_family : Family
        Reference to existing family in project (if any)
    
    Methods
    -------
    reload_family()
        Reloads family with forced overwrite
    get_family_info()
        Returns detailed information about family status
    get_symbols()
        Returns all family symbols for selective loading
    load_selective()
        Loads family with user-selected symbols
    load_all()
        Loads family with all symbols
    """
    
    def __init__(self, path):
        """
        Parameters
        ----------
        path : str
            Absolute path to family .rfa file
        """
        self.path = path
        self.name = os.path.basename(path).replace(".rfa", "")
        self._existing_family = None
        self._check_existing_family()
    
    def _check_existing_family(self):
        """Check if family already exists in project"""
        collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.Family)
        for family in collector:
            if family.Name == self.name:
                self._existing_family = family
                break
    
    @property
    def is_loaded(self):
        """Returns True if family is already loaded in project"""
        return self._existing_family is not None
    
    @property
    def existing_family(self):
        """Returns existing family element if loaded"""
        return self._existing_family
    
    def get_family_info(self):
        """
        Returns detailed information about family
        
        Returns
        -------
        dict
            Dictionary containing family information
        """
        info = {
            'name': self.name,
            'path': self.path,
            'is_loaded': self.is_loaded,
            'symbol_count': 0,
            'status': 'New' if not self.is_loaded else 'Existing'
        }
        
        if self.is_loaded and self._existing_family:
            try:
                symbol_ids = self._existing_family.GetFamilySymbolIds()
                info['symbol_count'] = len(symbol_ids)
            except:
                info['symbol_count'] = 0
        
        return info
    
    def reload_family(self):
        """
        Reload family with forced overwrite
        
        Returns
        -------
        tuple
            (success: bool, message: str, family: Family or None)
        """
        try:
            # Create custom load options for forced overwrite
            load_options = FamilyReloadOptions()
            
            # Load family with overwrite options
            ret_ref = clr.Reference[DB.Family]()
            success = revit.doc.LoadFamily(self.path, load_options, ret_ref)
            
            if success:
                loaded_family = ret_ref.Value
                # Update our reference to the family
                self._existing_family = loaded_family
                
                # Get symbol count for reporting
                symbol_count = len(loaded_family.GetFamilySymbolIds())
                
                if self.is_loaded:
                    message = "Successfully reloaded family '{}' with {} types".format(
                        self.name, symbol_count)
                else:
                    message = "Successfully loaded new family '{}' with {} types".format(
                        self.name, symbol_count)
                
                return True, message, loaded_family
            else:
                return False, "Failed to load family '{}'".format(self.name), None
                
        except Exception as e:
            error_msg = "Error loading family '{}': {}".format(self.name, str(e))
            logger.error(error_msg)
            return False, error_msg, None
    
    def get_symbols(self):
        """
        Get all family symbols for selective loading
        (Enhanced version of original method)
        
        Returns
        -------
        set
            Set of SmartSortableFamilySymbol objects
        """
        logger.debug('Getting symbols for family: {}'.format(self.name))
        symbol_set = set()
        
        try:
            with revit.ErrorSwallower():
                # DryTransaction will rollback all the changes
                with revit.DryTransaction('Fake load for symbols'):
                    ret_ref = clr.Reference[DB.Family]()
                    revit.doc.LoadFamily(self.path, ret_ref)
                    loaded_fam = ret_ref.Value
                    
                    # Get the symbols
                    for symbol_id in loaded_fam.GetFamilySymbolIds():
                        symbol = revit.doc.GetElement(symbol_id)
                        symbol_name = symbol.Name if symbol else "Unknown"
                        sortable_sym = SmartSortableFamilySymbol(symbol_name)
                        logger.debug('Importable Symbol: {}'.format(sortable_sym))
                        symbol_set.add(sortable_sym)
                        
        except Exception as e:
            logger.error("Error getting symbols for family '{}': {}".format(self.name, str(e)))
        
        return sorted(symbol_set)
    
    def load_selective(self):
        """
        Loads the family and selected symbols
        (Enhanced version with reload capability)
        """
        symbols = self.get_symbols()

        # Don't prompt if only 1 symbol available
        if len(symbols) == 1:
            return self.reload_family()

        # User input -> Select family symbols
        selected_symbols = forms.SelectFromList.show(
            symbols,
            title=self.name,
            button_name="Load type(s)",
            multiselect=True)
            
        if selected_symbols is None:
            logger.debug('No family symbols selected.')
            return False, "No symbols selected", None

        logger.debug('Selected symbols are: {}'.format(selected_symbols))

        # Load family with selected symbols using reload options
        try:
            load_options = FamilyReloadOptions()
            
            with revit.Transaction('Loaded {}'.format(self.name)):
                for symbol in selected_symbols:
                    logger.debug('Loading symbol: {}'.format(symbol))
                    revit.doc.LoadFamilySymbol(self.path, symbol.symbol_name, load_options)
                
                logger.debug('Successfully loaded all selected symbols')
                return True, "Successfully loaded {} selected symbols".format(len(selected_symbols)), None
                
        except Exception as load_err:
            error_msg = 'Error loading family symbol from {} | {}'.format(self.path, load_err)
            logger.error(error_msg)
            return False, error_msg, None

    def load_all(self):
        """
        Loads family and all its symbols
        (Enhanced version with reload capability)
        """
        return self.reload_family()
    
    def __str__(self):
        return self.name
    
    def __repr__(self):
        return "<FamilyLoaderEnhanced: {} (Loaded: {})>".format(self.name, self.is_loaded)

# Maintain compatibility with original FamilyLoader class name
# This allows existing scripts to work without modification
class FamilyLoader(FamilyLoaderEnhanced):
    """
    Compatibility wrapper for original FamilyLoader
    Inherits all enhanced functionality while maintaining original interface
    """
    pass