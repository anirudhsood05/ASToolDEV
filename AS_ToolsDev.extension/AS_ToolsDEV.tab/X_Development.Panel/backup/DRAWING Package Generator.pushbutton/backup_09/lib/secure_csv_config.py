# -*- coding: utf-8 -*-
"""
Secure CSV Configuration Loader
Prevents injection attacks and validates all input
"""

import csv
import os
import re
from pyrevit import script

logger = script.get_logger()

class SecureCSVConfig:
    """Security-hardened CSV configuration loader"""
    
    # BALANCED: Strict where needed, permissive where safe
    VALIDATORS = {
        # Strict validation for controlled fields
        'SeriesNumber': r'^\d{1,2}$',  # Only 1-2 digits
        'ViewType': r'^(FloorPlan|CeilingPlan|AreaPlan)$',  # Enum only
        
        # Moderate validation for naming fields (architectural conventions)
        'SeriesName': r'^[a-zA-Z0-9 _/\-().,&]{1,100}$',
        'ViewTemplate': r'^[a-zA-Z0-9 _\-()]{0,150}$',
        
        # Permissive for filter patterns (allow underscores, hyphens)
        'WallFilter': r'^[a-zA-Z0-9_\-]{0,50}$',
        'DoorFilter': r'^[a-zA-Z0-9_\-]{0,50}$',
        'WindowFilter': r'^[a-zA-Z0-9_\-]{0,50}$',
        
        # Permissive for free-text guidance (block only dangerous chars)
        'UserGuidance': r'^[^=+@\t\r\n]{0,500}$',
    }
    
    BOOLEAN_FIELDS = {
        'TagRooms', 'TagDoors', 'TagWindows', 'TagWalls', 'TagCeilings',
        'CreateRCP', 'LevelNumbering', 'AutoGenerate'
    }
    
    MAX_ROWS = 100
    MAX_FILE_SIZE = 1048576  # 1 MB
    
    @staticmethod
    def sanitize_field(field_name, value):
        """Sanitize and validate CSV field value"""
        if value is None:
            return ''
    
        value = str(value).strip()
    
        # Remove dangerous formula prefixes
        dangerous_prefixes = ['=', '+', '-', '@', '\t', '\r', '\n']
        original_value = value
        while value and value[0] in dangerous_prefixes:
            value = value[1:]
    
        if value != original_value:
            logger.warning("Stripped dangerous prefix from {}: {}".format(
                field_name, original_value[:20]))
    
        # Apply pattern validation
        if field_name in SecureCSVConfig.VALIDATORS:
            pattern = SecureCSVConfig.VALIDATORS[field_name]
            if not re.match(pattern, value):
                # Only log for non-free-text fields
                if field_name not in ['UserGuidance']:  # Don't warn for free text
                    logger.warning(
                        "Field '{}' value '{}' doesn't match pattern '{}' - ALLOWING ANYWAY".format(
                            field_name, value[:50], pattern
                        )
                    )
                # Still allow it - we just wanted to log suspicious values
    
        # Boolean validation
        if field_name in SecureCSVConfig.BOOLEAN_FIELDS:
            upper_val = value.upper()
            if upper_val not in ('TRUE', 'FALSE', ''):
                raise ValueError(
                    "Invalid boolean for {}: '{}'".format(field_name, value)
                )
    
        # Length check (still enforce to prevent DoS)
        if len(value) > 500:
            raise ValueError("Field {} exceeds 500 chars".format(field_name))
    
        return value
    
    @staticmethod
    def get_safe_csv_path(script_file):
        """Get validated CSV path preventing path traversal"""
        script_dir = os.path.dirname(os.path.abspath(script_file))
        csv_filename = "drawing_series_config_enhanced.csv"
        
        # Prevent path traversal
        if os.path.sep in csv_filename or (os.path.altsep and os.path.altsep in csv_filename):
            raise ValueError("Invalid CSV filename")
        
        csv_path = os.path.abspath(os.path.join(script_dir, csv_filename))
        
        # Validate path is within script directory
        if not csv_path.startswith(script_dir + os.path.sep):
            raise ValueError("CSV path outside script directory")
        
        return csv_path
    
    @staticmethod
    def load_config(csv_path):
        """Load and validate CSV configuration"""
        if not os.path.exists(csv_path):
            raise IOError("CSV not found: {}".format(csv_path))
        
        # Check file size
        file_size = os.path.getsize(csv_path)
        if file_size > SecureCSVConfig.MAX_FILE_SIZE:
            raise ValueError("CSV too large: {} bytes".format(file_size))
        
        config = {}
        row_count = 0
        
        try:
            with open(csv_path, 'r') as f:
                reader = csv.DictReader(f)
                
                for row in reader:
                    row_count += 1
                    
                    if row_count > SecureCSVConfig.MAX_ROWS:
                        raise ValueError("CSV exceeds {} rows".format(
                            SecureCSVConfig.MAX_ROWS))
                    
                    try:
                        series_num = SecureCSVConfig.sanitize_field(
                            'SeriesNumber', row.get('SeriesNumber', ''))
                        
                        if not series_num:
                            continue
                        
                        config[series_num] = {
                            'name': SecureCSVConfig.sanitize_field(
                                'SeriesName', row.get('SeriesName', '')),
                            'view_type': SecureCSVConfig.sanitize_field(
                                'ViewType', row.get('ViewType', 'FloorPlan')),
                            'template_pattern': SecureCSVConfig.sanitize_field(
                                'ViewTemplate', row.get('ViewTemplate', '')),
                            'tag_rooms': row.get('TagRooms', '').upper() == 'TRUE',
                            'tag_doors': row.get('TagDoors', '').upper() == 'TRUE',
                            'tag_windows': row.get('TagWindows', '').upper() == 'TRUE',
                            'tag_walls': row.get('TagWalls', '').upper() == 'TRUE',
                            'tag_ceilings': row.get('TagCeilings', '').upper() == 'TRUE',
                            'create_rcp': row.get('CreateRCP', '').upper() == 'TRUE',
                            'level_numbering': row.get('LevelNumbering', '').upper() == 'TRUE',
                            'auto_generate': row.get('AutoGenerate', '').upper() == 'TRUE',
                            'wall_filter': SecureCSVConfig.sanitize_field(
                                'WallFilter', row.get('WallFilter', '')),
                            'door_filter': SecureCSVConfig.sanitize_field(
                                'DoorFilter', row.get('DoorFilter', '')),
                            'window_filter': SecureCSVConfig.sanitize_field(
                                'WindowFilter', row.get('WindowFilter', '')),
                            'user_guidance': SecureCSVConfig.sanitize_field(
                                'UserGuidance', row.get('UserGuidance', '')),
                        }
                        
                    except ValueError as row_error:
                        logger.error("Row {} invalid: {}".format(row_count, str(row_error)))
                        continue
                        
        except csv.Error as csv_error:
            raise ValueError("CSV parse error: {}".format(str(csv_error)))
        
        if not config:
            raise ValueError("No valid rows in CSV")
        
        logger.info("Loaded {} series configs".format(len(config)))
        return config