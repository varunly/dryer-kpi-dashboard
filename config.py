"""
Configuration for Dryer KPI applications
"""

import os

# File paths - use environment variables for flexibility
CONFIG = {
    "energy_file": os.getenv("ENERGY_FILE", "data/Energieverbrauch Trockner 1, Stundenweise - Januar - September 2025.xlsx"),
    "energy_sheet": 0,
    "wagon_file": os.getenv("WAGON_FILE", "data/Hordenwagenverfolgung_Stand 2025_10_12.xlsm"),
    "wagon_sheet": "Hordenwagenverfolgung",
    "wagon_header_row": 6,
    "gas_to_kwh": 11.5,
    "takt_minutes": 65,
    "zones_seq": ["Z1", "Z2", "Z3", "Z4", "Z5"],
    "product_filter": None,
    "month_filter": None,
    "output_file": "Dryer_KPI_Results.xlsx",
    # Production configuration
    "plates_per_row": 4,
    "rows_per_wagen": 39,
    "plates_per_wagen": 156,
    "plate_width_m": 0.605,
    "plate_length_m": 0.605,
    "cutting_allowance_mm": 7
}

# Zone to column mapping
ZONE_ENERGY_MAPPING = {
    "Z2": "Zone 2",
    "Z3": "Zone 3",
    "Z4": "Zone 4",
    "Z5": "Zone 5"
}
