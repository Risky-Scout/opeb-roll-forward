"""
OPEB Roll-Forward Model

Production-ready GASB 75 roll-forward model for year-over-year
OPEB liability projections, experience analysis, and Excel automation.

Version: 2.1.0 (West Florida Planning Corrections - 2025-12-28)
Author: Actuarial Pipeline Project
"""

__version__ = "2.1.0"
__author__ = "Actuarial Pipeline Project"

from .engine import (
    RollForwardEngine,
    PriorValuation,
    RollForwardResults as EngineResults,
    create_engine,
    load_prior
)

from .excel_updater import (
    # Main functions
    run_roll_forward,
    update_roll_forward_excel,
    calculate_roll_forward,
    verify_roll_forward_output,
    print_roll_forward_summary,
    
    # Data classes
    RollForwardInputs,
    RollForwardResults,
    
    # Helper functions
    copy_cell_format,
    adjust_formula_row,
)

__all__ = [
    # Engine
    "RollForwardEngine",
    "PriorValuation", 
    "EngineResults",
    "create_engine",
    "load_prior",
    
    # Excel Updater
    "run_roll_forward",
    "update_roll_forward_excel",
    "calculate_roll_forward",
    "verify_roll_forward_output",
    "print_roll_forward_summary",
    "RollForwardInputs",
    "RollForwardResults",
    "copy_cell_format",
    "adjust_formula_row",
]
