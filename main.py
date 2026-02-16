#!/usr/bin/env python3
"""
Entry point for the DSF Pipeline execution.
"""
import sys
from pathlib import Path

# Add src to the python path so imports work correctly
sys.path.append(str(Path(__file__).parent / "src"))

try:
    from dsf_pipeline import main
except ImportError:
    # Fallback if specific import fails or structure is slightly different
    from src.dsf_pipeline import main

if __name__ == "__main__":
    main()
