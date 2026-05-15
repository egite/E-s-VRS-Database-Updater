"""Allow running as: python -m vrs_updater"""
import sys

if "--no-gui" in sys.argv or "--help" in sys.argv or "-h" in sys.argv:
    from .main import main
    main()
else:
    from .gui import run_gui
    run_gui()
