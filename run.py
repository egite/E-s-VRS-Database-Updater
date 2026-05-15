"""PyInstaller entry point — avoids relative import issues."""
import sys

if "--no-gui" in sys.argv or "--help" in sys.argv or "-h" in sys.argv:
    from vrs_updater.main import main
    main()
else:
    from vrs_updater.gui import run_gui
    run_gui()
