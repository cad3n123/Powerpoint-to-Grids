# Grid Maker Tool

Creates a contact sheet (grid) of slides from PowerPoint files.

## For the End User (How to Run)

1. Double-click `grid_maker.exe`.
2. **If you don't have LibreOffice**: The app will detect this and ask permission to install it for you automatically.
3. **Poppler**: You do NOT need to install this. It is built into the app.

## For the Developer (How to Build)

1. Save `grid_maker.py` and `build_windows.bat`.
2. Run `build_windows.bat`.
   - This script automatically downloads Python, downloads Poppler, bundles everything, and produces the final `.exe`.
