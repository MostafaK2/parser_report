# PyInstaller build recipe for the DAR app.
# Build with:  pyinstaller --noconfirm --clean dar_app.spec
# Result:      dist/ShiftReportDAR/ShiftReportDAR(.exe)  - copy the whole ShiftReportDAR folder.

a = Analysis(
    ["dar_app.py"],
    pathex=[],
    datas=[("report_config.json", ".")],   # bundled default config (an edited copy next to the .exe wins)
    hiddenimports=[],
    excludes=["pytest", "PIL", "numpy", "pandas"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ShiftReportDAR",
    console=False,        # window app, no black console window
    upx=False,
)
coll = COLLECT(exe, a.binaries, a.datas, name="ShiftReportDAR", upx=False)
