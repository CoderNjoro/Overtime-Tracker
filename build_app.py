import subprocess
import sys
import venv
from pathlib import Path
import shutil


ROOT = Path(__file__).parent.resolve()
# Reuse the existing project virtualenv if present to avoid slow venv creation.
# Fallback to a dedicated build venv if needed.
PREF_VENV_DIR = ROOT / ".venv"
FALLBACK_VENV_DIR = ROOT / ".venv_build"
VENV_DIR = PREF_VENV_DIR if PREF_VENV_DIR.exists() else FALLBACK_VENV_DIR
SCRIPT = ROOT / "overtime_calculator.py"
# Optional application icon (ICO); if present, it will be used for the EXE.
ICON = ROOT / "app_icon.ico"

# Packages that must be installed in the build venv so the resulting EXE works
# on other PCs without Python.
REQUIRED_IMPORTS = [
    "PyInstaller",
    "pandas",
    "openpyxl",
    "pdfplumber",
    "xlsxwriter",
    "xlrd",   # needed for legacy .xls support via pandas
    "lxml",   # needed for pandas read_html (e.g. .xls that are actually HTML)
    # html5lib is an optional fallback for pandas read_html, but lxml is sufficient.
]


def run(cmd, **kwargs):
    print(f"\n>> {' '.join(cmd)}")
    subprocess.check_call(cmd, **kwargs)


def ensure_venv():
    if VENV_DIR.exists():
        print(f"Using existing virtualenv at {VENV_DIR} ...")
        return

    # This path runs only if neither .venv nor .venv_build exists.
    # Creating the venv (especially ensurepip) can take a while.
    print(f"Creating build virtualenv at {VENV_DIR} ... (this may take a few minutes)")
    venv.EnvBuilder(with_pip=True).create(VENV_DIR)


def venv_python() -> Path:
    if sys.platform.startswith("win"):
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def main():
    if not SCRIPT.exists():
        print(f"ERROR: {SCRIPT.name} not found next to this script.")
        sys.exit(1)

    ensure_venv()
    py = str(venv_python())

    # Dependency install step can be very slow on some networks, so we do not auto-install.
    # Instead, we validate required imports and tell you exactly what to install if missing.
    print("Using existing virtualenv and its installed packages (no auto pip install).")

    print("Checking required libraries in the virtualenv...")
    missing = []
    for mod in REQUIRED_IMPORTS:
        try:
            run([py, "-c", f"import {mod}"])
        except subprocess.CalledProcessError:
            missing.append(mod)

    if missing:
        pkgs = " ".join(m.lower() for m in missing)
        print("\nERROR: Missing required libraries in this virtualenv:")
        for m in missing:
            print(f"  - {m}")
        print("\nInstall them once, then rerun this builder:")
        print(f"  {py} -m pip install {pkgs}")
        sys.exit(1)

    # Ensure we're not accidentally reusing a previous build (common cause of "missing module"
    # at runtime after installing new deps).
    for p in (ROOT / "build", ROOT / "dist"):
        if p.exists():
            shutil.rmtree(p, ignore_errors=True)

    print("Building standalone EXE with PyInstaller...")
    cmd = [
        py,
        "-m",
        "PyInstaller",
        "--clean",
        "--noconsole",
        "--onefile",
        "--name",
        "OvertimeCalculator",
        # pandas loads these dynamically; force-include so EXE works without Python.
        "--hidden-import",
        "xlrd",
        "--hidden-import",
        "pandas.io.excel._xlrd",
        "--hidden-import",
        "lxml",
        "--hidden-import",
        "lxml.etree",
        # Ensure the standard library 'xml' package is included (required by pandas HTML fallback)
        "--hidden-import",
        "xml",
        "--hidden-import",
        "xml.etree.ElementTree",
        # Optional: include html5lib if you want an extra fallback parser
        # "--hidden-import", "html5lib",
    ]
    if ICON.exists():
        cmd += ["--icon", str(ICON)]
    cmd.append(str(SCRIPT))

    run(cmd, cwd=str(ROOT))

    dist_exe = ROOT / "dist" / "OvertimeCalculator.exe"
    if not dist_exe.exists():
        print("Build did not produce expected EXE. Check PyInstaller output above.")
        sys.exit(1)

    print("\nBuild finished successfully.")
    print(f"Executable created at:\n  {dist_exe}")
    print(
        "\nYou can now copy this EXE to a USB drive and run it on other "
        "Windows machines like a normal app (no Python needed)."
    )

    # Also generate an Inno Setup script so you can build a proper installer.
    iss_path = ROOT / "OvertimeCalculatorInstaller.iss"
    iss_content = f"""\
; Inno Setup script for Overtime Calculator
; 1) Install Inno Setup from https://jrsoftware.org/
; 2) Open this .iss file in Inno Setup Compiler
; 3) Build to generate OvertimeCalculatorSetup.exe installer

[Setup]
AppName=Overtime Calculator
AppVersion=1.0.0
DefaultDirName={{pf}}\\Overtime Calculator
DefaultGroupName=Overtime Calculator
OutputBaseFilename=OvertimeCalculatorSetup
Compression=lzma
SolidCompression=yes
DisableDirPage=no
DisableProgramGroupPage=no

[Files]
Source: "{dist_exe}"; DestDir: "{{app}}"; Flags: ignoreversion

[Icons]
Name: "{{group}}\\Overtime Calculator"; Filename: "{{app}}\\OvertimeCalculator.exe"
Name: "{{commondesktop}}\\Overtime Calculator"; Filename: "{{app}}\\OvertimeCalculator.exe"
"""
    iss_path.write_text(iss_content, encoding="utf-8")
    print(f"\nInstaller script generated at:\n  {iss_path}")
    print(
        "To create a full Windows installer (setup .exe):\n"
        "  1) Install Inno Setup (if not already).\n"
        "  2) Open this .iss file in Inno Setup Compiler.\n"
        "  3) Click Build → Compile.\n"
        "The generated installer .exe can be put on USB and run on other PCs."
    )


if __name__ == "__main__":
    main()