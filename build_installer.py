"""Build a Windows distributable for the GA scheduler GUI.

輸出分成三層：
1. PyInstaller 先在系統暫存資料夾建立乾淨可攜版，避免 OneDrive 鎖檔。
2. `dist/FactoryGAScheduler_latest.../FactoryGAScheduler.exe`：專案內新版可攜版副本。
3. `installer/FactoryGAScheduler_portable_latest.zip` 與
   `installer/FactoryGAScheduler_SelfExtracting_Setup_latest.exe`：交付用壓縮檔與一鍵自解壓啟動包。

若電腦有安裝 Inno Setup 6，會再額外產生 `installer/FactoryGAScheduler_Setup.exe`。
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
APP_NAME = "FactoryGAScheduler"
APP_TITLE = "工廠訂單排程 GA 工具"
ENTRY_SCRIPT = PROJECT_DIR / "ga_tkinter_app.py"
DIST_APP_DIR = PROJECT_DIR / "dist" / APP_NAME
PUBLISHED_APP_DIR = PROJECT_DIR / "dist" / f"{APP_NAME}_latest"
INSTALLER_DIR = PROJECT_DIR / "installer"
BUILD_TMP_DIR = Path(tempfile.gettempdir()) / f"{APP_NAME}_build"
INNO_SCRIPT = INSTALLER_DIR / f"{APP_NAME}.iss"
IEXPRESS_SCRIPT = INSTALLER_DIR / f"{APP_NAME}_iexpress.sed"
PORTABLE_ZIP = INSTALLER_DIR / f"{APP_NAME}_portable.zip"
IEXPRESS_SETUP = INSTALLER_DIR / f"{APP_NAME}_SelfExtracting_Setup.exe"
DEMO_ASSETS = (
    (PROJECT_DIR / "workbook_bom_refresh.xlsx", "workbook_bom_refresh.xlsx"),
    (PROJECT_DIR.parent / "calendar_capacity.xlsx", "calendar_capacity.xlsx"),
    (PROJECT_DIR / "README.md", "README.md"),
)


def run(command: list[str]) -> None:
    print("執行：", " ".join(command), flush=True)
    subprocess.run(command, cwd=PROJECT_DIR, check=True)


def has_pyinstaller() -> bool:
    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--version"],
        cwd=PROJECT_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )
    if result.returncode == 0:
        print(f"PyInstaller：{result.stdout.strip()}")
        return True
    return False


def find_inno_compiler() -> Path | None:
    from_path = shutil.which("ISCC")
    if from_path:
        return Path(from_path)

    candidates = [
        Path(r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"),
        Path(r"C:\Program Files\Inno Setup 6\ISCC.exe"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def find_iexpress() -> Path | None:
    from_path = shutil.which("IExpress.exe") or shutil.which("iexpress.exe")
    if from_path:
        return Path(from_path)

    candidate = Path(r"C:\Windows\System32\iexpress.exe")
    if candidate.exists():
        return candidate
    return None


def write_inno_script() -> None:
    INSTALLER_DIR.mkdir(parents=True, exist_ok=True)
    script = f"""
#define MyAppName "{APP_TITLE}"
#define MyAppVersion "1.0.0"
#define MyAppExeName "{APP_NAME}.exe"

[Setup]
AppId={{{{D23451F7-2D75-4B8D-A8FD-8C965C3F2291}}}}
AppName={{#MyAppName}}
AppVersion={{#MyAppVersion}}
DefaultDirName={{autopf}}\\{APP_NAME}
DefaultGroupName={{#MyAppName}}
OutputDir=.
OutputBaseFilename={APP_NAME}_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "chinesetraditional"; MessagesFile: "compiler:Languages\\ChineseTraditional.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{DIST_APP_DIR}\\*"; DestDir: "{{app}}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{{group}}\\{{#MyAppName}}"; Filename: "{{app}}\\{{#MyAppExeName}}"
Name: "{{commondesktop}}\\{{#MyAppName}}"; Filename: "{{app}}\\{{#MyAppExeName}}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "建立桌面捷徑"; GroupDescription: "其他選項："; Flags: unchecked

[Run]
Filename: "{{app}}\\{{#MyAppExeName}}"; Description: "啟動 {{#MyAppName}}"; Flags: nowait postinstall skipifsilent
""".strip()
    INNO_SCRIPT.write_text(script + "\n", encoding="utf-8")
    print(f"已產生 Inno Setup 腳本：{INNO_SCRIPT}")


def build_pyinstaller_app() -> None:
    global DIST_APP_DIR

    if not ENTRY_SCRIPT.exists():
        raise FileNotFoundError(f"找不到 GUI 入口：{ENTRY_SCRIPT}")
    if not has_pyinstaller():
        raise RuntimeError(
            "找不到 PyInstaller。請先執行：\n"
            "python -m pip install -r requirements-build.txt"
        )

    BUILD_TMP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    work_path = BUILD_TMP_DIR / f"pyinstaller_work_{timestamp}"
    spec_path = BUILD_TMP_DIR / f"pyinstaller_spec_{timestamp}"
    dist_path = BUILD_TMP_DIR / f"pyinstaller_dist_{timestamp}"
    DIST_APP_DIR = dist_path / APP_NAME

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--windowed",
        "--name",
        APP_NAME,
        "--workpath",
        str(work_path),
        "--specpath",
        str(spec_path),
        "--distpath",
        str(dist_path),
        "--hidden-import",
        "genetic_algorithm",
        str(ENTRY_SCRIPT),
    ]
    run(command)
    for source_path, target_name in DEMO_ASSETS:
        if source_path.exists():
            shutil.copy2(source_path, DIST_APP_DIR / target_name)
            print(f"已加入示範資料：{target_name}")
        else:
            print(f"找不到示範資料，略過：{source_path}")
    published_dir = publish_portable_copy(DIST_APP_DIR)
    print(f"可攜版已產生：{DIST_APP_DIR}")
    print(f"專案內新版可攜版副本：{published_dir}")


def publish_portable_copy(source_dir: Path) -> Path:
    if not source_dir.exists():
        raise FileNotFoundError(f"找不到可攜版來源資料夾：{source_dir}")
    target_dir = PUBLISHED_APP_DIR
    if target_dir.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target_dir = target_dir.with_name(f"{target_dir.name}_{timestamp}")
    target_dir.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source_dir, target_dir)
    return target_dir


def next_output_path(path: Path) -> Path:
    if not path.exists():
        return path
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_{timestamp}{path.suffix}")


def build_inno_installer() -> None:
    write_inno_script()
    compiler = find_inno_compiler()
    if compiler is None:
        print("未偵測到 Inno Setup，因此先只產生可攜版。")
        print("若要產生一鍵安裝包，請安裝 Inno Setup 6 後重新執行本腳本。")
        build_iexpress_installer()
        return
    run([str(compiler), str(INNO_SCRIPT)])
    print(f"安裝包輸出位置：{INSTALLER_DIR / f'{APP_NAME}_Setup.exe'}")


def build_iexpress_installer() -> None:
    iexpress = find_iexpress()
    if iexpress is None:
        print("未偵測到 IExpress，無法產生自解壓備用安裝包。")
        return
    if not DIST_APP_DIR.exists():
        raise FileNotFoundError(f"找不到可攜版資料夾：{DIST_APP_DIR}")

    INSTALLER_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_root = Path(tempfile.gettempdir()) / f"{APP_NAME}_iexpress_{timestamp}"
    temp_payload_dir = temp_root / "payload"
    temp_sed = temp_root / f"{APP_NAME}.sed"
    temp_setup = temp_root / IEXPRESS_SETUP.name
    temp_payload_dir.mkdir(parents=True)
    temp_zip_base = temp_root / PORTABLE_ZIP.stem
    temp_zip = Path(
        shutil.make_archive(
            str(temp_zip_base),
            "zip",
            root_dir=DIST_APP_DIR.parent,
            base_dir=APP_NAME,
        )
    )
    shutil.copy2(temp_zip, temp_payload_dir / PORTABLE_ZIP.name)

    install_script = f"""@echo off
setlocal
set "TARGET_DIR=%LOCALAPPDATA%\\{APP_NAME}"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -LiteralPath '%~dp0{PORTABLE_ZIP.name}' -DestinationPath $env:LOCALAPPDATA -Force"
if errorlevel 1 (
    echo Failed to extract {APP_NAME}.
    pause
    exit /b 1
)
start "" "%TARGET_DIR%\\{APP_NAME}.exe"
exit /b 0
"""
    install_bat = temp_payload_dir / "install_and_run.bat"
    install_bat.write_text(install_script, encoding="utf-8")

    sed = f"""[Version]
Class=IEXPRESS
SEDVersion=3

[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=Factory GA Scheduler has been extracted and launched.
TargetName={temp_setup}
FriendlyName=Factory GA Scheduler
AppLaunched=install_and_run.bat
PostInstallCmd=<None>
AdminQuietInstCmd=install_and_run.bat
UserQuietInstCmd=install_and_run.bat
SourceFiles=SourceFiles

[SourceFiles]
SourceFiles0={temp_payload_dir}

[SourceFiles0]
%FILE0%=
%FILE1%=

[Strings]
FILE0={PORTABLE_ZIP.name}
FILE1=install_and_run.bat
"""
    temp_sed.write_text(sed, encoding="mbcs")
    sed_output = next_output_path(IEXPRESS_SCRIPT.with_name(f"{APP_NAME}_iexpress_latest.sed"))
    shutil.copy2(temp_sed, sed_output)
    run([str(iexpress), "/N", "/Q", str(temp_sed)])
    if temp_setup.exists():
        zip_output = next_output_path(PORTABLE_ZIP.with_name(f"{APP_NAME}_portable_latest.zip"))
        setup_output = next_output_path(IEXPRESS_SETUP.with_name(f"{APP_NAME}_SelfExtracting_Setup_latest.exe"))
        shutil.copy2(temp_zip, zip_output)
        shutil.copy2(temp_setup, setup_output)
        print(f"可攜版 zip 輸出位置：{zip_output}")
        print(f"自解壓備用安裝包輸出位置：{setup_output}")
    else:
        raise RuntimeError("IExpress 未產生預期的自解壓安裝包。")


def main() -> None:
    try:
        print("開始建立工廠訂單排程 GA 工具發行版")
        build_pyinstaller_app()
        build_inno_installer()
        print("完成。")
    except Exception as exc:
        print(f"打包中止：{exc}")
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
