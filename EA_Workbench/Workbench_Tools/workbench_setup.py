"""EA Repository Workbench - setup tool.

One place for the jobs that need the Python tools: check this computer, create a new model,
install a model into Excel, rebuild a model's Power BI project, and run the self-tests.

    python workbench_setup.py          window (falls back to the menu if no display is available)
    python workbench_setup.py --menu   menu in the console

Needs Python 3.9 or later. openpyxl is needed to build models; pytest only for the self-tests.
The tool can install both with pip. Everything else (tkinter, subprocess) comes with Python.
"""
import importlib
import os
import platform
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

TOOLS = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS))

APP = "EA Workbench Setup"
VERSION = "1.1"
IS_WINDOWS = platform.system() == "Windows"


# ================================================================ helpers
def find_template_folder():
    """The folder whose src, Guides and README a new model copies (the sample model, or tools/model_template)."""
    for cand in (TOOLS / "model_template", TOOLS.parent / "EA_Model_Sample"):
        if (cand / "src" / "modCore.bas").exists():
            return cand
    return None


def has_module(name):
    try:
        importlib.import_module(name)
        return True
    except ImportError:
        return False


def folder_name_for(model_name):
    clean = "".join(ch if ch.isalnum() else "_" for ch in model_name.strip())
    while "__" in clean:
        clean = clean.replace("__", "_")
    clean = clean.strip("_")
    return "EA_Model_" + clean if clean else ""


def windows_path(p: Path):
    s = str(p)
    if IS_WINDOWS:
        return s.rstrip("\\") + "\\"
    return s.replace("/", "\\").rstrip("\\") + "\\"


def drawio_exe():
    if not IS_WINDOWS:
        return None
    for c in (Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "draw.io" / "draw.io.exe",
              Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "draw.io" / "draw.io.exe",
              Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "draw.io" / "draw.io.exe"):
        if c.exists():
            return c
    return None


def excel_installed():
    if not IS_WINDOWS:
        return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application"))
        return True
    except OSError:
        return False


def run_powershell(script, log):
    """Runs a PowerShell script and passes its output to log; returns the exit code."""
    fd, path = tempfile.mkstemp(suffix=".ps1")
    with os.fdopen(fd, "w", encoding="utf-8-sig") as f:
        f.write(script)
    try:
        proc = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", path],
                              capture_output=True, text=True)
        for line in (proc.stdout + proc.stderr).splitlines():
            if line.strip():
                log("  " + line.rstrip())
        return proc.returncode
    finally:
        os.remove(path)


def ps_quote(s):
    return "'" + str(s).replace("'", "''") + "'"


# ================================================================ tasks
# Every task takes log (a function that prints one line) and returns True on success.

def task_check(log):
    ok = True
    log(f"Python {platform.python_version()} on {platform.system()} {platform.release()}")
    if sys.version_info < (3, 9):
        log("  Python 3.9 or later is needed.")
        ok = False
    for mod, why in (("openpyxl", "building models"), ("pytest", "the self-tests (optional)")):
        log(f"  {mod}: {'installed' if has_module(mod) else 'MISSING - needed for ' + why}")
        if mod == "openpyxl" and not has_module(mod):
            ok = False
    tpl = find_template_folder()
    log(f"  Workbench code (src): {tpl / 'src' if tpl else 'NOT FOUND - keep Workbench_Tools next to EA_Model_Sample'}")
    if not tpl:
        ok = False
    if IS_WINDOWS:
        log(f"  Excel: {'found' if excel_installed() else 'not found'}")
        exe = drawio_exe()
        log(f"  draw.io Desktop: {exe if exe else 'not found (optional, needed for diagrams)'}")
    else:
        log("  Not Windows: models can be built here, but Excel setup runs on Windows only.")
    log("Ready." if ok else "Fix the items marked MISSING or NOT FOUND, then check again.")
    return ok


def task_install_packages(log):
    log("Installing openpyxl and pytest with pip (for this user)...")
    proc = subprocess.run([sys.executable, "-m", "pip", "install", "--user", "openpyxl", "pytest"],
                          capture_output=True, text=True)
    for line in (proc.stdout + proc.stderr).splitlines()[-8:]:
        log("  " + line)
    importlib.invalidate_caches()
    if proc.returncode == 0:
        log("Done. Restart this tool if a check still shows a package as missing.")
        return True
    log("pip did not finish. If you are offline or pip is blocked, ask IT to install openpyxl.")
    return False


def task_create_model(log, parent, model_name, with_sample):
    if not has_module("openpyxl"):
        log("openpyxl is not installed: use 'Install Python packages' first.")
        return False
    tpl = find_template_folder()
    if not tpl:
        log("The Workbench code (src) was not found next to the tools.")
        return False
    name = folder_name_for(model_name)
    if not name:
        log("Enter a model name, for example Group Landscape.")
        return False
    parent = Path(parent).expanduser()
    if not parent.is_dir():
        log(f"The folder {parent} does not exist.")
        return False
    folder = parent / name
    if folder.exists():
        log(f"{folder} already exists. Pick another name.")
        return False
    import builder
    import pbip
    log(f"Creating {folder} ...")
    log("  " + builder.build(folder, with_sample))
    pbip.write(folder, windows_path(folder))
    log("  Power BI project written to Reports\\")
    shutil.copytree(tpl / "src", folder / "src")
    log("  Workbench code copied to src\\")
    for extra in ("Guides",):
        if (tpl / extra).is_dir():
            shutil.copytree(tpl / extra, folder / extra)
    for f in ["README.md"] + [f"TEST_Phase{i}.md" for i in range(1, 6)]:
        if (tpl / f).exists():
            shutil.copy2(tpl / f, folder / f)
    log(f"Model created: {folder}")
    log("Next: Install into Excel (step 3), with this folder.")
    return True


def task_unblock(log, folder):
    folder = Path(folder)
    if not IS_WINDOWS:
        log("Only needed on Windows.")
        return False
    if not folder.is_dir():
        log(f"{folder} is not a folder.")
        return False
    log(f"Removing the 'downloaded from the internet' mark from the files in {folder} ...")
    code = run_powershell(f"Get-ChildItem -LiteralPath {ps_quote(folder)} -Recurse -File | Unblock-File\n"
                          "Write-Output 'Files unblocked.'", log)
    return code == 0


INSTALL_PS = r"""
$ErrorActionPreference = 'Stop'
$folder = __FOLDER__
$core = Join-Path $folder 'EA_Core.xlsx'
if (-not (Test-Path -LiteralPath $core)) {
    if (Test-Path -LiteralPath (Join-Path $folder 'EA_Core.xlsm')) { Write-Output 'ALREADY: this model is installed already (EA_Core.xlsm exists).'; exit 3 }
    Write-Output "MISSING: no EA_Core.xlsx in $folder"; exit 4
}
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$wb = $xl.Workbooks.Open($core)
try { $null = $wb.VBProject.VBComponents.Count }
catch {
    Write-Output 'TRUST: Excel is blocking access to the VBA project.'
    Write-Output 'In Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings >'
    Write-Output 'tick "Trust access to the VBA project object model", click OK, then run Install into Excel again.'
    $wb.Close($false); $xl.Quit(); exit 2
}
$null = $wb.VBProject.VBComponents.Import((Join-Path $folder 'src\modInstall.bas'))
Write-Output 'Installer imported; Excel now asks you to confirm the install.'
$xl.Run('InstallWorkbench')
Write-Output 'Finished. Excel stays open on the installed model.'
exit 0
"""


def task_install_excel(log, folder):
    folder = Path(folder)
    if not IS_WINDOWS:
        log("Installing into Excel runs on Windows only.")
        return False
    if not (folder / "src" / "modInstall.bas").exists():
        log(f"{folder} is not a model folder (no src\\modInstall.bas).")
        return False
    log(f"Opening Excel to install {folder.name} ...")
    log("  Keep this window open; Excel asks you to confirm, then reports when it is done.")
    code = run_powershell(INSTALL_PS.replace("__FOLDER__", ps_quote(folder)), log)
    if code == 0:
        log("Installed. Tip: add the folder that holds your models as an Excel Trusted Location.")
    return code == 0


def task_rebuild_pbip(log, folder):
    folder = Path(folder)
    if not (folder / "model.json").exists():
        log(f"{folder} is not a model folder (no model.json).")
        return False
    import pbip
    pbip.write(folder, windows_path(folder))
    log(f"Power BI project rewritten with ModelFolder = {windows_path(folder)}")
    log("Report pages you laid out are in Reports\\EA_Reports.Report and are replaced too; save a copy first if needed.")
    return True


def task_self_test(log):
    if not (has_module("pytest") and has_module("openpyxl")):
        log("pytest and openpyxl are needed: use 'Install Python packages' first.")
        return False
    tpl = find_template_folder()
    tmp = Path(tempfile.mkdtemp(prefix="eawb_test_"))
    try:
        import builder
        import pbip
        model = tmp / "EA_Model_Sample"
        log("Building a fresh sample model to test ...")
        builder.build(model, True)
        pbip.write(model, windows_path(model))
        if tpl:
            shutil.copytree(tpl / "src", model / "src")
        env = dict(os.environ, EAWB_MODEL=str(model))
        proc = subprocess.run([sys.executable, "-m", "pytest", "-q", str(TOOLS / "test_workbench.py")],
                              capture_output=True, text=True, env=env, cwd=str(TOOLS))
        lines = (proc.stdout + proc.stderr).strip().splitlines()
        for line in [l for l in lines if "FAILED" in l or "Error" in l][:20] + lines[-1:]:
            log("  " + line)
        log("All tests passed." if proc.returncode == 0 else "Some tests failed - see above.")
        return proc.returncode == 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def open_folder(path):
    path = str(path)
    if IS_WINDOWS:
        os.startfile(path)  # noqa: S606 - opens Explorer on a folder the user chose
    elif platform.system() == "Darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


def default_parent():
    tpl = find_template_folder()
    return str(tpl.parent if tpl else Path.home())


def default_model():
    tpl = find_template_folder()
    return str(tpl) if tpl and tpl.name.startswith("EA_Model_") else ""


# ================================================================ menu
def run_menu():
    def log(line):
        print(line)

    def ask(prompt, default=""):
        s = input(f"{prompt}{f' [{default}]' if default else ''}: ").strip()
        return s or default

    def yes(prompt, default=True):
        s = input(f"{prompt} [{'Y/n' if default else 'y/N'}]: ").strip().lower()
        return default if not s else s.startswith("y")

    items = [
        ("Check this computer", lambda: task_check(log)),
        ("Install Python packages (openpyxl, pytest)", lambda: task_install_packages(log)),
        ("Create a new model", lambda: task_create_model(
            log, ask("Folder to create it in", default_parent()), ask("Model name"),
            yes("Include the sample data?", False))),
        ("Unblock downloaded files in a model folder", lambda: task_unblock(log, ask("Model folder", default_model()))),
        ("Install a model into Excel", lambda: task_install_excel(log, ask("Model folder", default_model()))),
        ("Rebuild a model's Power BI project", lambda: task_rebuild_pbip(log, ask("Model folder", default_model()))),
        ("Run the self-tests", lambda: task_self_test(log)),
        ("Open a folder", lambda: open_folder(ask("Folder", default_parent()))),
    ]
    while True:
        print(f"\n{APP} {VERSION}\n" + "-" * (len(APP) + 4))
        for i, (label, _) in enumerate(items, 1):
            print(f"  {i}. {label}")
        print("  0. Exit")
        try:
            choice = input("Choose: ").strip()
        except EOFError:
            return
        if choice in ("0", "q", "exit", ""):
            return
        if choice.isdigit() and 1 <= int(choice) <= len(items):
            print()
            try:
                items[int(choice) - 1][1]()
            except Exception as e:  # keep the menu alive and say what went wrong
                print(f"Stopped: {e}")
            try:
                input("\nPress Enter to go back to the menu...")
            except EOFError:
                return
        else:
            print("Type a number from the list.")


# ================================================================ window
def run_window():
    import tkinter as tk
    from tkinter import filedialog, ttk

    root = tk.Tk()
    root.title(f"{APP} {VERSION}")
    root.geometry("820x720")
    root.minsize(700, 600)
    msgs = queue.Queue()
    busy = tk.BooleanVar(value=False)
    buttons = []

    def log(line):
        msgs.put(line)

    def pump():
        while not msgs.empty():
            out.configure(state="normal")
            out.insert("end", msgs.get() + "\n")
            out.see("end")
            out.configure(state="disabled")
        root.after(100, pump)

    def run(fn, *args):
        if busy.get():
            return

        def work():
            try:
                fn(log, *args)
            except Exception as e:
                log(f"Stopped: {e}")
            finally:
                root.after(0, done)

        def done():
            busy.set(False)
            for b in buttons:
                b.configure(state="normal")
            status.set("Ready")

        busy.set(True)
        for b in buttons:
            b.configure(state="disabled")
        status.set("Working...")
        log("")
        threading.Thread(target=work, daemon=True).start()

    def browse(var):
        d = filedialog.askdirectory(initialdir=var.get() or default_parent())
        if d:
            var.set(d)

    def button(parent, text, cmd):
        b = ttk.Button(parent, text=text, command=cmd)
        buttons.append(b)
        return b

    pad = {"padx": 8, "pady": 4}
    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)
    ttk.Label(frm, text="EA Repository Workbench setup", font=("Segoe UI", 14, "bold")).pack(anchor="w")
    ttk.Label(frm, text="Steps 1 to 3 set up a model; step 4 maintains one.", foreground="#595959").pack(anchor="w")

    s1 = ttk.LabelFrame(frm, text="1. Check this computer", padding=6)
    s1.pack(fill="x", pady=6)
    button(s1, "Check", lambda: run(task_check)).pack(side="left", **pad)
    button(s1, "Install Python packages", lambda: run(task_install_packages)).pack(side="left", **pad)
    button(s1, "Run self-tests", lambda: run(task_self_test)).pack(side="left", **pad)

    s2 = ttk.LabelFrame(frm, text="2. Create a new model", padding=6)
    s2.pack(fill="x", pady=6)
    parent_var = tk.StringVar(value=default_parent())
    name_var = tk.StringVar(value="")
    sample_var = tk.BooleanVar(value=False)
    ttk.Label(s2, text="Create in folder").grid(row=0, column=0, sticky="w", **pad)
    ttk.Entry(s2, textvariable=parent_var, width=60).grid(row=0, column=1, sticky="we", **pad)
    button(s2, "Browse...", lambda: browse(parent_var)).grid(row=0, column=2, **pad)
    ttk.Label(s2, text="Model name").grid(row=1, column=0, sticky="w", **pad)
    ttk.Entry(s2, textvariable=name_var, width=60).grid(row=1, column=1, sticky="we", **pad)
    preview = tk.StringVar(value="")
    ttk.Label(s2, textvariable=preview, foreground="#595959").grid(row=2, column=1, sticky="w", padx=8)
    name_var.trace_add("write", lambda *_: preview.set(
        f"Folder: {folder_name_for(name_var.get())}" if folder_name_for(name_var.get()) else ""))
    ttk.Checkbutton(s2, text="Include the sample data (14 insurance entities)", variable=sample_var).grid(
        row=3, column=1, sticky="w", **pad)

    def created_then_select():
        def go(log_):
            new = str(Path(parent_var.get()) / folder_name_for(name_var.get()))
            if task_create_model(log_, parent_var.get(), name_var.get(), sample_var.get()):
                root.after(0, lambda: model_var.set(new))     # Tk variables are set on the window's thread
        run(lambda log_: go(log_))
    button(s2, "Create model", created_then_select).grid(row=3, column=2, **pad)
    s2.columnconfigure(1, weight=1)

    s3 = ttk.LabelFrame(frm, text="3. Install a model into Excel (Windows)", padding=6)
    s3.pack(fill="x", pady=6)
    model_var = tk.StringVar(value=default_model())
    ttk.Label(s3, text="Model folder").grid(row=0, column=0, sticky="w", **pad)
    ttk.Entry(s3, textvariable=model_var, width=60).grid(row=0, column=1, sticky="we", **pad)
    button(s3, "Browse...", lambda: browse(model_var)).grid(row=0, column=2, **pad)
    row = ttk.Frame(s3)
    row.grid(row=1, column=0, columnspan=3, sticky="w")
    button(row, "Unblock downloaded files", lambda: run(task_unblock, model_var.get())).pack(side="left", **pad)
    button(row, "Install into Excel", lambda: run(task_install_excel, model_var.get())).pack(side="left", **pad)
    button(row, "Open folder", lambda: model_var.get() and open_folder(model_var.get())).pack(side="left", **pad)
    ttk.Label(s3, text="Excel must allow 'Trust access to the VBA project object model' while installing "
                       "(File > Options > Trust Center > Macro Settings).", foreground="#595959",
              wraplength=740).grid(row=2, column=0, columnspan=3, sticky="w", padx=8)
    s3.columnconfigure(1, weight=1)

    s4 = ttk.LabelFrame(frm, text="4. Maintain a model (uses the model folder above)", padding=6)
    s4.pack(fill="x", pady=6)
    button(s4, "Rebuild Power BI project", lambda: run(task_rebuild_pbip, model_var.get())).pack(side="left", **pad)
    ttk.Label(s4, text="After moving a model folder, so Power BI points at its new place.",
              foreground="#595959").pack(side="left", padx=8)

    out = tk.Text(frm, height=14, wrap="word", state="disabled", font=("Consolas", 9))
    out.pack(fill="both", expand=True, pady=(8, 0))
    status = tk.StringVar(value="Ready")
    ttk.Label(frm, textvariable=status, foreground="#595959").pack(anchor="w")
    root.after(100, pump)
    run(task_check)
    root.mainloop()


def main():
    if "--menu" in sys.argv or "--cli" in sys.argv:
        run_menu()
        return
    try:
        run_window()
    except Exception as e:  # no tkinter or no display: use the console menu instead
        print(f"(No window available: {e}. Using the menu.)")
        run_menu()


if __name__ == "__main__":
    main()
