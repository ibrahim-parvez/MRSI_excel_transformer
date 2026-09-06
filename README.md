<div align="center">
  <img src="assets/images/mrsi_logo.png" alt="MRSI" width="140">

  # MRSI Data Normalization Tool

  **Isotope-ratio mass spectrometer data, normalized and reported — without touching a spreadsheet formula.**

  Built for the McMaster Research Group for Stable Isotopologues

  `v2.1.0` · Windows & macOS · Python 3.13 · PyQt6

  [**Download the installer →**](https://github.com/ibrahim-parvez/MRSI_Data_Normalization_Tool/releases/tag/Installer)
</div>

---

## What this is

Raw output from an IRMS run is a wall of numbers. Turning it into publishable δ¹³C and δ¹⁸O values means sorting runs into groups, averaging the right replicates, flagging outliers, fitting a slope and intercept against your reference materials, normalizing every sample against that fit, and laying the whole thing out so the next person can follow it. Done by hand in Excel it takes an afternoon per run, and every manual step is a chance to paste into the wrong column.

The DNT does all of it in one click, and writes the result back into the **same workbook** as a set of new sheets — so the raw data, every intermediate stage, and the final report all live together and stay auditable.

It handles **Carbonate** and **Water** in dedicated pipelines, with optional **CO₂** and **Yield** calculations layered on for carbonate work.

---

## Highlights

| | |
|---|---|
| **Two full pipelines** | Carbonate and Water, each with its own reference materials, normalization maths, and report layout |
| **Run all 7 steps or just one** | Tick exactly the steps you need — re-run only the normalization after changing a setting, for example |
| **Batch + Combine** | Point it at a folder's worth of raw files, process them all, and merge every standard into one comparison workbook with charts |
| **Outlier control** | 1σ / 2σ / 3σ thresholds, and a choice between discarding a whole run or keeping the good half of it |
| **Editable reference materials** | Add, remove, recolor and regroup your standards — no code changes |
| **Your originals stay safe** | Automatic backups before every run, crash-safe saves, and a temp-copy mode for batch jobs |
| **Drag and drop** | Drop an `.xlsx` (or a legacy `.xls`, converted automatically) anywhere on the window |
| **Dark mode** | `Ctrl/Cmd + D` |

---

## The processing pipeline

Every step reads the sheet the previous one produced and appends a new `_DNT` sheet. Nothing is overwritten, so you can open any stage and see exactly what happened.

| Step | Creates | What it does |
|:---:|---|---|
| **1** | `Data_DNT` | Reads the raw export, finds the header row wherever it sits, groups replicate runs, computes per-group C/O averages and standard deviations, flags outliers |
| **2** | `To Sort_DNT` | Converts formulas to cached values and applies the run filter (Last 6, Ref Avg, Delta, …) |
| **3** | `Last 6_DNT` | Keeps the analytically meaningful replicates — optionally with statistical outliers already removed |
| **4** | `Pre-Group_DNT` | Stages rows for grouping and copies formatting across |
| **5** | `Group_DNT` | Sorts reference materials and samples into blocks, marks outliers with red strikethrough, writes per-block Average / Stdev / Count |
| **6** | `Normalization_DNT` | The core: fits slope and intercept from your standards, normalizes every sample to VPDB and VSMOW, builds the blue reference box, and adds the CO₂ and Yield tables when enabled |
| **7** | `Report_DNT` | The clean, presentable summary sheet |

After step 7 the tool re-runs step 6 as a **"Finalizing Formatting"** pass, so the workbook opens on `Normalization_DNT` with everything laid out.

### Carbonate extras

- **CO₂ table** — a parallel normalization against CO₂ reference values, with a Quick Switch for the 25 °C and 72 °C acid-fractionation profiles, or a **Custom** profile you can name yourself (there's a `°` insert button so you can type `70 ° C` without hunting for the character).
- **Yield table** — theoretical vs. calculated yield per sample, driven by a compound you choose for reference materials and a second one for samples. The two are color-coded (green for reference, pink for sample) all the way through the Theoretical column so you can see at a glance which is which.

---

## Combine Data

For comparing standards across many runs.

1. Drop in as many raw files as you like.
2. For each file, drag identifiers between the **References** and **Samples** lists — per-file overrides, so a material treated as a standard in one run can be a sample in another.
3. Choose whether to work on **temporary copies** (originals untouched) or **the original files**.
4. Choose **Create New Combined File** or **Append to an Existing** one.

Each file is taken through all seven steps, then every reference material gets its own sheet in the combined workbook, with each run's block stacked in time order and native Excel scatter charts plus matplotlib trend plots drawn underneath.

---

## Advanced Settings

Press `Ctrl/Cmd + Shift + S` (or long-press the menu button) and enter the password. Two extra tabs appear: **Combine Data** and **Advanced Settings**.

**Outliers**
- **Sigma threshold** — 1σ, 2σ or 3σ. Anything outside mean ± (σ × stdev) is flagged.
- **Exclusion logic** — *Individual* keeps a run's good δ¹⁸O even when its δ¹³C is an outlier; *Exclude Entire Row* discards the whole measurement if either value fails.

**Calculation modes**
- **Step 3** — `Last 6` or `Last 6 Outliers Excluded`.
- **Step 7** — report from `All Values` or `Outliers Excluded`.

**Standard deviation threshold** — optional; highlights any group whose stdev exceeds your limit.

**Reference materials** — one editable table per type (Carbonate / Water / CO₂): name, published δ values, and display color. **Slope & Intercept Groups** define which standards are fitted together, and you can have several groups side by side.

**Yield** — atomic weights and compound definitions, plus the reference/sample compound selection.

> Settings live in memory for the session and reset to defaults on relaunch. Use **Export / Import** to keep a configuration, and **Manage Settings** for the rolling auto-saves taken as you move between tabs.

---

## Keyboard shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl/Cmd + 1…4` | Jump to a tab |
| `Ctrl/Cmd + R` | Run the selected steps |
| `Ctrl/Cmd + D` | Toggle dark mode |
| `Ctrl/Cmd + Shift + S` | Unlock / re-lock the advanced tabs |

---

## Your data is protected

Processing writes into your workbook, so the tool takes that seriously:

- **A backup before every run.** Snapshots are kept for the session and reachable from the History button, so you can view or restore any earlier state.
- **Crash-safe saves.** Every sheet is written to a temporary file and swapped into place atomically. If the disk fills, a network drive drops, or the machine dies mid-save, your workbook is either the old version or the new one — never a half-written file.
- **Read-only files are respected** rather than silently overwritten.
- **Temp-copy mode** in Combine leaves your raw files byte-for-byte untouched.
- **Steps fail loudly.** If a prerequisite sheet is missing, you get a message naming the sheet and the step to run first — never a green "completed" on a step that did nothing.

---

## Installing

### For end users

Grab the installer from the [releases page](https://github.com/ibrahim-parvez/MRSI_Data_Normalization_Tool/releases/tag/Installer). It ships as a standalone `.exe` / `.app` — no Python needed. The app checks GitHub for new versions and can update itself in place.

**Microsoft Excel must be installed.** The tool drives a hidden Excel instance to recalculate formulas between steps, which Excel alone can do correctly.

### From source

```bash
git clone https://github.com/ibrahim-parvez/MRSI_Data_Normalization_Tool.git
cd MRSI_Data_Normalization_Tool

python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

pip install -r requirements.txt

python src/main.py
```

On macOS you'll be asked to allow the app to control Microsoft Excel the first time it recalculates — this is xlwings driving Excel via AppleScript, and it needs to be granted for processing to work.

---

## Building

Everything the build needs lives in [`packaging/`](packaging/) — a `.spec` file
per target, plus a script that runs them the same way on both platforms.

```bash
python packaging/build.py app          # the tool itself
python packaging/build.py installer    # the standalone installer
python packaging/build.py all
python packaging/build.py clean        # remove build/, dist/ and __pycache__
```

Output lands in `dist/`. The spec picks the right shape for the host platform:
a single self-contained `.exe` with a startup splash on Windows, and a `.app`
bundle on macOS — the layout the in-app updater expects when it swaps a new
version into place. Cross-compiling isn't possible; build each platform on
that platform.

`assets/images` is copied into the build by the spec and read back at runtime
through `utils.resources`, which resolves paths against `sys._MEIPASS` when
frozen. That's what makes the logo appear in a packaged build.

---

## Project layout

```
assets/images/                     Logos and platform icons (single source of truth)
packaging/                         Build definitions
├── build.py                       Build/clean entry point
├── app.spec                       PyInstaller spec for the tool
├── installer.spec                 PyInstaller spec for the installer
└── entitlements.plist             macOS codesigning entitlements
tools/
└── generate_password_hash.py      Regenerate the Advanced Settings credentials
src/
├── main.py                        Entry point: splash, then staged startup
├── gui/
│   ├── splash.py                  Startup splash and update check
│   ├── main_window.py             Main window, tabs, theming, history, updates
│   └── tabs/
│       ├── carbonate_tab.py       Carbonate run configuration
│       ├── water_tab.py           Water run configuration
│       ├── combine_tab.py         Batch file list and combine options
│       ├── advanced_settings_tab.py
│       └── combine_processors/    Combine workers (one per mode)
├── processors/
│   ├── carbonate/step1…step7      Carbonate pipeline
│   └── water/step1…step7          Water pipeline
└── utils/
    ├── resources.py               Asset lookup, in a checkout and when frozen
    ├── settings.py                Runtime configuration + import/export
    ├── excel_engine.py            Shared background Excel instance
    ├── common_utils.py            Atomic saves, embedded settings comments
    ├── calculators/               Isotope fractionation maths
    ├── updater.py                 GitHub release checks and self-update
    └── installer/                 Standalone installer app
```

---

## Troubleshooting

**"Sheet 'X_DNT' not found. Run Step N first."**
Steps depend on the sheet the previous one creates. Either tick the earlier step or run the full chain.

**"Required column(s) missing from sheet …"**
The raw export doesn't have a column the pipeline groups on (`Identifier 1` and `Time Code` for carbonate, `Line` for water). The message lists what it did find.

**"… is a sheet this tool generates, not raw data."**
The sheet name box is pointing at a `_DNT` sheet. Pick the original instrument sheet — the tool auto-detects the right-most sheet when you drop a file.

**Excel windows flashing during processing**
Expected. Each step opens the workbook to recalculate formulas, then closes it. On macOS you may see Excel bounce in the Dock.

**`The RPC server is unavailable` (Windows)**
The background Excel instance was killed mid-run. The tool now restarts Excel and retries automatically; if you still see it, check whether antivirus or a cleanup tool is terminating `EXCEL.EXE`.

**Nothing happens when I click Remove Row**
Select a row in the table first — the button stays greyed out until there's a selection.

---

<div align="center">
  <sub>Built by Ibrahim Parvez for the McMaster Research Group for Stable Isotopologues</sub>
</div>
