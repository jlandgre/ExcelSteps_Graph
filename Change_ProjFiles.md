## Purpose
Create and validate ProjFiles class for handling directory paths and filenames for a project. Scope is path/filename holder (attributes only) for now; may expand to file-operation methods in a future change.

## Background

This change follows [[VBA Project Changes]] sequencing and is scoped to planning-level architecture before coding.

Project code instances `ProjFiles` as `files` and uses its attributes for wayfinding. `ProjFiles` is decoupled from any specific project — it references the project global `IsTest` (always declared `Public Boolean` in project workbook, default `False`) directly rather than receiving it as an argument.

## Data I/O Descriptions

**Inputs to `.Init`:**
- `files As ProjFiles` — self-reference (standard project convention: instance passed as first arg to its own methods)
- `wkbk As Workbook` — calling project workbook; caller passes `ThisWorkbook`. Required because `ProjFiles` resides in ExcelSteps add-in where `ThisWorkbook` would refer to the add-in, not the project.
- `Optional subdir_tests As String` — subdirectory under `pathTests` used as data path when `IsTest = True`

**Implicit inputs used by `SetGenericPaths`:**
- `wkbk.Path` — used as `pathSrc`
- `IsTest` — project global Boolean; controls `pathData` / `pathColInfo` branching

**Outputs — all attributes set after `.Init`:**

| Attribute | Value |
|---|---|
| `IsDevelopment` | `True` if last path segment of `pathSrc` = `"src"` |
| `subdir_tests` | Value of `Optional subdir_tests` arg (empty string if not passed) |
| `pathSrc` | `wkbk.Path & sep` |
| `pathRoot` | `pathSrc` one level up, trailing sep (meaningful when `IsDevelopment = True`) |
| `pathTests` | `pathRoot & "tests" & sep` (meaningful when `IsDevelopment = True`) |
| `pathData` | trailing-sep path: `pathSrc` if `IsTest = False`; else `pathTests` if `subdir_tests = ""`; else `pathTests & subdir_tests & sep` |
| `pathColInfo` | Same logic as `pathData`; trailing sep |
| `fColInfo` | `"ColInfo.xlsx"` |
| `pfColInfo` | `pathColInfo & fColInfo` (no sep needed; pathColInfo already ends with sep) |

## Project Architecture

**New class: `ProjFiles.cls`**
- Instanced as `files` in project code (e.g. `Set files = New ProjFiles`)
- All attributes are public; no `Property Get/Let`

**Methods:**
- `Public Function Init(files, ByVal wkbk As Workbook, Optional subdir_tests As String) As Boolean` — sets `.subdir_tests` first, then calls `SetGenericPaths(files)`
- `Private Function SetGenericPaths(files) As Boolean` — derives and sets all path/filename attributes from `wkbk.Path` (stored as `.wkbk`) and `IsTest`

**Factory function:**
- `New_ProjFiles` in `Validation.bas` — enables cross-workbook instantiation from test suite

## Test Architecture

Tests housed in existing `Tests_ProjFiles.bas`. New `Procedure` group `procs.ProjFiles` added to `Procedures.cls`.

## Discussion: Environment Detection

`IsDevelopment` is derived automatically — no argument needed. If `pathSrc` ends in `"src"`, the workbook is in the standard development layout (`root/src/`), so `pathRoot` and `pathTests` are meaningful. In non-development (deployed) environments, these attributes still get set but callers should not rely on them; `pathData`/`pathColInfo` default to `pathSrc` when `IsTest = False` regardless of environment.

## Discussion: SetGenericPaths Visibility

`SetGenericPaths` is `Private` — called once from `.Init`. If caller needs different paths (e.g. different `subdir_tests`), they call `.Init` again. No partial re-init risk.

## Testing Considerations

- **Test module**: `tests_ToolboxClasses.bas` (shared with ColInfo and other toolbox classes); new driver sub `TestingDriver_ProjFiles`
- **Procedure wiring**: Add `Public ProjFiles As Object` to `Procedures.cls` declarations; add `Set .ProjFiles = New Procedure` + `.ProjFiles.Name = "ProjFiles"` in `Procedures.Init` under `' tests_ToolboxClasses` group comment. Follow [[vba-testing-create-new-test-procedure]] skill for all wiring steps.
- **Driver**: `TestingDriver_ProjFiles` in `tests_ToolboxClasses.bas`; `procs.Init` args: `ThisWorkbook`, `"ProjFiles"`, `"Tests_ToolboxClasses"`; enable `procs.ProjFiles.Enabled = True`
- **Cross-workbook instantiation**: Add `New_ProjFiles` factory function to `Validation.bas`
- **Test setup**: `Set files = DemoProj.New_ProjFiles` then `tst.Assert tst, files.Init(files, ThisWorkbook)`
- **Coverage**: Unit tests for `Init` (with and without `subdir_tests`) and `SetGenericPaths` path derivation; verify full attribute set after init
- **Key edge cases**: `subdir_tests` empty vs. populated; `IsTest = True` vs. `False`; `IsDevelopment` detection when path ends in `"src"` vs. other folder name
- **Test data**: No external files needed — all inputs derived from workbook path and globals

## Procedure Outline

**ProjFiles Class**
- **`Init`** — store `.wkbk`; set `.subdir_tests` from optional arg; call `SetGenericPaths`
- **`SetGenericPaths`** (Private) — set `pathSrc` from `.wkbk.Path`; detect `IsDevelopment`; derive `pathRoot`, `pathTests`; branch `pathData`/`pathColInfo` on `IsTest` and `subdir_tests`; set `fColInfo`, `pfColInfo`
