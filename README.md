# TestLink to Squash TM – Test Case Import Converter

This script converts a TestLink test case export (XML) into a Squash TM import file. The output is an **XLS** file (Excel 97–2003) in the same format as the [Squash TM test case import template](https://tm-en.doc.squashtest.com/v7/user-guide/manage-test-cases/import-test-cases.html), so you can import it directly in Squash TM.

## What it does

- Reads a TestLink **XML export** (full project / test specification).
- Preserves **folder structure**: project → folders (test suites) → test cases.
- Converts **test case names**, **descriptions**, **prerequisites**, and **steps** (actions and expected results).
- Maps TestLink **execution type** (Manual/Automated) into a Squash TM **custom field** (e.g. `EXEC_TYPE`).
- Writes **.xls** files (like Squash TM template) with sheets: TEST_CASES, STEPS, PARAMETERS, DATASETS, LINK_REQ_TC.
- Can split large imports into multiple files to avoid 413 "Request Entity Too Large" errors.

## Requirements

- Python 3.7+
- Dependencies: `pandas`, `openpyxl`, `xlwt`

## Installation

From the project directory:

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install pandas openpyxl xlwt
```

## Usage

### 1. Export from TestLink

In TestLink, export your test project as **XML** (test specification / full export). Save the file (e.g. `MyProject.testproject-deep.xml`).

### 2. Configure the script

Edit the configuration block at the top of `tl2squash.py`:

| Variable | Description |
|----------|-------------|
| `PROJECT_NAME` | Squash TM project name. Must match exactly (case-sensitive). |
| `IMPORT_ROOT_FOLDER` | Leave `""` to get the same structure as TestLink (project → folders → test cases). Set to e.g. `"TestLink_Import_20260213"` to import under a dedicated root folder (avoids conflicts if the project already has content). |
| `SQUASH_CUF_CODE` | Code of the Squash TM custom field for execution type (e.g. `EXEC_TYPE`). Values written: "Manual" / "Automated". |
| `SQUASH_CUF_TESTLINK_ID` | Code of the Squash TM custom field for TestLink ID (e.g. `testlink_id`). Leave empty to use TC_REFERENCE instead. |
| `MIGRATE_TESTLINK_ID` | `True` to include TestLink external ID (e.g. 100-855) for referencing between test cases. `False` to leave it empty. |
| `SPLIT_INTO_PARTS` | Split large imports into multiple files to avoid 413 "Request Entity Too Large" errors. `1` = single file (default), `4` = split into 4 equal parts, etc. |

### 3. Run the script

**Option A – XML in current directory**

Place your `.xml` file in the script directory and run:

```bash
python tl2squash.py
```

The script uses the first `.xml` file it finds.

**Option B – Specify XML path**

```bash
python tl2squash.py /path/to/MyProject.testproject-deep.xml
```

The output file is written next to the XML (or in the current directory if the path has no directory).

### 4. Output

- **Success with XLS:**  
  `SUCCESS! File created (XLS, like Squash TM template): <path>/<name>_SquashTM_Import.xls`  
  If `SPLIT_INTO_PARTS > 1`, creates multiple files: `<name>_SquashTM_Import_part1.xls`, etc.

- If `xlwt` is missing or .xls writing fails, the script falls back to **.xlsx** and prints a short message.

### 5. Import in Squash TM

1. In Squash TM, open the **Test cases** workspace and select the project (same name as `PROJECT_NAME`).
2. Use **Import** and choose the generated **.xls** file(s).
3. If split into multiple parts, import each file separately (part1, part2, etc.).
4. Run a **simulation** first to check for warnings.
5. Then run the real **Import**.

**Important:** If you get a **500 error** or **DuplicateNameException** (“… already exists within the same container”), the project already has test case library content from a previous import. Either:

- **Delete all test case library content** under that project in Squash TM and import again, or  
- Set **`IMPORT_ROOT_FOLDER`** to a non-empty value (e.g. `"TestLink_Import_20260213"`) so the import goes into a new subtree.

## Output file structure

- **TEST_CASES:** One row per test case (path, description, prerequisite, nature, type, custom field EXEC_TYPE, etc.).
- **STEPS:** One row per step (action, expected result); order is preserved.
- **PARAMETERS / DATASETS / LINK_REQ_TC:** Filled with headers and one placeholder row so Squash TM accepts the file; no TestLink data is mapped there.

Folder and test case names are taken from the TestLink XML. Unique suffixes are only added to test cases with duplicate names within the same folder. TestLink IDs are preserved in the configured custom field (or TC_REFERENCE) when `MIGRATE_TESTLINK_ID = True`.

## Notes

- **XLS cell limit:** Excel 97–2003 allows at most 32,767 characters per cell. Longer values (e.g. very long descriptions or steps) are truncated in the .xls file. If you need full content, you can force .xlsx output by temporarily removing or not installing `xlwt`.
- **Execution type:** TestLink `execution_type` 1 → "Manual", 2 → "Automated". Ensure your Squash TM custom field (e.g. EXEC_TYPE) uses the same option labels.
- **Encoding:** The script writes UTF-8; the .xls is produced with `xlwt` in UTF-8 mode. Squash TM supports XLS import in UTF-8.

## Reference

- [Squash TM – Import test cases](https://tm-en.doc.squashtest.com/v7/user-guide/manage-test-cases/import-test-cases.html)
