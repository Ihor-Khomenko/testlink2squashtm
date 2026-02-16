# TestLink to Squash TM Converter

Converts TestLink XML exports to Squash TM XLS import files.

## Installation

Requires Python 3.7+ and dependencies:

```bash
pip install pandas openpyxl xlwt
```

## Parameters

Edit these at the top of `tl2squash.py`:

| Parameter | Description |
|-----------|-------------|
| `PROJECT_NAME` | Squash TM project name (must match exactly) |
| `IMPORT_ROOT_FOLDER` | Root folder for import (leave `""` for TestLink structure) |
| `SQUASH_CUF_CODE` | Custom field code for execution type (e.g. `EXEC_TYPE`) |
| `SQUASH_CUF_TESTLINK_ID` | Custom field code for TestLink ID (e.g. `testlink_id`) |
| `MIGRATE_TESTLINK_ID` | Include TestLink IDs for cross-referencing |
| `SPLIT_INTO_PARTS` | Split large imports (1 = single file, 4 = 4 parts) |

## Usage

```bash
python tl2squash.py [path/to/testlink.xml]
```

## Reference

- [Squash TM – Import test cases](https://tm-en.doc.squashtest.com/v7/user-guide/manage-test-cases/import-test-cases.html)
