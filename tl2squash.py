import xml.etree.ElementTree as ET
import pandas as pd
import os
import re
import html
import random
import string
import sys
from datetime import datetime

# CONFIGURATION
PROJECT_NAME = "OpenProject"
IMPORT_ROOT_FOLDER = ""
SQUASH_CUF_CODE = "EXEC_TYPE"
SQUASH_CUF_TESTLINK_ID = "testlink_id"
MIGRATE_TESTLINK_ID = True
SPLIT_INTO_PARTS = 4

def generate_short_id():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))

def sanitize_text(text):
    if not text: return ""
    text = str(text)
    text = html.unescape(text)
    cleanr = re.compile('<.*?>')
    text = re.sub(cleanr, '', text)
    text = "".join(ch for ch in text if (ord(ch) >= 32 or ch in "\n\t\r"))
    return " ".join(text.split())

def rich_text_to_html(raw):
    if not raw: return ""
    text = str(raw)
    text = html.unescape(text)
    cleanr = re.compile('<.*?>')
    text = re.sub(cleanr, '', text)
    text = "".join(ch for ch in text if (ord(ch) >= 32 or ch in "\n\t\r"))
    text = " ".join(text.split())
    text = html.escape(text)
    return f"<p>{text}</p>" if text else ""

def sanitize_folder_name(name_str):
    if not name_str: return "Unnamed"
    clean = name_str.replace('/', '-').replace('\\', '-')
    return clean.strip()

def build_folder_path(path_prefix, suite_name, seen_folders=None):
    # Empty suite name: do not add a path segment (avoids "Unnamed" for root)
    if not suite_name:
        return path_prefix
    clean_name = sanitize_folder_name(suite_name)
    # Avoid DuplicateNameException: same folder name under same parent must be unique in Squash
    if seen_folders is not None:
        key = (path_prefix, clean_name)
        if key in seen_folders:
            clean_name = f"{clean_name}_{generate_short_id()}"
        else:
            seen_folders.add(key)
    if not path_prefix:
        return clean_name
    return f"{path_prefix}/{clean_name}"

def format_path_for_squash(raw_path, import_root=None):
    """Build Squash path: /ProjectName[/root]/folder1/folder2/... Leave root empty to match TestLink structure."""
    base = f"/{PROJECT_NAME}"
    root = import_root if import_root is not None else IMPORT_ROOT_FOLDER
    if root:
        base = f"{base}/{root}"
    if not raw_path:
        return base
    clean_raw = raw_path.replace('\\', '/').replace('//', '/')
    if clean_raw.startswith('/'): clean_raw = clean_raw[1:]
    return f"{base}/{clean_raw}"

def get_node_text(element, tag_name):
    """Get text from child element; use itertext() for CDATA content."""
    for child in element:
        if child.tag.endswith(tag_name):
            text = child.text
            if text is not None:
                return text.strip()
            # CDATA / nested text
            parts = "".join(child.itertext()).strip()
            return parts if parts else ""
    return ""


def _xls_cell(val, max_len=32767):
    if pd.isna(val):
        return ''
    s = str(val)
    return s[:max_len] if len(s) > max_len else s

def write_xls(path, df_tc, df_steps, df_param, df_dataset, df_link):
    import xlwt
    wb = xlwt.Workbook(encoding='utf-8')
    style_header = xlwt.easyxf('font: bold on')
    sheet_names = ['TEST_CASES', 'STEPS', 'PARAMETERS', 'DATASETS', 'LINK_REQ_TC']
    dfs = [df_tc, df_steps, df_param, df_dataset, df_link]
    for name, df in zip(sheet_names, dfs):
        sheet = wb.add_sheet(name)
        for c, col in enumerate(df.columns):
            sheet.write(0, c, str(col), style_header)
        for r in range(len(df)):
            for c, col in enumerate(df.columns):
                sheet.write(r + 1, c, _xls_cell(df.iloc[r][col]))
    wb.save(path)

def split_and_write_files(df_tc, df_steps, xml_file):
    """Split test cases into multiple files to avoid 413 errors on large imports"""
    if df_tc.empty:
        print("No test cases to split")
        return

    # Split test cases into equal groups
    tc_groups = []
    tc_list = df_tc.to_dict('records')
    group_size = len(tc_list) // SPLIT_INTO_PARTS
    remainder = len(tc_list) % SPLIT_INTO_PARTS

    start_idx = 0
    for i in range(SPLIT_INTO_PARTS):
        current_size = group_size + (1 if i < remainder else 0)
        end_idx = start_idx + current_size
        tc_groups.append(tc_list[start_idx:end_idx])
        start_idx = end_idx

    base_name = os.path.splitext(os.path.basename(xml_file))[0]
    output_dir = os.path.dirname(os.path.abspath(xml_file)) or '.'

    for part_num, tc_group in enumerate(tc_groups, 1):
        df_tc_part = pd.DataFrame(tc_group)
        tc_paths = set(df_tc_part['TC_PATH'].tolist())
        df_steps_part = df_steps[df_steps['TC_OWNER_PATH'].isin(tc_paths)]

        print(f"Part {part_num}: {len(df_tc_part)} test cases, {len(df_steps_part)} steps")
        param_cols = ['ACTION', 'TC_OWNER_PATH', 'TC_PARAM_NAME', 'TC_PARAM_DESCRIPTION']
        dataset_cols = ['ACTION', 'TC_OWNER_PATH', 'TC_DATASET_NAME', 'TC_PARAM_OWNER_PATH', 'TC_DATASET_PARAM_NAME', 'TC_DATASET_PARAM_VALUE']
        link_cols = ['REQ_PATH', 'REQ_VERSION_NUM', 'TC_PATH']
        df_param = pd.DataFrame([{c: '' for c in param_cols}], columns=param_cols)
        df_dataset = pd.DataFrame([{c: '' for c in dataset_cols}], columns=dataset_cols)
        df_link = pd.DataFrame([{c: '' for c in link_cols}], columns=link_cols)

        # Output filename with part number
        part_suffix = f"_part{part_num}"
        output_xls = base_name + f"_SquashTM_Import{part_suffix}.xls"
        output_xlsx = base_name + f"_SquashTM_Import{part_suffix}.xlsx"
        if output_dir:
            output_xls = os.path.join(output_dir, output_xls)
            output_xlsx = os.path.join(output_dir, output_xlsx)

        try:
            write_xls(output_xls, df_tc_part, df_steps_part, df_param, df_dataset, df_link)
            print(f"SUCCESS! Part {part_num} created (XLS): {output_xls}")
        except ImportError:
            with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
                df_tc_part.to_excel(writer, sheet_name='TEST_CASES', index=False)
                df_steps_part.to_excel(writer, sheet_name='STEPS', index=False)
                df_param.to_excel(writer, sheet_name='PARAMETERS', index=False)
                df_dataset.to_excel(writer, sheet_name='DATASETS', index=False)
                df_link.to_excel(writer, sheet_name='LINK_REQ_TC', index=False)
            print(f"SUCCESS! Part {part_num} created (XLSX). Install xlwt for .xls: pip install xlwt")
        except Exception as xls_err:
            with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
                df_tc_part.to_excel(writer, sheet_name='TEST_CASES', index=False)
                df_steps_part.to_excel(writer, sheet_name='STEPS', index=False)
                df_param.to_excel(writer, sheet_name='PARAMETERS', index=False)
                df_dataset.to_excel(writer, sheet_name='DATASETS', index=False)
                df_link.to_excel(writer, sheet_name='LINK_REQ_TC', index=False)
            print(f"SUCCESS! Part {part_num} created (XLSX). (.xls failed: {xls_err})")

def main():
    print(f"Target Project: '{PROJECT_NAME}'")
    print(f"Import root folder: '{IMPORT_ROOT_FOLDER or '(none - TestLink structure)'}'")
    if not IMPORT_ROOT_FOLDER:
        print("NOTE: For 500 errors, delete existing content or set IMPORT_ROOT_FOLDER")

    if len(sys.argv) >= 2:
        xml_file = sys.argv[1]
        if not os.path.isfile(xml_file):
            print(f"ERROR: File not found: {xml_file}")
            return
    else:
        files = [f for f in os.listdir('.') if f.endswith('.xml')]
        if not files:
            print("ERROR: No .xml file found. Usage: python tl2squash.py [path/to/export.xml]")
            return
        xml_file = files[0]
    print(f"Processing: {xml_file}")

    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        top_suites = [c for c in root if c.tag.endswith('testsuite')]
        first_level = [s.get('name', '') for s in top_suites if s.get('name')]
        print(f"TestLink folders: {', '.join(first_level[:10])}{'...' if len(first_level) > 10 else ''}")

        tc_rows = []
        step_rows = []
        seen_folders = set()
        seen_test_cases = {}
        effective_root = IMPORT_ROOT_FOLDER

        def get_children(elem, tag_suffix):
            return [child for child in elem if child.tag.endswith(tag_suffix)]

        def format_path(raw_path):
            return format_path_for_squash(raw_path, effective_root)

        def parse_suite(suite_element, path_prefix=""):
            suite_name = suite_element.get('name', '')
            if not suite_name:
                current_raw_path = path_prefix
            else:
                current_raw_path = build_folder_path(path_prefix, suite_name, seen_folders)
            
            testcases = get_children(suite_element, 'testcase')
            for testcase in testcases:
                raw_tc_name = testcase.get('name', 'Unnamed Case')
                tc_name = sanitize_text(raw_tc_name)
                if not tc_name: tc_name = "Unnamed Case"

                squash_folder_path = format_path(current_raw_path)

                key = (squash_folder_path, tc_name)
                if key in seen_test_cases:
                    seen_test_cases[key] += 1
                    final_tc_name = f"{tc_name}_{generate_short_id()}"
                else:
                    seen_test_cases[key] = 1
                    final_tc_name = tc_name

                full_tc_path = f"{squash_folder_path}/{final_tc_name}"

                tc_description = rich_text_to_html(get_node_text(testcase, 'summary'))
                tc_prerequisite = rich_text_to_html(get_node_text(testcase, 'preconditions'))
                exec_val = get_node_text(testcase, 'execution_type')
                execution_status = "Automated" if exec_val == "2" else "Manual"
                testlink_id = get_node_text(testcase, 'externalid') if MIGRATE_TESTLINK_ID else ''

                tc_row = {
                    'ACTION': 'C',
                    'TC_PATH': full_tc_path,
                    'TC_NUM': '',
                    'TC_UUID': '',
                    'TC_REFERENCE': testlink_id if not SQUASH_CUF_TESTLINK_ID else '',
                    'TC_NAME': '',
                    'TC_MILESTONE': '',
                    'TC_WEIGHT_AUTO': '0',
                    'TC_WEIGHT': 'LOW',
                    'TC_NATURE': 'NAT_UNDEFINED',
                    'TC_TYPE': 'TYP_UNDEFINED',
                    'TC_STATUS': 'WORK_IN_PROGRESS',
                    'TC_DESCRIPTION': tc_description,
                    'TC_PRE_REQUISITE': tc_prerequisite,
                    'TC_CREATED_ON': '',
                    'TC_CREATED_BY': '',
                    'DRAFTED_BY_AI': '0',
                    'TC_KIND': 'STANDARD',
                    'TC_SCRIPT': '',
                    'TC_AUTOMATABLE': 'M',
                    f'TC_CUF_{SQUASH_CUF_CODE}': execution_status
                }

                if SQUASH_CUF_TESTLINK_ID and testlink_id:
                    tc_row[f'TC_CUF_{SQUASH_CUF_TESTLINK_ID}'] = testlink_id

                tc_rows.append(tc_row)
                steps_container_list = get_children(testcase, 'steps')
                if steps_container_list:
                    steps = get_children(steps_container_list[0], 'step')
                    for i, step in enumerate(steps, 1):
                        action = rich_text_to_html(get_node_text(step, 'actions'))
                        expected = rich_text_to_html(get_node_text(step, 'expectedresults'))
                        step_num = 1
                        step_rows.append({
                            'ACTION': 'C',
                            'TC_OWNER_PATH': full_tc_path,
                            'TC_STEP_NUM': step_num,
                            'TC_STEP_IS_CALL_STEP': 0,
                            'TC_STEP_CALL_DATASET': '',
                            'TC_STEP_ACTION': action,
                            'TC_STEP_EXPECTED_RESULT': expected
                        })

            sub_suites = get_children(suite_element, 'testsuite')
            for sub in sub_suites:
                parse_suite(sub, current_raw_path)

        if root.tag.endswith('testsuite'):
            parse_suite(root, "")
        else:
            top_suites = get_children(root, 'testsuite')
            for s in top_suites: parse_suite(s, "")
        
        TC_COLUMNS = [
            'ACTION', 'TC_PATH', 'TC_NUM', 'TC_UUID', 'TC_REFERENCE', 'TC_NAME',
            'TC_MILESTONE', 'TC_WEIGHT_AUTO', 'TC_WEIGHT', 'TC_NATURE', 'TC_TYPE',
            'TC_STATUS', 'TC_DESCRIPTION', 'TC_PRE_REQUISITE', 'TC_CREATED_ON',
            'TC_CREATED_BY', 'DRAFTED_BY_AI', 'TC_KIND', 'TC_SCRIPT', 'TC_AUTOMATABLE',
            f'TC_CUF_{SQUASH_CUF_CODE}'
        ]

        if SQUASH_CUF_TESTLINK_ID:
            TC_COLUMNS.append(f'TC_CUF_{SQUASH_CUF_TESTLINK_ID}')
        STEP_COLUMNS = [
            'ACTION', 'TC_OWNER_PATH', 'TC_STEP_NUM', 'TC_STEP_IS_CALL_STEP',
            'TC_STEP_CALL_DATASET', 'TC_STEP_ACTION', 'TC_STEP_EXPECTED_RESULT'
        ]
        df_tc = pd.DataFrame(tc_rows, columns=TC_COLUMNS).fillna("")
        df_steps = pd.DataFrame(step_rows, columns=STEP_COLUMNS).fillna("")

        if not df_tc.empty:
            df_tc = df_tc[df_tc['TC_PATH'].str.strip() != ""]

        print(f"Generated {len(df_tc)} Test Cases, {len(df_steps)} Steps")

        if SPLIT_INTO_PARTS > 1:
            split_and_write_files(df_tc, df_steps, xml_file)
            return
        param_cols = ['ACTION', 'TC_OWNER_PATH', 'TC_PARAM_NAME', 'TC_PARAM_DESCRIPTION']
        dataset_cols = ['ACTION', 'TC_OWNER_PATH', 'TC_DATASET_NAME', 'TC_PARAM_OWNER_PATH', 'TC_DATASET_PARAM_NAME', 'TC_DATASET_PARAM_VALUE']
        link_cols = ['REQ_PATH', 'REQ_VERSION_NUM', 'TC_PATH']
        df_param = pd.DataFrame([{c: '' for c in param_cols}], columns=param_cols)
        df_dataset = pd.DataFrame([{c: '' for c in dataset_cols}], columns=dataset_cols)
        df_link = pd.DataFrame([{c: '' for c in link_cols}], columns=link_cols)

        # Output next to script if input was path, else next to XML. Use .xls like Squash TM template.
        base_name = os.path.splitext(os.path.basename(xml_file))[0]
        output_dir = os.path.dirname(os.path.abspath(xml_file))
        output_xls = base_name + "_SquashTM_Import.xls"
        output_xlsx = base_name + "_SquashTM_Import.xlsx"
        if output_dir:
            output_xls = os.path.join(output_dir, output_xls)
            output_xlsx = os.path.join(output_dir, output_xlsx)
        
        try:
            write_xls(output_xls, df_tc, df_steps, df_param, df_dataset, df_link)
            print(f"SUCCESS! File created (XLS, like Squash TM template): {output_xls}")
        except ImportError:
            with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
                df_tc.to_excel(writer, sheet_name='TEST_CASES', index=False)
                df_steps.to_excel(writer, sheet_name='STEPS', index=False)
                df_param.to_excel(writer, sheet_name='PARAMETERS', index=False)
                df_dataset.to_excel(writer, sheet_name='DATASETS', index=False)
                df_link.to_excel(writer, sheet_name='LINK_REQ_TC', index=False)
            print(f"SUCCESS! File created (XLSX). For .xls like the template, install xlwt: pip install xlwt")
        except Exception as xls_err:
            with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
                df_tc.to_excel(writer, sheet_name='TEST_CASES', index=False)
                df_steps.to_excel(writer, sheet_name='STEPS', index=False)
                df_param.to_excel(writer, sheet_name='PARAMETERS', index=False)
                df_dataset.to_excel(writer, sheet_name='DATASETS', index=False)
                df_link.to_excel(writer, sheet_name='LINK_REQ_TC', index=False)
            print(f"SUCCESS! File created (XLSX). (.xls failed: {xls_err})")
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()