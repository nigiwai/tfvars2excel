import openpyxl
import sys
import json
import os
from dotenv import load_dotenv

# .envファイルを読み込む
load_dotenv()
SHEET_PREFIXES = os.getenv("3_SHEET_NAME_PREFIXES").split(",")
data_row_start = int(os.getenv("3_DATA_ROW_START", "5"))

def pretty_format_tf(value, indent=0):
    spaces = "  " * indent  # インデントを2スペースに変更
    if isinstance(value, dict):
        if not value:
            return "{}"
        lines = []
        for k, v in value.items():
            key_str = k if k.isidentifier() else json.dumps(k, ensure_ascii=False)
            lines.append(f"{spaces}  {key_str} = {pretty_format_tf(v, indent+1)}")
        return "{\n" + "\n".join(lines) + f"\n{spaces}}}"
    elif isinstance(value, list):
        if not value:
            return "[]"
        lines = []
        for idx, item in enumerate(value):
            line = pretty_format_tf(item, indent + 1)
            if idx < len(value) - 1:
                line += ","
            lines.append(line)
        inner = "\n".join("  " * (indent + 1) + line for line in lines)
        return "[\n" + inner + f"\n{spaces}]"
    else:
        if isinstance(value, str):
            return json.dumps(value, ensure_ascii=False)
        # 修正: bool 値を小文字で出力する
        elif isinstance(value, bool):
            return "true" if value else "false"
        else:
            return str(value)

def pretty_format_map(value, key_string="key", indent=0):
    spaces = "  " * indent
    if isinstance(value, list):
        new_dict = {}
        for i, item in enumerate(value, start=1):
            new_dict[f"{key_string}{i:02d}"] = item
        return pretty_format_tf(new_dict, indent)
    else:
        return pretty_format_tf(value, indent)

def pretty_format_list_object(value, indent=0):
    spaces = "  " * indent
    if isinstance(value, list):
        lines = []
        for i, item in enumerate(value, start=1):
            # 各行末尾に常にカンマを付与
            item_str = pretty_format_tf(item, indent + 1)
            line = f"{spaces}  {item_str},"
            lines.append(line)
        inner = "\n".join(lines)
        return "[\n" + inner + f"\n{spaces}]"
    else:
        return pretty_format_tf(value, indent)

def convert(elem_type,val):
    if elem_type == "number":
        try:
            return (
                float(val) if "." in val else int(val)
            )
        except:
            return val
    elif elem_type == "bool":
        return val.lower() == "true"
    elif elem_type == "list":
        if val is None:
            return []
        return [item.strip() for item in val.split(",")]
    elif elem_type == "list(object)":
        if val is None:
            return []
        return val
    elif elem_type == "object":
        if val is None:
            return {}
        return val
    elif elem_type == "map(object)":
        if val is None:
            return {}
        return val
    else:
        return val

# 値フォーマットをまとめるヘルパー関数
def format_value(value, typ, field=None, object_defs=None):
    # 空やNoneの場合は type に応じて適切に処理
    if typ in ["string", "number", "bool"]:
        # Noneや空文字の場合はnullで出力
        if not value and value != 0:
            return "null"
        if typ == "bool":
            # 修正: bool値は必ず小文字で出力
            return "true" if value else "false"
        elif typ == "string":
            return f'"{value}"'
        else:  # number
            return str(value)
    elif typ == "list":
        # 空の場合は空リスト
        return json.dumps(value or [], ensure_ascii=False)
    elif typ == "map":
        # 空の場合は {}
        return pretty_format_tf(value or {}, 2)
    elif typ == "object":
        # Noneの場合は null
        if value is None:
            return "null"
        return pretty_format_tf(value, 2)
    elif typ == "list(object)":
        # Noneの場合は空リスト
        if value is None:
            return "[]"
        return pretty_format_list_object(value, 2)
    elif typ == "map(object)":
        # Noneの場合は {}
        if value is None:
            return "{}"
        key_str = object_defs.get(field, {}).get("key", field) if object_defs else field
        return pretty_format_map(value, key_str, 2)
    else:
        raise ValueError(f"Unsupported type: {typ}")

def excel_to_tfvars(excel_file, sheet_name, output_file):
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb[sheet_name]

    # ヘッダー等の読み込み（先頭3行）
    max_col = sheet.max_column
    headers = []
    col_types = []
    object_defs = {}
    object_nums = []
    for col in range(1, max_col + 1):
        header = sheet.cell(row=1, column=col).value
        headers.append(header)
        col_type = (sheet.cell(row=2, column=col).value or "string")
        col_types.append(col_type)
        if col_type in ["map(object)", "object", "list(object)"]:
            definitions = sheet.cell(row=3, column=col).value
            object_def = {}
            if definitions:
                # 定義は "aaa:string" といったフォーマット
                for line in definitions.splitlines():
                    if ":" in line:
                        elem, elem_type = line.split(":", 1)
                        object_def[elem.strip()] = elem_type.strip()
            if col_type == "map(object)" and "key" not in object_def:
                raise ValueError(f"Error: 'key' field is missing for map(object) in {header}")
            object_defs[col - 1] = object_def
        object_num = sheet.cell(row=4, column=col).value
        try:
            object_nums.append(int(object_num))
        except (ValueError, TypeError):
            object_nums.append(0)
    # 変更: 重複ヘッダーの場合は定義をマージする
    object_definitions = {}
    for i in range(len(headers)):
        if col_types[i] in ["map(object)", "object", "list(object)"]:
            key = headers[i]
            current_def = object_defs.get(i, {})
            if key in object_definitions:
                object_definitions[key].update(current_def)
            else:
                object_definitions[key] = current_def

    # 変更: 参照元・参照先対応のための ref_map 作成
    # 参照元の maps 定義内に "xheader:maps" のような記述があれば、
    # ref_map[<参照元ヘッダー>] = <参照先ヘッダー>
    ref_map = {}
    for i, hdr in enumerate(headers):
        if (
            col_types[i] in ["map(object)", "object", "list(object)"]
            and hdr in object_definitions
        ):
            for field, ftype in object_definitions[hdr].items():
                if ftype in ["map(object)", "object", "list(object)", "list"]:
                    if hdr not in ref_map:
                        ref_map[hdr] = []
                    ref_map[hdr].append(field)

    header_types = { headers[i]: col_types[i] for i in range(1, max_col) }

    # データ行は5行目以降
    tfvars_map = {}
    for row_idx, row in enumerate(sheet.iter_rows(min_row=data_row_start, values_only=True), start=data_row_start):
        key = row[0]
        if key is None:
            raise ValueError(f"Error: Key is None in row {row_idx}")
        values = {}
        for idx in range(1, max_col):
            cell_value = row[idx]
            col_type = col_types[idx]
            header = headers[idx]
            object_num = object_nums[idx]
            # 型に応じた変換処理
            if cell_value is not None or cell_value == "":
                if col_type == "string":
                    converted_value = str(cell_value) 
                elif col_type == "number":
                    try:
                        converted_value = (
                            float(cell_value)
                            if "." in str(cell_value)
                            else int(cell_value)
                        )
                    except:
                        converted_value = cell_value

                elif col_type == "bool":
                    # 修正: cell_valueがNoneの場合はNoneにする
                    if cell_value != "true" and cell_value != "false":
                        raise ValueError(
                            f"Error in sheet '{sheet.title}' for key '{key}', type:{col_type} '{header}' expects 'true' or 'false' but got '{cell_value}'"
                        )
                    converted_value = (
                        str(cell_value).lower()
                    )
                elif col_type == "list":
                    converted_value = (
                        cell_value.splitlines()
                        if (isinstance(cell_value, str) and "\n" in cell_value)
                        else [cell_value]
                    )
                elif col_type in ["map(object)", "object", "list(object)"]:
                    # 各改行ごとにひとつのオブジェクトとして処理し、常にリストとして返す
                    lines = str(cell_value).splitlines()
                    if header in object_definitions:
                        objects = []
                        for line in lines:
                            values_list = [p.strip() for p in line.split(":")]
                            if len(values_list) != object_nums[idx]:
                                raise ValueError(
                                    f"Error in sheet '{sheet.title}' for key '{key}', type:{col_type} '{header}' expects {object_num} elements but got {len(values_list)} in line: {line}"
                                )
                            obj = {}
                            # "maps" は先頭項目を除外
                            for i, key_elem in enumerate(
                                object_definitions[header].keys()
                            ):
                                # 変更: "map(object)" の場合は先頭項目を除外
                                if col_type == "map(object)" and i == 0:
                                    continue
                                idx_offset = i - 1 if col_type == "map(object)" else i
                                # 変更: インデックス範囲外なら空文字列（またはNone）を設定
                                raw = (
                                    values_list[idx_offset]
                                    if idx_offset < len(values_list)
                                    else None
                                )
                                elem_type = object_definitions[header][key_elem]
                                obj[key_elem] = convert(elem_type,raw)
                            objects.append(obj)
                        # 変更: "object" の場合は単一オブジェクトを返す
                        if col_type == "object":
                            converted_value = objects[0] if objects else None
                        else:
                            converted_value = objects
                    else:
                        converted_value = cell_value
                else:
                    converted_value = cell_value
                values[header] = converted_value
            else:
                if col_type in ["map(object)", "object"]:
                    values[header] = {}
                elif col_type in ["list", "list(object)"]:
                    values[header] = []
                else:
                    values[header] = None
        # 変更: 参照先の値を参照元のmapsにネストして結合する
        for src_hdr, dest_list in ref_map.items():
            for dest_hdr in dest_list:
                if (
                    src_hdr in {h for h in headers} 
                    and dest_hdr in {h for h in headers}
                ):
                    # 参照先セルが None または空リスト等の場合は処理をスキップ
                    if (
                        src_hdr in values
                        and dest_hdr in values
                        and values[dest_hdr] is not None
                    ):
                        src_val = values[src_hdr]
                        if not isinstance(src_val, list):
                            src_val = [src_val]
                        dest_val = values[dest_hdr]
                        if isinstance(dest_val, list):
                            dest_objs = dest_val
                        else:
                            dest_objs = [dest_val]
                        # 追加: 参照先が空の場合はスキップ
                        if not dest_objs:
                            continue
                        # 各参照元マップオブジェクトに対して、参照先の内容をネストする
                        for obj_index, obj in enumerate(src_val):
                            if obj is None or not isinstance(obj, dict):
                                obj = {}
                                src_val[obj_index] = obj
                            counter = 1
                            dest_def = object_definitions.get(dest_hdr, {})
                            dest_prefix = dest_def.get("key", dest_hdr)
                            # 参照先フィールドごとの処理
                            for d_obj in dest_objs:
                                dest_field_type = header_types.get(dest_hdr)
                                if dest_field_type == "list":
                                    if dest_hdr not in obj or not isinstance(obj[dest_hdr], list):
                                        obj[dest_hdr] = []
                                    obj[dest_hdr].append(d_obj)
                                elif dest_field_type == "object":
                                    if dest_hdr not in obj or not isinstance(obj[dest_hdr], dict):
                                        obj[dest_hdr] = {}
                                    obj[dest_hdr] = d_obj
                                elif dest_field_type in ["map(object)"]:
                                    if dest_hdr not in obj or not isinstance(obj[dest_hdr], dict):
                                        obj[dest_hdr] = {}
                                    sub_key = f"{dest_prefix}{counter:02d}"
                                    obj[dest_hdr][sub_key] = d_obj
                                elif dest_field_type == "list(object)":
                                    if dest_hdr not in obj or not isinstance(obj[dest_hdr], list):
                                        obj[dest_hdr] = []
                                    obj[dest_hdr].append(d_obj)
                                else:
                                    raise ValueError(
                                        f"Error: Unsupported type '{dest_field_type}' for '{dest_hdr}'"
                                    )
                                counter += 1
                            values[src_hdr] = src_val
                    values.pop(dest_hdr, None)
        tfvars_map[key] = values
    # 各ヘッダーの型情報をマッピング（先頭列以外）

    # tfvarsファイルとして出力（Terraformの各typeをシンプルに処理）
    mode = "a" if os.path.exists(output_file) else "w"
    with open(output_file, mode, encoding="utf-8", newline="\n") as tfvars_file:
        tfvars_file.write(f"{sheet.title} = {{\n")

        # シート内の各キーごとにブロックを出力
        for key, data in tfvars_map.items():
            tfvars_file.write(f'  "{key}" = {{\n')

            for field, value in data.items():
                typ = header_types.get(field)
                # 新しいヘルパー関数でフォーマットを取得
                val_str = format_value(value, typ, field, object_definitions)
                tfvars_file.write(f"    {field} = {val_str}\n")

            tfvars_file.write("  },\n")
        tfvars_file.write("}\n")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python 3_excel2map.py <excel_filepath>")
        sys.exit(1)

    excel_file_path = sys.argv[1]

    # Excelファイル名からフォルダ名を作成
    excel_file_name = os.path.splitext(os.path.basename(excel_file_path))[0]
    output_folder = os.path.join("output", excel_file_name)
    os.makedirs(output_folder, exist_ok=True)
    output_tfvars_file = os.path.join(output_folder, "terraform.tfvars")

    wb = openpyxl.load_workbook(excel_file_path)
    for sheet_name in wb.sheetnames:
        if any(sheet_name.startswith(prefix) for prefix in SHEET_PREFIXES):
            print(f"Converting {sheet_name} sheet to terraform.tfvars")
            excel_to_tfvars(excel_file_path, sheet_name, output_tfvars_file)
