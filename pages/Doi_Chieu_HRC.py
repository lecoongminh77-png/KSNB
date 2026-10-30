À, đúng rồi — bạn cần **nguyên một file code** để dán vào `app.py`. Đây là bản đầy đủ, bắt đầu bằng `import io` và `import re`. Cài thư viện bằng `pip install streamlit pandas openpyxl`, rồi chạy `streamlit run app.py`.

```python
import io
import re
import math
import numbers
import posixpath
import shutil
import unicodedata
import xml.etree.ElementTree as ET
from collections import defaultdict
from io import BytesIO
from typing import Any, Iterable, Sequence
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

RESULT_HEADERS = (
    (14, "NHẬP/XUẤT"),
    (15, "Mã hóa"),
    (16, "Khối lượng net theo PL"),
    (17, "Đối chiếu"),
    (18, "KL theo bc nhập hàng"),
    (19, "Số phiếu cân"),
    (20, "Đối chiếu số phiếu cân"),
    (21, "KL theo bc trạm cân"),
    (22, ""),
    (23, "Đối chiếu KL TC & KL BCBV"),
    (24, "Đối chiếu KL TC & KL BBNT"),
    (25, "Chênh lệch (kg)"),
    (26, "Chênh lệch dung sai"),
    (27, "Đối chiếu dung sai"),
    (28, "So với quy định (>+-0.3%)"),
    (29, "KL Nhập kho\n035"),
    (30, "Chênh lệch"),
    (31, "KL Nhập kho\n031"),
    (32, "Chênh lệch"),
    (33, "KL Nhập kho\n03G-3H"),
    (34, "Chênh lệch"),
)


class ReconciliationError(ValueError):
    pass


def as_bytes(file_obj):
    if isinstance(file_obj, bytes):
        return file_obj
    if isinstance(file_obj, bytearray):
        return bytes(file_obj)
    if hasattr(file_obj, "getvalue"):
        return file_obj.getvalue()
    if hasattr(file_obj, "read"):
        position = file_obj.tell() if hasattr(file_obj, "tell") else None
        data = file_obj.read()
        if position is not None and hasattr(file_obj, "seek"):
            file_obj.seek(position)
        return data
    raise TypeError("Đầu vào phải là file Excel hoặc dữ liệu bytes.")


def fold(value):
    text = unicodedata.normalize("NFKD", str(value if value is not None else ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.replace("đ", "d").replace("Đ", "D")
    return re.sub(r"\s+", " ", text).strip().upper()


def clean_id(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(value, numbers.Integral):
        return str(int(value))
    if isinstance(value, numbers.Real) and math.isfinite(float(value)):
        number = float(value)
        if number.is_integer():
            return str(int(number))

    return re.sub(r"\.0+$", "", str(value).strip())


def normalize_ticket(value):
    return clean_id(value)


def normalize_lot(value):
    text = clean_id(value).upper()
    if not text:
        return ""
    return text if text.endswith("A") else text + "A"


def normalize_warehouse(value):
    text = clean_id(value).upper()
    if text.isdigit() and len(text) < 3:
        return text.zfill(3)
    return text


def as_number(value, default=0.0):
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass

    if isinstance(value, numbers.Real):
        number = float(value)
        return number if math.isfinite(number) else default

    text = str(value).strip().replace(" ", "")
    if not text:
        return default

    if "," in text and "." in text:
        text = text.replace(",", "")
    elif "," in text:
        if re.fullmatch(r"-?\d{1,3}(,\d{3})+", text):
            text = text.replace(",", "")
        else:
            text = text.replace(",", ".")

    try:
        number = float(text)
        return number if math.isfinite(number) else default
    except ValueError:
        return default


def workbook_sheet_names(file_obj):
    workbook = load_workbook(
        BytesIO(as_bytes(file_obj)),
        read_only=True,
        data_only=True,
        keep_links=False,
    )
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def part_name(target):
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join("xl", target))


def register_xml_namespaces(xml_bytes):
    try:
        for _, (prefix, uri) in ET.iterparse(
            BytesIO(xml_bytes), events=("start-ns",)
        ):
            ET.register_namespace(prefix or "", uri)
    except ET.ParseError:
        pass


def serialize_xml(root):
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def select_workbook_sheets(file_obj, selected_names):
    """Giữ lại các sheet đã chọn trước khi openpyxl đọc dữ liệu."""
    content = as_bytes(file_obj)
    wanted = {str(name).strip() for name in selected_names}

    if not wanted:
        raise ReconciliationError("Bạn cần chọn ít nhất một sheet BBNT.")

    try:
        source_zip = ZipFile(BytesIO(content), "r")
    except Exception as exc:
        raise ReconciliationError(
            "Không đọc được file BBNT. Hãy dùng file .xlsx hợp lệ."
        ) from exc

    with source_zip:
        required_parts = {
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "[Content_Types].xml",
        }
        if not required_parts.issubset(set(source_zip.namelist())):
            raise ReconciliationError("File BBNT không có cấu trúc .xlsx hợp lệ.")

        workbook_bytes = source_zip.read("xl/workbook.xml")
        rels_bytes = source_zip.read("xl/_rels/workbook.xml.rels")
        types_bytes = source_zip.read("[Content_Types].xml")

        register_xml_namespaces(workbook_bytes)
        register_xml_namespaces(rels_bytes)
        register_xml_namespaces(types_bytes)

        workbook_root = ET.fromstring(workbook_bytes)
        rels_root = ET.fromstring(rels_bytes)
        types_root = ET.fromstring(types_bytes)

        sheets_element = workbook_root.find(f"{{{MAIN_NS}}}sheets")
        if sheets_element is None:
            raise ReconciliationError("Không tìm thấy danh sách sheet trong BBNT.")

        relationship_targets = {
            rel.attrib.get("Id", ""): rel.attrib.get("Target", "")
            for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
        }

        sheet_nodes = list(sheets_element)
        selected_nodes = []

        for original_index, sheet in enumerate(sheet_nodes):
            name = sheet.attrib.get("name", "")
            relationship_id = sheet.attrib.get(f"{{{DOC_REL_NS}}}id", "")
            target = relationship_targets.get(relationship_id)

            if name in wanted and target:
                selected_nodes.append(
                    (original_index, sheet, relationship_id, part_name(target))
                )

        selected_found = {
            sheet.attrib.get("name", "")
            for _, sheet, _, _ in selected_nodes
        }
        missing_names = wanted.difference(selected_found)

        if missing_names:
            raise ReconciliationError(
                "Không tìm thấy sheet BBNT: " + ", ".join(sorted(missing_names))
            )

        selected_original_indexes = {
            original_index for original_index, _, _, _ in selected_nodes
        }
        new_sheet_index = {
            original_index: new_index
            for new_index, (original_index, _, _, _) in enumerate(selected_nodes)
        }
        selected_relationship_ids = {
            relationship_id for _, _, relationship_id, _ in selected_nodes
        }
        selected_parts = {part for _, _, _, part in selected_nodes}

        removed_sheet_parts = set()
        removed_names = []

        for sheet in sheet_nodes:
            name = sheet.attrib.get("name", "")
            if name not in wanted:
                removed_names.append(name)
                relationship_id = sheet.attrib.get(f"{{{DOC_REL_NS}}}id", "")
                target = relationship_targets.get(relationship_id)
                if target:
                    removed_sheet_parts.add(part_name(target))
                sheets_element.remove(sheet)

        book_views = workbook_root.find(f"{{{MAIN_NS}}}bookViews")
        if book_views is not None:
            for view in book_views.findall(f"{{{MAIN_NS}}}workbookView"):
                view.set("activeTab", "0")
                view.set("firstSheet", "0")

        defined_names = workbook_root.find(f"{{{MAIN_NS}}}definedNames")
        if defined_names is not None:
            for defined_name in list(defined_names):
                local_sheet_id = defined_name.attrib.get("localSheetId")

                if local_sheet_id is not None:
                    try:
                        original_index = int(local_sheet_id)
                    except ValueError:
                        defined_names.remove(defined_name)
                        continue

                    if original_index not in selected_original_indexes:
                        defined_names.remove(defined_name)
                    else:
                        defined_name.set(
                            "localSheetId",
                            str(new_sheet_index[original_index]),
                        )
                else:
                    formula = (defined_name.text or "").casefold()
                    refers_to_removed_sheet = any(
                        f"'{name.casefold()}'!" in formula
                        or f"{name.casefold()}!" in formula
                        for name in removed_names
                        if name
                    )
                    if refers_to_removed_sheet:
                        defined_names.remove(defined_name)

            if not list(defined_names):
                workbook_root.remove(defined_names)

        removed_parts = set(removed_sheet_parts)

        for part in removed_sheet_parts:
            rel_part = posixpath.join(
                posixpath.dirname(part),
                "_rels",
                posixpath.basename(part) + ".rels",
            )
            removed_parts.add(rel_part)

        for relationship in list(
            rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
        ):
            rel_type = relationship.attrib.get("Type", "")
            relationship_id = relationship.attrib.get("Id", "")
            target = relationship.attrib.get("Target", "")

            if (
                rel_type.endswith("/worksheet")
                and relationship_id not in selected_relationship_ids
            ):
                rels_root.remove(relationship)
            elif rel_type.endswith("/calcChain"):
                removed_parts.add(part_name(target))
                rels_root.remove(relationship)

        removed_parts.add("xl/calcChain.xml")

        for override in list(
            types_root.findall(f"{{{CONTENT_TYPES_NS}}}Override")
        ):
            part = override.attrib.get("PartName", "").lstrip("/")
            if part in removed_parts or part == "xl/calcChain.xml":
                types_root.remove(override)

        replacements = {
            "xl/workbook.xml": serialize_xml(workbook_root),
            "xl/_rels/workbook.xml.rels": serialize_xml(rels_root),
            "[Content_Types].xml": serialize_xml(types_root),
        }

        output = BytesIO()
        with ZipFile(
            output,
            "w",
            compression=ZIP_DEFLATED,
            compresslevel=6,
        ) as result_zip:
            for info in source_zip.infolist():
                name = info.filename

                if name in removed_parts:
                    continue

                if name in replacements:
                    result_zip.writestr(info, replacements[name])
                    continue

                if info.is_dir():
                    result_zip.writestr(info, b"")
                    continue

                with source_zip.open(info, "r") as input_stream:
                    with result_zip.open(info, "w") as output_stream:
                        shutil.copyfileobj(
                            input_stream,
                            output_stream,
                            length=1024 * 1024,
                        )

        return output.getvalue()


def find_header_row(
    file_content,
    required_groups,
    sheet_name=0,
    max_scan=40,
    file_label="file Excel",
):
    try:
        raw = pd.read_excel(
            BytesIO(file_content),
            sheet_name=sheet_name,
            header=None,
            nrows=max_scan,
            engine="openpyxl",
        )
    except Exception as exc:
        raise ReconciliationError(
            f"Không đọc được phần đầu của {file_label}."
        ) from exc

    folded_groups = [
        [fold(keyword) for keyword in group]
        for group in required_groups
    ]

    for row_index, row in raw.iterrows():
        row_text = " ".join(
            fold(value)
            for value in row.values
            if not pd.isna(value)
        )

        if all(
            any(keyword in row_text for keyword in group)
            for group in folded_groups
        ):
            return int(row_index)

    requirements = " và ".join(
        "/".join(group) for group in required_groups
    )
    raise ReconciliationError(
        f"Không tìm thấy dòng tiêu đề gồm {requirements} trong {file_label}."
    )


def read_table(
    file_content,
    required_groups,
    sheet_name=0,
    file_label="file Excel",
):
    header_row = find_header_row(
        file_content,
        required_groups,
        sheet_name=sheet_name,
        file_label=file_label,
    )
    try:
        return pd.read_excel(
            BytesIO(file_content),
            sheet_name=sheet_name,
            skiprows=header_row,
            engine="openpyxl",
        )
    except Exception as exc:
        raise ReconciliationError(
            f"Không đọc được bảng dữ liệu trong {file_label}."
        ) from exc


def find_column(
    columns,
    label,
    aliases=(),
    contains=(),
    excludes=(),
    required=True,
):
    columns = list(columns)
    aliases = [fold(value) for value in aliases]
    contains = [fold(value) for value in contains]
    excludes = [fold(value) for value in excludes]

    for alias in aliases:
        for column in columns:
            current = fold(column)
            if (current == alias or alias in current) and not any(
                item in current for item in excludes
            ):
                return column

    for column in columns:
        current = fold(column)
        if not contains or not all(item in current for item in contains):
            continue
        if any(item in current for item in excludes):
            continue
        return column

    if required:
        raise ReconciliationError(
            f"Không tìm thấy cột “{label}”. Hãy kiểm tra dòng tiêu đề của file."
        )
    return None


def read_protection_records(file_content):
    df = read_table(
        file_content,
        required_groups=(
            ("MÃ PHIẾU CÂN", "PHIẾU CÂN"),
            ("KHỐI LƯỢNG",),
        ),
        file_label="file Bảo vệ",
    )

    ticket_col = find_column(
        df.columns,
        label="Mã phiếu cân trong file Bảo vệ",
        aliases=("MÃ PHIẾU CÂN", "PHIẾU CÂN"),
        contains=("PHIẾU CÂN",),
    )
    weight_col = find_column(
        df.columns,
        label="Khối lượng trong file Bảo vệ",
        contains=("KHỐI LƯỢNG",),
    )
    code_col = find_column(
        df.columns,
        label="Mã hàng hóa trong file Bảo vệ",
        aliases=("MÃ HÀNG HÓA",),
        required=False,
    )
    name_col = find_column(
        df.columns,
        label="Tên hàng hóa trong file Bảo vệ",
        aliases=("TÊN HÀNG HÓA",),
        required=False,
    )

    if code_col is None and name_col is None:
        raise ReconciliationError(
            "File Bảo vệ cần có cột Mã hàng hóa hoặc Tên hàng hóa."
        )

    records = []
    for _, row in df.iterrows():
        description = " ".join(
            clean_id(row.get(column, ""))
            for column in (code_col, name_col)
            if column is not None
        )
        ticket = normalize_ticket(row.get(ticket_col))

        if description and ticket:
            records.append(
                (description, ticket, as_number(row.get(weight_col), 0.0))
            )

    return records


def read_scale_data(file_content):
    df = read_table(
        file_content,
        required_groups=(("PHIẾU CÂN",), ("KHỐI LƯỢNG",)),
        file_label="file Trạm cân",
    )

    ticket_col = find_column(
        df.columns,
        label="Phiếu cân trong file Trạm cân",
        aliases=("PHIẾU CÂN",),
    )
    weight_col = find_column(
        df.columns,
        label="Khối lượng trong file Trạm cân",
        contains=("KHỐI LƯỢNG",),
    )
    class_col = find_column(
        df.columns,
        label="Phân loại trong file Trạm cân",
        contains=("PHÂN LOẠI",),
        required=False,
    )

    weights = {}
    classifications = {}

    for _, row in df.iterrows():
        ticket = normalize_ticket(row.get(ticket_col))
        if not ticket:
            continue

        weights[ticket] = as_number(row.get(weight_col), 0.0)

        if class_col is not None:
            classifications[ticket] = (
                clean_id(row.get(class_col)) or "Nhập hàng"
            )

    return weights, classifications


def read_opm_data(file_content):
    df = read_table(
        file_content,
        required_groups=(("LOẠI GIAO DỊCH",), ("MÃ LÔ",)),
        file_label="file ERP OPM 105",
    )

    transaction_col = find_column(
        df.columns,
        label="Loại giao dịch trong file OPM",
        contains=("LOẠI GIAO DỊCH",),
    )
    lot_col = find_column(
        df.columns,
        label="Mã lô trong file OPM",
        contains=("MÃ LÔ",),
    )
    weight_col = find_column(
        df.columns,
        label="Khối lượng net trong file OPM",
        contains=("KHỐI LƯỢNG",),
        excludes=("GROSS",),
    )
    warehouse_col = find_column(
        df.columns,
        label="Mã kho trong file OPM",
        aliases=("MÃ KHO", "MĂ KHO"),
    )

    # Một số file OPM có dòng hướng dẫn bắt đầu bằng dấu "[" ngay dưới tiêu đề.
    if len(df) > 0 and len(df.columns) > 1 and "[" in str(df.iloc[0, 1]):
        df = df.iloc[1:]

    po_035 = defaultdict(float)
    rtv_035 = defaultdict(float)
    intransit_031 = defaultdict(float)
    po_03g = defaultdict(float)
    rtv_03g = defaultdict(float)

    for _, row in df.iterrows():
        key = normalize_lot(row.get(lot_col))
        if not key:
            continue

        transaction = clean_id(row.get(transaction_col)).casefold()
        warehouse = normalize_warehouse(row.get(warehouse_col))
        weight = as_number(row.get(weight_col), 0.0)

        if warehouse == "035" and "po receipt" in transaction:
            po_035[key] += weight
        if warehouse == "035" and "return to vendor" in transaction:
            rtv_035[key] += weight
        if warehouse in ("03G", "03H") and "intransit receipt" in transaction:
            intransit_031[key] += weight
        if warehouse == "03G" and "po receipt" in transaction:
            po_03g[key] += weight
        if warehouse == "03G" and "return to vendor" in transaction:
            rtv_03g[key] += weight

    return (
        dict(po_035),
        dict(rtv_035),
        dict(intransit_031),
        dict(po_03g),
        dict(rtv_03g),
    )


def read_packing_list(file_content, selected_sheet, available_sheets):
    sheet_map = {fold(sheet): sheet for sheet in available_sheets}
    matched_sheet = sheet_map.get(fold(selected_sheet))

    if matched_sheet is None:
        raise ReconciliationError(
            f"Không tìm thấy sheet “{selected_sheet}” trong Packing List."
        )

    df = read_table(
        file_content,
        required_groups=(
            ("MÃ SỐ NSX", "MÃ SỐ HSPM", "MÃ SỐ"),
            ("NET",),
        ),
        sheet_name=matched_sheet,
        file_label=f"sheet Packing List “{matched_sheet}”",
    )

    lot_col = find_column(
        df.columns,
        label="Mã số cuộn trong Packing List",
        aliases=("MÃ SỐ NSX", "MÃ SỐ HSPM"),
        contains=("MÃ SỐ",),
    )
    net_col = find_column(
        df.columns,
        label="Khối lượng net trong Packing List",
        contains=("NET",),
    )

    weights = defaultdict(float)
    for _, row in df.iterrows():
        key = normalize_lot(row.get(lot_col))
        if key:
            weights[key] += as_number(row.get(net_col), 0.0)

    return dict(weights)


def find_bbnt_header_row(worksheet):
    for row_number in range(1, min(worksheet.max_row or 14, 20) + 1):
        value = worksheet.cell(row=row_number, column=4).value
        if value and "CUỘN" in str(value).upper():
            return row_number

    raise ReconciliationError(
        f"Không tìm thấy tiêu đề “Mã số cuộn” trong sheet “{worksheet.title}”."
    )


def first_protection_match(coil_code, protection_records):
    needle = coil_code.casefold()

    for description, ticket, weight in protection_records:
        if needle in description.casefold():
            return ticket, weight

    return "", 0.0


def write_result_headers(worksheet, header_row):
    yellow_fill = PatternFill(
        start_color="FFFFFF00",
        end_color="FFFFFF00",
        fill_type="solid",
    )

    for column, label in RESULT_HEADERS:
        cell = worksheet.cell(row=header_row, column=column)
        cell.value = label
        cell.fill = yellow_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            wrap_text=True,
            horizontal="center",
            vertical="center",
        )


def process_workbooks(
    bbnt_file,
    packing_file,
    scale_file,
    protection_file,
    opm_file,
    selected_sheets,
):
    if not selected_sheets:
        raise ReconciliationError("Bạn cần chọn ít nhất một sheet BBNT.")

    bbnt_bytes = as_bytes(bbnt_file)
    packing_bytes = as_bytes(packing_file)
    scale_bytes = as_bytes(scale_file)
    protection_bytes = as_bytes(protection_file)
    opm_bytes = as_bytes(opm_file)

    try:
        with pd.ExcelFile(
            BytesIO(packing_bytes),
            engine="openpyxl",
        ) as excel_file:
            packing_sheets = list(excel_file.sheet_names)
    except Exception as exc:
        raise ReconciliationError(
            "Không đọc được danh sách sheet trong Packing List."
        ) from exc

    packing_names = {fold(name) for name in packing_sheets}
    missing_sheets = [
        name for name in selected_sheets if fold(name) not in packing_names
    ]

    if missing_sheets:
        raise ReconciliationError(
            "Thiếu sheet Packing List trùng tên: "
            + ", ".join(missing_sheets)
        )

    protection_records = read_protection_records(protection_bytes)
    scale_weights, scale_classes = read_scale_data(scale_bytes)

    (
        opm_035_po,
        opm_035_rtv,
        opm_031,
        opm_03g_po,
        opm_03g_rtv,
    ) = read_opm_data(opm_bytes)

    # Chỉ giữ sheet được chọn trước khi openpyxl đọc dữ liệu,
    # tránh nạp các sheet lớn không cần thiết vào bộ nhớ.
    selected_workbook = select_workbook_sheets(
        bbnt_bytes,
        selected_sheets,
    )

    try:
        workbook = load_workbook(
            BytesIO(selected_workbook),
            data_only=True,
            keep_links=False,
        )
    except Exception as exc:
        raise ReconciliationError(
            "Không mở được sheet BBNT đã chọn. Hãy kiểm tra file có bị khóa hoặc hỏng không."
        ) from exc

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    red_fill = PatternFill(
        start_color="FFFFC7CE",
        end_color="FFFFC7CE",
        fill_type="solid",
    )
    red_font = Font(color="9C0006", bold=True)

    total_rows = 0
    rows_with_issues = 0

    try:
        for sheet_name in selected_sheets:
            if sheet_name not in workbook.sheetnames:
                raise ReconciliationError(
                    f"Không tìm thấy sheet “{sheet_name}” trong file BBNT."
                )

            worksheet = workbook[sheet_name]
            packing_by_lot = read_packing_list(
                packing_bytes,
                sheet_name,
                packing_sheets,
            )

            header_row = find_bbnt_header_row(worksheet)
            data_rows = []

            for row_number in range(
                header_row + 1,
                (worksheet.max_row or header_row) + 1,
            ):
                coil_code = clean_id(
                    worksheet.cell(row=row_number, column=4).value
                )
                serial = clean_id(
                    worksheet.cell(row=row_number, column=1).value
                )

                if not coil_code:
                    continue
                if "TỔNG" in serial.upper() or "KẾT LUẬN" in serial.upper():
                    continue

                data_rows.append((row_number, coil_code))

            write_result_headers(worksheet, header_row)

            for row_number, coil_code in data_rows:
                key = normalize_lot(coil_code)

                bbnt_ticket = normalize_ticket(
                    worksheet.cell(row=row_number, column=5).value
                )
                net_document = as_number(
                    worksheet.cell(row=row_number, column=8).value
                )
                gross_document = as_number(
                    worksheet.cell(row=row_number, column=9).value
                )
                table_weight = as_number(
                    worksheet.cell(row=row_number, column=10).value
                )
                allowed_difference = round(
                    as_number(
                        worksheet.cell(row=row_number, column=12).value
                    ),
                    6,
                )

                protection_ticket, protection_weight = first_protection_match(
                    coil_code,
                    protection_records,
                )

                weight_type = scale_classes.get(bbnt_ticket, "Nhập hàng")
                net_packing = packing_by_lot.get(key, 0.0)
                packing_difference = net_document - net_packing

                ticket_matches = (
                    bbnt_ticket == protection_ticket
                    if protection_ticket
                    else False
                )

                ticket_for_scale = protection_ticket or bbnt_ticket
                scale_weight = scale_weights.get(ticket_for_scale, 0.0)

                scale_matches_protection = (
                    math.isclose(
                        scale_weight,
                        protection_weight,
                        abs_tol=1e-6,
                    )
                    if protection_weight > 0
                    else False
                )
                scale_matches_bbnt = (
                    math.isclose(
                        scale_weight,
                        table_weight,
                        abs_tol=1e-6,
                    )
                    if table_weight > 0
                    else False
                )

                weight_difference = scale_weight - gross_document
                tolerance_difference = (
                    weight_difference / gross_document
                    if gross_document != 0
                    else 0.0
                )
                tolerance_matches = (
                    round(tolerance_difference, 6) == allowed_difference
                )
                tolerance_status = (
                    "VUOT DUNG SAI"
                    if abs(tolerance_difference) > 0.003
                    else "OK"
                )

                received_035 = (
                    opm_035_po.get(key, 0.0)
                    - abs(opm_035_rtv.get(key, 0.0))
                )
                difference_035 = net_packing - received_035

                received_031 = opm_031.get(key, 0.0)
                difference_031 = (
                    0.0
                    if received_031 == 0
                    else net_packing - received_031
                )

                received_03g = (
                    opm_03g_po.get(key, 0.0)
                    - abs(opm_03g_rtv.get(key, 0.0))
                )
                difference_03g = net_packing - received_03g

                row_values = {
                    14: weight_type,
                    15: key,
                    16: net_packing,
                    17: packing_difference,
                    18: protection_weight,
                    19: protection_ticket,
                    20: ticket_matches,
                    21: scale_weight,
                    22: protection_ticket + "A" if protection_ticket else "",
                    23: scale_matches_protection,
                    24: scale_matches_bbnt,
                    25: weight_difference,
                    26: tolerance_difference,
                    27: tolerance_matches,
                    28: tolerance_status,
                    29: received_035,
                    30: difference_035,
                    31: received_031,
                    32: difference_031,
                    33: received_03g,
                    34: difference_03g,
                }

                row_has_issue = False

                for column, value in row_values.items():
                    cell = worksheet.cell(row=row_number, column=column)
                    cell.value = value
                    cell.border = thin_border

                    is_issue = False

                    if column in (20, 23, 24, 27) and value is False:
                        is_issue = True
                    elif column == 28 and value == "VUOT DUNG SAI":
                        is_issue = True
                    elif (
                        column in (17, 30)
                        and isinstance(value, (int, float))
                        and value != 0
                    ):
                        is_issue = True

                    if is_issue:
                        cell.fill = red_fill
                        cell.font = red_font
                        row_has_issue = True

                total_rows += 1
                rows_with_issues += int(row_has_issue)

        output = BytesIO()
        workbook.save(output)

        summary = {
            "sheets": len(selected_sheets),
            "rows": total_rows,
            "rows_with_issues": rows_with_issues,
        }
        return output.getvalue(), summary

    finally:
        workbook.close()


# -------------------- GIAO DIỆN STREAMLIT --------------------

st.set_page_config(
    page_title="Đối Chiếu Nghiệm Thu HRC",
    page_icon="📑",
    layout="wide",
)

st.title("Đối Chiếu Nghiệm Thu HRC")
st.write(
    "Tải lên 5 file Excel nguồn, chọn sheet BBNT cần đối chiếu, "
    "rồi tải file kết quả."
)

left, right = st.columns(2, gap="large")

with left:
    bbnt_file = st.file_uploader(
        "1. Biên bản nghiệm thu (BBNT)",
        type=["xlsx"],
        key="bbnt",
        help="Workbook có thể chứa nhiều sheet.",
    )
    packing_file = st.file_uploader(
        "2. Packing List nhà cung cấp",
        type=["xlsx"],
        key="packing",
    )
    scale_file = st.file_uploader(
        "3. Dữ liệu trạm cân",
        type=["xlsx"],
        key="scale",
    )

with right:
    protection_file = st.file_uploader(
        "4. Dữ liệu bảo vệ",
        type=["xlsx"],
        key="protection",
    )
    opm_file = st.file_uploader(
        "5. ERP OPM 105",
        type=["xlsx"],
        key="opm",
    )
    st.caption(
        "File được xử lý trong bộ nhớ của phiên làm việc; "
        "ứng dụng không ghi bản tải lên vào kho lưu trữ."
    )


selected_sheets = []
bbnt_bytes = bbnt_file.getvalue() if bbnt_file else None
packing_bytes = packing_file.getvalue() if packing_file else None

if bbnt_bytes and packing_bytes:
    try:
        bbnt_sheets = workbook_sheet_names(bbnt_bytes)
        packing_sheets = workbook_sheet_names(packing_bytes)

        packing_keys = {name.strip().casefold() for name in packing_sheets}
        matching_sheets = [
            name
            for name in bbnt_sheets
            if name.strip().casefold() in packing_keys
        ]

        if not matching_sheets:
            st.error(
                "Không tìm thấy sheet Packing List trùng tên với sheet BBNT. "
                "Hãy kiểm tra lại hai workbook."
            )
        else:
            selected_sheets = st.multiselect(
                "Chọn sheet BBNT cần đối chiếu",
                options=matching_sheets,
                default=[],
                help="Chỉ các sheet có tên trùng với Packing List mới được chọn.",
            )
    except Exception as exc:
        st.error(f"Không đọc được danh sách sheet: {exc}")
elif bbnt_bytes:
    try:
        bbnt_sheets = workbook_sheet_names(bbnt_bytes)
        selected_sheets = st.multiselect(
            "Chọn sheet BBNT cần đối chiếu",
            options=bbnt_sheets,
            default=[],
        )
    except Exception as exc:
        st.error(f"Không đọc được danh sách sheet BBNT: {exc}")
else:
    st.info("Tải file BBNT lên để hiển thị danh sách sheet.")


all_files_uploaded = all(
    file is not None
    for file in (
        bbnt_file,
        packing_file,
        scale_file,
        protection_file,
        opm_file,
    )
)

if not all_files_uploaded:
    st.caption("Cần đủ 5 file Excel để chạy đối chiếu.")

if st.button(
    "Bắt đầu đối chiếu",
    type="primary",
    use_container_width=True,
    disabled=not all_files_uploaded or not selected_sheets,
):
    try:
        with st.spinner(
            f"Đang đối chiếu {len(selected_sheets)} sheet..."
        ):
            result_bytes, summary = process_workbooks(
                bbnt_file=bbnt_bytes,
                packing_file=packing_bytes,
                scale_file=scale_file.getvalue(),
                protection_file=protection_file.getvalue(),
                opm_file=opm_file.getvalue(),
                selected_sheets=selected_sheets,
            )

        st.success(
            f"Đã đối chiếu {summary['rows']} dòng trong "
            f"{summary['sheets']} sheet."
        )

        if summary["rows_with_issues"]:
            st.warning(
                f"Có {summary['rows_with_issues']} dòng cần kiểm tra; "
                "các ô sai lệch đã được tô đỏ trong file kết quả."
            )
        else:
            st.info("Không phát hiện dòng cần cảnh báo theo các tiêu chí hiện tại.")

        st.download_button(
            "Tải file Excel kết quả",
            data=result_bytes,
            file_name="BBNT_Doi_Chieu_Ket_Qua.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    except ReconciliationError as exc:
        st.error(str(exc))
    except Exception as exc:
        st.error(
            "Có lỗi khi đối chiếu. Hãy kiểm tra các workbook và tiêu đề cột."
        )
        st.exception(exc)
```
