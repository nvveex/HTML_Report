#!/usr/bin/env python3
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile


NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


@dataclass
class SheetData:
    name: str
    rows: list[list[str]]


def _col_to_idx(ref: str) -> int:
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _load_shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    out: list[str] = []
    for si in root.findall("a:si", NS):
        texts = [t.text or "" for t in si.iterfind(".//a:t", NS)]
        out.append("".join(texts))
    return out


def _sheet_targets(zf: ZipFile) -> list[tuple[str, str]]:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    sheets: list[tuple[str, str]] = []
    for sheet in wb.find("a:sheets", NS):
        rid = sheet.attrib["{" + NS["r"] + "}id"]
        sheets.append((sheet.attrib.get("name", "Sheet1"), rid_to_target[rid]))
    return sheets


def read_workbook(path: str | Path, max_rows: int | None = None) -> list[SheetData]:
    path = Path(path)
    with ZipFile(path) as zf:
        shared_strings = _load_shared_strings(zf)
        sheets = _sheet_targets(zf)
        out: list[SheetData] = []
        for sheet_name, target in sheets:
            xml = ET.fromstring(zf.read("xl/" + target))
            rows: list[list[str]] = []
            for row_idx, row in enumerate(xml.findall(".//a:sheetData/a:row", NS), start=1):
                if max_rows is not None and row_idx > max_rows:
                    break
                values: dict[int, str] = {}
                for cell in row.findall("a:c", NS):
                    ref = cell.attrib.get("r", "")
                    col_idx = _col_to_idx("".join(ch for ch in ref if ch.isalpha()))
                    value_node = cell.find("a:v", NS)
                    inline_node = cell.find("a:is", NS)
                    value = ""
                    if inline_node is not None:
                        value = "".join(t.text or "" for t in inline_node.iterfind(".//a:t", NS))
                    elif value_node is not None and value_node.text is not None:
                        raw = value_node.text
                        if cell.attrib.get("t") == "s" and raw.isdigit():
                            idx = int(raw)
                            value = shared_strings[idx] if idx < len(shared_strings) else raw
                        else:
                            value = raw
                    values[col_idx] = value.strip()
                if values:
                    max_col = max(values)
                    rows.append([values.get(i, "") for i in range(1, max_col + 1)])
            out.append(SheetData(name=sheet_name, rows=rows))
        return out


def read_result_rows(path: str | Path) -> list[dict[str, str]]:
    sheets = read_workbook(path)
    if not sheets:
        return []
    target = next((s for s in sheets if s.name == "result"), sheets[0])
    if not target.rows:
        return []
    headers = [h.strip() for h in target.rows[0]]
    records: list[dict[str, str]] = []
    for raw in target.rows[1:]:
        if not any(cell.strip() for cell in raw):
            continue
        record: dict[str, str] = {}
        for idx, header in enumerate(headers):
            if header == "":
                continue
            record[header] = raw[idx].strip() if idx < len(raw) else ""
        records.append(record)
    return records


def maybe_number(value: Any) -> float | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    s = s.replace(",", "").replace("%", "")
    try:
        return float(s)
    except ValueError:
        return None


def maybe_percent(value: Any) -> float | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    if s.endswith("%"):
        try:
            return float(s[:-1]) / 100.0
        except ValueError:
            return None
    num = maybe_number(s)
    return None if num is None else num
