#!/usr/bin/env python3

from __future__ import annotations

import json
import posixpath
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile


ROOT = Path("/home/korot/projects/storyboard")
WORKBOOK_PATH = ROOT / "Режиссёрский сценарий - Кулинар Вкус крови_4.xlsx"
OUTPUT_JSON = ROOT / "revision-board-data.json"
OUTPUT_IMAGE_DIR = ROOT / "assets" / "revision-board" / "excel"
THUMBNAIL_DIR = ROOT / "assets" / "revision-board" / "thumbnails"
GENERATED_IMAGE_MAP = {
    75: [
        "assets/revision-board/generated/scene-8-shot-10-ai-v1.png",
        "assets/revision-board/generated/scene-8-shot-10-tail-ai-v1.png",
        "assets/revision-board/generated/scene-8-shot-10-cat-reaction-ai-v1.png",
    ],
}

NS = {
    "sheet": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "draw": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


@dataclass(frozen=True)
class ManualNote:
    priority: str
    complexity: str
    status: str
    transcript_ranges: tuple[tuple[int, int], ...]
    cleaned_note: str
    raw_override: str | None = None


MANUAL_NOTES: dict[int, ManualNote] = {
    4: ManualNote(
        priority="high",
        complexity="high",
        status="нужен новый проход",
        transcript_ranges=((5, 14),),
        cleaned_note=(
            "Нужен длинный отъезд: макродеталь циферблата -> часы -> стена с рисунком и календарём -> "
            "общий план следователя. Важно правильно показать географию кабинета и посадку Коли на углу стола."
        ),
    ),
    6: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((29, 46),),
        cleaned_note=(
            "Сохранить направление, но подчистить географию: Коля должен быть ближе к камере и сидеть на углу стола."
        ),
    ),
    7: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((29, 46),),
        cleaned_note=(
            "Сделать чистую восьмёрку на Колю без плеча следователя. Кадр должен ощущаться как POV следователя."
        ),
    ),
    8: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((29, 46),),
        cleaned_note=(
            "Сделать более крупный анфасный кадр на следователя, без плеча Коли в кадре."
        ),
    ),
    9: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужен новый кадр",
        transcript_ranges=((49, 58),),
        cleaned_note=(
            "Добавить отсутствующий кадр: непрерывная POV-панорама Коли по кабинету с ключевыми деталями интерьера."
        ),
    ),
    10: ManualNote(
        priority="medium",
        complexity="low",
        status="нужно поправить",
        transcript_ranges=((61, 66),),
        cleaned_note=(
            "Крупный кадр на следователя из-под ноутбука. Нижняя полоса в кадре должна читаться как край ноутбука."
        ),
    ),
    17: ManualNote(
        priority="high",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((77, 94),),
        cleaned_note=(
            "Сцена 2, кадр 4: сохранить географию, но добавить читаемый отъезд от крупного профиля Коли к общему плану со спины."
        ),
    ),
    21: ManualNote(
        priority="low",
        complexity="low",
        status="под вопросом",
        transcript_ranges=((93, 96),),
        cleaned_note=(
            "Кадр пока не финализирован: возможен вылет из сцены. Оставлен на доске как спорный."
        ),
    ),
    22: ManualNote(
        priority="high",
        complexity="high",
        status="нужен новый проход",
        transcript_ranges=((97, 118),),
        cleaned_note=(
            "Зеркалить направление движения. Кадр должен быть проходом MOVI с отъездом от среднего к общему, "
            "с входом в зал и видимым столиком в дальнем углу."
        ),
    ),
    25: ManualNote(
        priority="high",
        complexity="high",
        status="нужен новый проход",
        transcript_ranges=((117, 166),),
        cleaned_note=(
            "Сложный подвижный кадр с серией перебросок внимания внутри одного прохода: Коля -> Дмитрий -> Коля -> пустое кресло -> "
            "поиск Дмитрия -> движение к проходу. Нужно собрать в одну ясную хореографию."
        ),
    ),
    27: ManualNote(
        priority="high",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((169, 182),),
        cleaned_note=(
            "Разговорная восьмёрка Дмитрия и Коли должна происходить в проходе/на пороге, а не уже за столом."
        ),
    ),
    28: ManualNote(
        priority="high",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((183, 206),),
        cleaned_note=(
            "Восьмёрка Сони и Луки тоже остаётся у порога. Нужен маленький наезд к Соне на момент гипноза."
        ),
    ),
    29: ManualNote(
        priority="high",
        complexity="high",
        status="нужен новый проход",
        transcript_ranges=((209, 248),),
        cleaned_note=(
            "Собрать длинный вход Якова: из тёмного коридора через стол к посадке. Важно поправить посадку за столом: кот на самом краю."
        ),
    ),
    31: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((118, 120),),
        cleaned_note=(
            "Крупный на кота оставить по смыслу, но поменять направление камеры."
        ),
    ),
    33: ManualNote(
        priority="high",
        complexity="high",
        status="нужен новый кадр",
        transcript_ranges=((120, 156),),
        cleaned_note=(
            "Построить общий кадр сцены 4 как оммаж «Тайной вечере» с точной рассадкой всех персонажей."
        ),
    ),
    36: ManualNote(
        priority="medium",
        complexity="low",
        status="нужно уточнить",
        transcript_ranges=((155, 156),),
        cleaned_note=(
            "Крупный кадр-пара к Коле. Пока рабочее решение: делать на Дмитрия."
        ),
    ),
    38: ManualNote(
        priority="high",
        complexity="high",
        status="нужен новый проход",
        transcript_ranges=((158, 170),),
        cleaned_note=(
            "Собрать длинный проход: крупно тумбочка и чашка -> следователь с кружкой -> разворот против часовой стрелки -> выход к Коле."
        ),
    ),
    39: ManualNote(
        priority="high",
        complexity="high",
        status="нужен новый проход",
        transcript_ranges=((174, 194),),
        cleaned_note=(
            "Построить проход с тремя фазами: наседание на Колю, отход к столу, посадка за стол и первое поедание конфеты."
        ),
    ),
    41: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((196, 204),),
        cleaned_note=(
            "Поменять направление старта панорамы: кот отдельно, без лишних персонажей в первом фрагменте движения."
        ),
    ),
    47: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((228, 235),),
        cleaned_note=(
            "Смена оси: Коля справа, Яков слева. Оставить конфликт через стол и убрать лишних людей."
        ),
    ),
    49: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((238, 240),),
        cleaned_note=(
            "Нужен отъезд от Коли с выходом на силуэт/фигуру Сони за его плечом."
        ),
    ),
    51: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((240, 242),),
        cleaned_note=(
            "Крупный портретный план на Луку с лёгкой нижней точки, без книги и без рук на столе."
        ),
    ),
    52: ManualNote(
        priority="medium",
        complexity="low",
        status="нужно поправить",
        transcript_ranges=((242, 242),),
        cleaned_note=(
            "Кадр повторяет базовую композицию 6.9: Коля в кадре, лицо Сони не видно, только охраняющий силуэт сзади."
        ),
    ),
    55: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((248, 248),),
        cleaned_note=(
            "Сделать чуть более общий повтор композиции 6.12 и открыть лицо Сони, когда Коля поднимает на неё глаза."
        ),
    ),
    59: ManualNote(
        priority="high",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((252, 254),),
        cleaned_note=(
            "Поправить рассадку и панораму: старт на Якове/Филиппе, затем по часовой стрелке на Соню и Колю."
        ),
    ),
    60: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((256, 260),),
        cleaned_note=(
            "Кадр должен работать как переброс с Якова на Дмитрия; Дмитрий не уходит с позиции за Лукой."
        ),
    ),
    65: ManualNote(
        priority="high",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((264, 270),),
        cleaned_note=(
            "Сцена 8, кадр 1: поправить географию за столом и построить отъезд от лица Якова к группе из четырёх вампиров."
        ),
    ),
    66: ManualNote(
        priority="high",
        complexity="high",
        status="нужен новый проход",
        transcript_ranges=((280, 290),),
        cleaned_note=(
            "Собрать весь экшен в один фронтальный отъезд без монтажных перебросок: четыре действия читаются последовательно внутри одного прохода."
        ),
    ),
    67: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((290, 290),),
        cleaned_note=(
            "Сохранить смысл кадра, но сделать его чуть крупнее и сократить число фигур в фоне."
        ),
    ),
    68: ManualNote(
        priority="high",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((292, 298),),
        cleaned_note=(
            "Чёткая восьмёрка из-за плеча Дмитрия: Коли нет в кадре, Яков в центре под столом, к нему подступают Соня и Филипп."
        ),
    ),
    69: ManualNote(
        priority="high",
        complexity="medium",
        status="нужен новый кадр",
        transcript_ranges=((302, 306),),
        cleaned_note=(
            "Добавить следующий beat после приземления: Яков вылезает из-под стола и хватает ножки стула."
        ),
    ),
    70: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((306, 306),),
        cleaned_note=(
            "Крупный анфасный защитный кадр на Дмитрия; в фоне у стены читается растерянный Коля."
        ),
    ),
    71: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((308, 314),),
        cleaned_note=(
            "Опустить камеру ниже, в настоящий POV Якова из-под стола, и добавить недостающих персонажей по географии."
        ),
    ),
    72: ManualNote(
        priority="high",
        complexity="high",
        status="нужен новый проход",
        transcript_ranges=((316, 328),),
        cleaned_note=(
            "Построить профильный кадр-стенку 3х3 с подъездом по камере и большой глубиной пространства."
        ),
    ),
    73: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((330, 332),),
        cleaned_note=(
            "Сделать общий по пояс на сторону Дмитрия/Луки/Графа как пару к следующему общему кадру."
        ),
    ),
    74: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((330, 332),),
        cleaned_note=(
            "Сделать парный общий кадр на противоположную сторону стола."
        ),
    ),
    75: ManualNote(
        priority="high",
        complexity="low",
        status="нужен новый кадр",
        transcript_ranges=((334, 335),),
        cleaned_note=(
            "Добавить отсутствующий крупный insert на ногу Дмитрия и хвост/кота."
        ),
        raw_override=(
            "334. [01:29:03.613] — Марина\n"
            "Спасибо. Десятый кадр. У тебя, по-моему, он тоже был. А, нет, не было. "
            "Десятый — это крупный на ногу Дмитрия, который наступает на кота.\n"
            "335. [01:29:15.293] — Раскадровщик\n"
            "Ок."
        ),
    ),
    76: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((336, 336),),
        cleaned_note=(
            "Перевести прыжок Якова в анфасный POV-кадр, где он летит прямо на камеру."
        ),
    ),
    77: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((336, 336),),
        cleaned_note=(
            "Сделать анфасный встречный прыжок Дмитрия на камеру."
        ),
    ),
    78: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((336, 336),),
        cleaned_note=(
            "Столкновение Дмитрия и Якова держать в той же оси и крупности, что и у профильной стенки 3х3."
        ),
    ),
    79: ManualNote(
        priority="low",
        complexity="low",
        status="нужно подчистить",
        transcript_ranges=((336, 338),),
        cleaned_note=(
            "Оставить фехтование Филиппа с вилкой, но почистить фон и лишних людей."
        ),
    ),
    80: ManualNote(
        priority="low",
        complexity="low",
        status="нужно подчистить",
        transcript_ranges=((336, 338),),
        cleaned_note=(
            "Оставить как рабочий дубль боя Луки и Филиппа; позже возможен merge с предыдущим кадром."
        ),
    ),
    82: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((340, 340),),
        cleaned_note=(
            "Сделать реакцию Коли крупным планом, а не со спины."
        ),
    ),
    83: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((342, 342),),
        cleaned_note=(
            "Уточнить конец phone-shot: в камеру летят одежда Графа и сам кот, затем уход в затемнение."
        ),
    ),
    85: ManualNote(
        priority="medium",
        complexity="medium",
        status="нужно поправить",
        transcript_ranges=((342, 354),),
        cleaned_note=(
            "Сцена 8-1, кадр 1: POV после драки упростить, сократить лишние фазы и яснее показать маршрут к проходу."
        ),
    ),
    86: ManualNote(
        priority="low",
        complexity="low",
        status="нужно поправить",
        transcript_ranges=((353, 354),),
        cleaned_note=(
            "Оставить кадр почти как есть, но уложить Дмитрия в корректную позицию."
        ),
    ),
    94: ManualNote(
        priority="low",
        complexity="low",
        status="нужно поправить",
        transcript_ranges=((712, 722),),
        cleaned_note=(
            "Крупный на следователя сверху, лежащего на спине."
        ),
    ),
    95: ManualNote(
        priority="high",
        complexity="high",
        status="нужен новый проход",
        transcript_ranges=((722, 746),),
        cleaned_note=(
            "Собрать подвижный кадр сцены 9: Коля встаёт, обходит стол, удаляет протокол с ноутбука и забирает коробку конфет."
        ),
    ),
}


def extract_transcript_paragraphs(docx_path: Path) -> list[str]:
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(docx_path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))

    paragraphs: list[str] = []
    for paragraph in root.findall(".//w:p", ns):
        text = "".join((node.text or "") for node in paragraph.findall(".//w:t", ns)).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


TRANSCRIPT_PARAGRAPHS = extract_transcript_paragraphs(ROOT / "2026.06.01 по раскадровкам.docx")


def transcript_text(ranges: tuple[tuple[int, int], ...]) -> str:
    chunks: list[str] = []
    for start, end in ranges:
        for index in range(start, end + 1):
            chunks.append(TRANSCRIPT_PARAGRAPHS[index - 1])
    return "\n".join(chunks)


def ensure_thumbnail(source_rel: str) -> str:
    source = ROOT / source_rel
    category = Path(source_rel).parts[2] if len(Path(source_rel).parts) > 2 else "misc"
    thumb_dir = THUMBNAIL_DIR / category
    thumb_dir.mkdir(parents=True, exist_ok=True)
    thumb_name = f"{Path(source_rel).stem}.jpg"
    thumb_path = thumb_dir / thumb_name

    if thumb_path.exists() and thumb_path.stat().st_mtime >= source.stat().st_mtime:
        return str(thumb_path.relative_to(ROOT))

    subprocess.run(
        [
            "ffmpeg",
            "-y",
            "-i",
            str(source),
            "-vf",
            "scale=960:-1:force_original_aspect_ratio=decrease",
            "-q:v",
            "4",
            str(thumb_path),
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    return str(thumb_path.relative_to(ROOT))


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str | None:
    value = cell.find("sheet:v", NS)
    if value is None:
        return None
    text = value.text or ""
    if cell.attrib.get("t") == "s":
        return shared_strings[int(text)]
    return text


def parse_shared_strings(archive: ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    shared: list[str] = []
    for item in root.findall("sheet:si", NS):
        parts = [node.text or "" for node in item.findall(".//sheet:t", NS)]
        shared.append("".join(parts))
    return shared


def load_sheet_path(archive: ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    sheets = workbook.find("sheet:sheets", NS)
    if sheets is None:
        raise KeyError("Workbook has no sheets node")
    for sheet in sheets:
        if sheet.attrib["name"] == sheet_name:
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            return "xl/" + rel_map[rid]
    raise KeyError(f"Sheet {sheet_name!r} not found")


def parse_sheet_rows(archive: ZipFile, sheet_path: str, shared_strings: list[str]) -> tuple[dict[int, dict[str, str]], dict[int, str]]:
    sheet_root = ET.fromstring(archive.read(sheet_path))
    rows: dict[int, dict[str, str]] = {}
    scene_labels: dict[int, str] = {}
    current_scene = ""

    sheet_data = sheet_root.find("sheet:sheetData", NS)
    if sheet_data is None:
        return rows, scene_labels

    for row in sheet_data:
        row_number = int(row.attrib["r"])
        values: dict[str, str] = {}
        for cell in row.findall("sheet:c", NS):
            column = re.match(r"([A-Z]+)", cell.attrib["r"]).group(1)
            text = cell_text(cell, shared_strings)
            if text is not None:
                values[column] = text
        if values:
            rows[row_number] = values
        marker = values.get("A", "")
        if marker.startswith("СЦЕНА "):
            current_scene = marker.split("—", 1)[0].strip()
        if current_scene:
            scene_labels[row_number] = current_scene

    return rows, scene_labels


def parse_row_images(archive: ZipFile, sheet_path: str) -> dict[int, list[str]]:
    sheet_root = ET.fromstring(archive.read(sheet_path))
    drawing_ref = sheet_root.find("sheet:drawing", NS)
    if drawing_ref is None:
        return {}

    rel_id = drawing_ref.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
    rels_path = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
    rels_root = ET.fromstring(archive.read(rels_path))

    drawing_target = ""
    for rel in rels_root:
        if rel.attrib["Id"] == rel_id:
            drawing_target = posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), rel.attrib["Target"]))
            break
    if not drawing_target:
        return {}

    drawing_root = ET.fromstring(archive.read(drawing_target))
    drawing_rels_path = posixpath.join(
        posixpath.dirname(drawing_target),
        "_rels",
        posixpath.basename(drawing_target) + ".rels",
    )
    drawing_rels_root = ET.fromstring(archive.read(drawing_rels_path))
    image_rel_map = {
        rel.attrib["Id"]: posixpath.normpath(posixpath.join(posixpath.dirname(drawing_target), rel.attrib["Target"]))
        for rel in drawing_rels_root
    }

    row_images: dict[int, list[str]] = {}
    for anchor in drawing_root:
        pic = anchor.find(".//draw:pic", NS)
        marker = anchor.find("draw:from", NS)
        if pic is None or marker is None:
            continue
        row_number = int(marker.find("draw:row", NS).text) + 1
        blip = pic.find(".//a:blip", NS)
        if blip is None:
            continue
        embed_id = blip.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"]
        image_path = image_rel_map.get(embed_id)
        if image_path is None:
            continue
        row_images.setdefault(row_number, []).append(image_path)
    return row_images


def slugify(text: str) -> str:
    lowered = text.lower()
    lowered = lowered.replace("сцена ", "scene-")
    lowered = lowered.replace("кадр ", "shot-")
    lowered = lowered.replace("сцена-", "scene-")
    lowered = lowered.replace("кадр-", "shot-")
    lowered = lowered.replace("сцена", "scene")
    lowered = lowered.replace("кадр", "shot")
    lowered = lowered.replace(" ", "-")
    lowered = lowered.replace(".", "-")
    lowered = lowered.replace("/", "-")
    lowered = re.sub(r"-+", "-", lowered)
    return lowered.strip("-")


def normalize_scene(scene_label: str) -> str:
    match = re.search(r"СЦЕНА\s+([0-9]+(?:-[0-9]+)?)", scene_label)
    if not match:
        return scene_label.title()
    value = match.group(1).replace("-", ".")
    return f"Сцена {value}"


def normalize_shot(shot_value: str) -> str:
    text = shot_value.strip()
    if text.upper() == "КАДР ЗАКАТА  - ФУТАЖ":
        return "Футаж заката"
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return f"Кадр {text}"


def build_entries() -> list[dict[str, Any]]:
    OUTPUT_IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    THUMBNAIL_DIR.mkdir(parents=True, exist_ok=True)

    with ZipFile(WORKBOOK_PATH) as archive:
        shared_strings = parse_shared_strings(archive)
        sheet_path = load_sheet_path(archive, "ОБЩИЙ ДОК")
        rows, scene_labels = parse_sheet_rows(archive, sheet_path, shared_strings)
        row_images = parse_row_images(archive, sheet_path)

        exported: dict[str, str] = {}
        entries: list[dict[str, Any]] = []

        for row_number, manual in sorted(MANUAL_NOTES.items()):
            row = rows[row_number]
            scene_label = normalize_scene(scene_labels.get(row_number, ""))
            shot_label = normalize_shot(row.get("B") or row.get("A") or str(row_number))
            scene_slug = slugify(scene_label)
            shot_slug = slugify(shot_label)
            image_paths = row_images.get(row_number, [])
            public_images: list[str] = []
            public_image_thumbs: list[str] = []

            for index, image_path in enumerate(image_paths, start=1):
                if image_path not in exported:
                    suffix = Path(image_path).suffix or ".png"
                    filename = f"{scene_slug}-{shot_slug}"
                    if len(image_paths) > 1:
                        filename += f"-ref-{index}"
                    filename += suffix
                    destination = OUTPUT_IMAGE_DIR / filename
                    destination.write_bytes(archive.read(image_path))
                    exported[image_path] = str(destination.relative_to(ROOT))
                public_images.append(exported[image_path])
                public_image_thumbs.append(ensure_thumbnail(exported[image_path]))

            generated_images = GENERATED_IMAGE_MAP.get(row_number, [])
            generated_thumbs = [ensure_thumbnail(path) for path in generated_images]

            entries.append(
                {
                    "row": row_number,
                    "scene": scene_label,
                    "shot": shot_label,
                    "title": f"{scene_label} · {shot_label}",
                    "status": manual.status,
                    "priority": manual.priority,
                    "complexity": manual.complexity,
                    "originalDescription": (row.get("E") or "").strip(),
                    "rawTranscriptExcerpt": manual.raw_override or transcript_text(manual.transcript_ranges),
                    "cleanedDirection": manual.cleaned_note,
                    "excelImages": public_images,
                    "excelThumbnails": public_image_thumbs,
                    "generatedImages": generated_images,
                    "generatedThumbnails": generated_thumbs,
                    "sourceSheet": "ОБЩИЙ ДОК",
                    "placeholderLabel": "Место для будущей AI-генерации",
                }
            )

    return entries


def main() -> None:
    entries = build_entries()
    payload = {
        "meta": {
            "title": "Доработка раскадровок",
            "subtitle": "Рабочий реестр кадров из листа «ОБЩИЙ ДОК», которым нужны правки после встречи 1 июня 2026 года.",
            "total": len(entries),
            "sheet": "ОБЩИЙ ДОК",
        },
        "entries": entries,
    }
    OUTPUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {OUTPUT_JSON}")
    print(f"Entries: {len(entries)}")


if __name__ == "__main__":
    main()
