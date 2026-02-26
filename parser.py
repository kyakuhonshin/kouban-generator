import re
import csv
import io
from docx import Document

# ===== 時間語マッピング =====
TIME_MAP = {
    "明け方": "M",
    "朝": "M",
    "昼": "D",
    "夕方": "E",
    "夕": "E",
    "夜": "N",
    "深夜": "N",
}
TIME_KEYS = list(TIME_MAP.keys())

def normalize_numbers(text: str) -> str:
    """全角数字を半角へ"""
    return text.translate(str.maketrans("０１２３４５６７８９", "0123456789"))

def clean_spaces(text: str) -> str:
    return re.sub(r"[ \t]+", "", text)

def is_scene_heading(line: str) -> bool:
    """
    柱判定：以下のどちらか
    - 行頭が数字（全角/半角）
    - 行頭が '○'
    """
    s = line.strip()
    if not s:
        return False
    if s.startswith("○"):
        return True
    s2 = normalize_numbers(s)
    return re.match(r"^[0-9]+", s2) is not None

def extract_scene_no(line: str, fallback_no: int) -> int:
    """柱から番号抽出。取れなければfallback"""
    s = line.strip()
    if s.startswith("○"):
        return fallback_no
    s2 = normalize_numbers(s)
    m = re.match(r"^([0-9]+)", s2)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return fallback_no
    return fallback_no

def extract_dn(line: str) -> str:
    """柱から M/D/E/N を抽出。見つからなければ空欄"""
    s = line.strip()
    for key in TIME_KEYS:
        if key in s:
            return TIME_MAP[key]
    return ""

def extract_location(line: str) -> str:
    """
    柱から場所文字列を抽出
    - 数字始まりの場合：先頭数字を落とす
    - ○始まりの場合：先頭○を落とす
    - 時間語があれば、その手前を場所とする
    """
    s = line.strip()
    if s.startswith("○"):
        s = s.lstrip("○").strip()
    else:
        s = normalize_numbers(s)
        s = re.sub(r"^[0-9]+\s*", "", s).strip()
    for key in TIME_KEYS:
        if key in s:
            before = s.split(key)[0]
            return before.strip("（）() \t")
    return s.strip("（）() \t")

def is_dialogue_line(line: str) -> bool:
    """セリフ行：先頭〜12文字程度 + 「 で始まる"""
    return re.match(r"^(.{1,12})「", line.strip()) is not None

def extract_character_from_dialogue(line: str):
    """セリフ行から役名抽出"""
    m = re.match(r"^(.{1,12})「", line.strip())
    if not m:
        return None
    name = m.group(1).strip()
    name = clean_spaces(name)
    name = name.strip("：:・")
    return name or None

def summarize_scene(scene_lines, max_chars=40) -> str:
    """簡易要約（LLMなし）"""
    non_dialogue = [l.strip() for l in scene_lines if l.strip() and not is_dialogue_line(l)]
    base = "".join(non_dialogue[:3]).replace("\n", "").strip()
    if not base:
        base = "".join([l.strip() for l in scene_lines[:1]])
    if len(base) <= max_chars:
        return base
    return base[:max_chars-1] + "…"

def parse_docx_to_csv(docx_path: str) -> str:
    """docxをパースしてCSV文字列を返す"""
    doc = Document(docx_path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    
    scenes = []
    current = None
    auto_no = 1
    
    for line in lines:
        if is_scene_heading(line):
            if current:
                scenes.append(current)
            scene_no = extract_scene_no(line, auto_no)
            location = extract_location(line)
            dn = extract_dn(line)
            current = {
                "scene_no": scene_no,
                "location": location,
                "DN": dn,
                "lines": [],
                "dialogue_characters": set(),
                "all_characters": set(),
            }
            auto_no += 1
        else:
            if current:
                current["lines"].append(line)
                c = extract_character_from_dialogue(line)
                if c:
                    current["dialogue_characters"].add(c)
    
    if current:
        scenes.append(current)
    
    # 全人物一覧（セリフベース）
    all_chars = []
    for s in scenes:
        for c in s["dialogue_characters"]:
            if c not in all_chars:
                all_chars.append(c)
    
    # ト書き登場も◯（名前長>=2のみ）
    for s in scenes:
        s["all_characters"] = set(s["dialogue_characters"])
        for line in s["lines"]:
            if is_dialogue_line(line):
                continue
            for c in all_chars:
                if len(c) < 2:
                    continue
                if c in line:
                    s["all_characters"].add(c)
    
    # CSV出力（メモリ上）
    output = io.StringIO()
    w = csv.writer(output)
    header = ["scene_no", "location", "D/N", "summary"] + all_chars + ["props_art", "notes"]
    w.writerow(header)
    
    for s in scenes:
        summary = summarize_scene(s["lines"], max_chars=40)
        row = [s["scene_no"], s["location"], s["DN"], summary]
        for c in all_chars:
            row.append("◯" if c in s["all_characters"] else "")
        row += ["", ""]
        w.writerow(row)
    
    return output.getvalue()
