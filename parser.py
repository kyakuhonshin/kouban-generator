import re
import csv
import io
from docx import Document

# ===== 時間語マッピング =====
TIME_MAP = {
    "明け方": "M",
    "朝": "M",
    "早朝": "M",
    "午前": "M",
    "昼": "D",
    "午後": "D",
    "日中": "D",
    "夕方": "E",
    "夕": "E",
    "黄昏": "E",
    "夜": "N",
    "深夜": "N",
    "M": "M",
    "D": "D",
    "E": "E",
    "N": "N",
}
TIME_KEYS = list(TIME_MAP.keys())

def normalize_numbers(text: str) -> str:
    """全角数字を半角へ"""
    return text.translate(str.maketrans("０１２３４５６７８９", "0123456789"))

def clean_spaces(text: str) -> str:
    """全種類の空白を削除（半角スペース、タブ、全角スペース）"""
    return re.sub(r"[ \t\u3000]+", "", text)

def normalize_character_name(name: str) -> str:
    """
    人物名の正規化：
    - 前後空白トリム
    - 文中の空白を全削除（半角・全角）
    - 末尾の括弧注釈を削除（(off),（声）等）
    - 大文字小文字統一（英語名の場合）
    """
    if not name:
        return ""
    
    # 空白削除
    normalized = clean_spaces(name.strip())
    
    # 末尾の括弧注釈を削除
    normalized = re.sub(r"[（(][^）)]+[）)]", "", normalized)
    
    # 末尾の制作付記を削除（M, N, off, 声等）
    normalized = re.sub(r"(off|OFF|声|OS|ナレーション|M|N)$", "", normalized, flags=re.IGNORECASE)
    
    # 英語名は小文字に統一
    if re.match(r"^[a-zA-Z0-9]+$", normalized):
        normalized = normalized.lower()
    
    return normalized.strip()

def is_valid_character_name(name: str) -> bool:
    """人物名として妥当か（1文字は除外）"""
    normalized = normalize_character_name(name)
    return len(normalized) > 1

def is_scene_heading(line: str) -> bool:
    """
    柱判定：以下のいずれか
    - 行頭が S#数字（新形式）
    - 行頭が数字（全角/半角）
    - 行頭が '○'
    """
    s = line.strip()
    if not s:
        return False
    # S#形式（新）
    if re.match(r"^S#\s*\d+", s, re.IGNORECASE):
        return True
    if s.startswith("○"):
        return True
    s2 = normalize_numbers(s)
    return re.match(r"^[0-9]+", s2) is not None

def extract_scene_no(line: str, fallback_no: int) -> int:
    """柱から番号抽出。取れなければfallback"""
    s = line.strip()
    # S#形式を優先
    m = re.match(r"^S#\s*(\d+)", s, re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return fallback_no
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
    
    # S#形式の場合：・区切りで最後の要素をチェック
    if re.match(r"^S#", s, re.IGNORECASE):
        parts = s.split("・")
        if len(parts) > 1:
            last_part = parts[-1].strip()
            if last_part in TIME_MAP:
                return TIME_MAP[last_part]
    
    for key in TIME_KEYS:
        if key in s:
            return TIME_MAP[key]
    return ""

def extract_location(line: str) -> str:
    """
    柱から場所文字列を抽出
    - S#形式：S#数字を除去し、・区切りで時間語を除いた部分
    - 数字始まりの場合：先頭数字を落とす
    - ○始まりの場合：先頭○を落とす
    - 時間語があれば、その手前を場所とする
    """
    s = line.strip()
    
    # S#形式
    m = re.match(r"^S#\s*\d+\s*(.+)", s, re.IGNORECASE)
    if m:
        rest = m.group(1)
        parts = rest.split("・")
        # 最後が時間語なら除外
        if parts[-1].strip() in TIME_MAP:
            parts = parts[:-1]
        return "・".join(parts).strip()
    
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

def extract_character_table(lines: list) -> list:
    """
    【人物表】セクションから人物リストを抽出
    ・金沢 和希（27）・・・ のような形式から「金沢和希」を抽出
    """
    characters = []
    in_table = False
    empty_count = 0
    
    for line in lines:
        stripped = line.strip()
        
        # セクション開始
        if re.match(r"^【人物表】", stripped):
            in_table = True
            empty_count = 0
            continue
        
        if not in_table:
            continue
        
        # セクション終了条件
        if re.match(r"^(S#|【|本編|脚本|タイトル|シーン|場面)", stripped):
            break
        if not stripped:
            empty_count += 1
            if empty_count >= 2:
                break
            continue
        else:
            empty_count = 0
        
        # 人物行の抽出（先頭が「・」）
        if stripped.startswith("・"):
            # 「・」を除去
            name = stripped[1:].strip()
            # 「（」または「(」以降を削除（年齢や注釈）
            name = re.split(r"[（(]", name)[0].strip()
            # 末尾の「・・・」等を除去
            name = re.sub(r"[・\.]+$", "", name).strip()
            
            normalized = normalize_character_name(name)
            if is_valid_character_name(normalized) and normalized not in characters:
                characters.append(normalized)
    
    return characters

def is_dialogue_line(line: str) -> bool:
    """セリフ行：先頭〜12文字程度 + 「 または 『 で始まる"""
    return re.match(r"^(.{1,12})[「『]", line.strip()) is not None

def extract_character_from_dialogue(line: str):
    """
    セリフ行から役名抽出（強化版）
    - 「名前（off）「セリフ」」形式に対応
    - 括弧注釈を除去して正規化
    """
    m = re.match(r"^(.{1,12})[「『]", line.strip())
    if not m:
        return None
    name = m.group(1).strip()
    # 正規化（空白削除・括弧注釈除去）
    name = normalize_character_name(name)
    return name or None

def extract_speaker_from_line(line: str) -> str:
    """
    行から話者名を抽出（括弧注釈対応）
    mito(off)「〜」や 和希（声）「〜」に対応
    """
    # 「 または 『 の前の部分を話者候補として取得
    m = re.match(r"^(.+?)[「『]", line.strip())
    if not m:
        return None
    speaker_raw = m.group(1).strip()
    # 正規化
    normalized = normalize_character_name(speaker_raw)
    if is_valid_character_name(normalized):
        return normalized
    return None

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
    """docxをパースしてCSV文字列を返す（【人物表】対応版）"""
    doc = Document(docx_path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    
    # 【人物表】から人物マスターを抽出（最優先）
    char_table = extract_character_table(lines)
    
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
                # 強化された話者抽出
                speaker = extract_speaker_from_line(line)
                if speaker:
                    current["dialogue_characters"].add(speaker)
    
    if current:
        scenes.append(current)
    
    # 全人物一覧（人物表ベース > セリフベース）
    all_chars = char_table.copy() if char_table else []
    
    # セリフから抽出した人物を追加
    for s in scenes:
        for c in s["dialogue_characters"]:
            normalized = normalize_character_name(c)
            if normalized not in all_chars and is_valid_character_name(normalized):
                all_chars.append(normalized)
    
    # ト書き登場も◯（名前長>=2のみ）
    for s in scenes:
        s["all_characters"] = set(s["dialogue_characters"])
        for line in s["lines"]:
            if is_dialogue_line(line):
                continue
            for c in all_chars:
                if len(c) < 2:
                    continue
                # 正規化した名前で照合
                normalized_c = normalize_character_name(c)
                if normalized_c in line or c in line:
                    s["all_characters"].add(normalized_c)
    
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
