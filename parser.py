"""
香盤ジェネレーター - パーサーモジュール
v0.9.3 - モード分離版（Mode-S:標準 / Mode-L:区切り行）
"""
import re
import csv
import io
from docx import Document
from datetime import datetime

# ===== バージョン =====
PARSER_VERSION = "v0.9.3"

# ===== 時間語マッピング =====
TIME_MAP = {
    "明け方": "M", "朝": "M", "早朝": "M", "午前": "M",
    "昼": "D", "午後": "D", "日中": "D",
    "夕方": "E", "夕": "E", "夕刻": "E", "黄昏": "E",
    "夜": "N", "深夜": "N", "翌朝": "M",
    "M": "M", "D": "D", "E": "E", "N": "N",
}
TIME_KEYS = list(TIME_MAP.keys())

# ===== ブラックリスト =====
SPEAKER_BLACKLIST = {
    "テロップ", "文字", "SE", "BGM", "効果音", "タイトル", "サブタイトル",
    "提供", "画", "ナレーション", "ナレーター", "実況", "解説", "音声",
    "音楽", "CM", "広告", "注釈", "注記", "キャプション", "字幕", "映像",
    "NA", "ＮＡ", "N.A", "Ｎ．Ａ", "ナレ", "ナレータ",
}

# ===== 行頭スキップパターン =====
SKIP_PREFIX_PATTERNS = [
    r"^【", r"^〔", r"^［", r"^〈", r"^《",
    r"^◆", r"^■", r"^▼", r"^●",
    r"^（", r"^\(",
    r"^[\-―＝]+\s*$",
    r"^\s*#",
    r"^テロップ", r"^解説", r"^実況", r"^映像", r"^ナレ",
]

# ===== Mode-L 検出パターン =====
RE_DELIMITER = re.compile(r"[|‖]{2,}\s*（(\d+)）\s*[|‖]{2,}")
RE_TITLE_CARD = re.compile(r"^[◯○〇]\s*タイトル")
RE_LOCATION_TIME = re.compile(r"^(?P<loc>.+?)[\s\u3000]+(?P<tod>朝|昼|夕|夕方|夕刻|夜|深夜|早朝|明け方|翌朝|日中)$")
RE_CAST_LINE = re.compile(r"^[・◯○〇]?\s*(.+?)[（(][0-9０-９]+[）)]")

# ===== ユーティリティ関数 =====
def normalize_line(line: str) -> str:
    """行の正規化（判定用）"""
    normalized = line.replace("\u3000", " ")
    normalized = re.sub(r"[ \t]+", " ", normalized)
    return normalized.strip()

def clean_spaces(text: str) -> str:
    """全種類の空白を削除"""
    return re.sub(r"[ \t\u3000]+", "", text)

def normalize_character_name(name: str) -> str:
    """人物名の正規化"""
    if not name:
        return ""
    normalized = clean_spaces(name.strip())
    # 括弧内注記を削除
    normalized = re.sub(r"[（(][^）)]+[）)]", "", normalized)
    # 末尾のNA等を削除
    normalized = re.sub(r"(NA|ＮＡ|N\.A|Ｎ．Ａ|ナレ)$", "", normalized, flags=re.IGNORECASE).strip()
    return normalized

def is_valid_character_name(name: str) -> bool:
    """人物名として妥当か（1文字は除外）"""
    return len(normalize_character_name(name)) > 1

def is_blacklisted(name: str) -> bool:
    """ブラックリストチェック（完全一致＋部分一致）"""
    normalized = normalize_character_name(name)
    if normalized in SPEAKER_BLACKLIST:
        return True
    for black in SPEAKER_BLACKLIST:
        if normalized.startswith(black) or black in normalized:
            return True
    return False

def should_skip_line(line: str) -> bool:
    """行頭パターンでスキップ判定"""
    stripped = line.strip()
    for pattern in SKIP_PREFIX_PATTERNS:
        if re.match(pattern, stripped):
            return True
    return False

def extract_given_name(full_name: str) -> str:
    """下の名前を抽出（スペース区切り対応）"""
    if " " in full_name or "\u3000" in full_name:
        parts = re.split(r"[ \t\u3000]+", full_name.strip())
        given = parts[-1]
        given = re.split(r"[（(]", given)[0].strip()
        return clean_spaces(given)
    normalized = normalize_character_name(full_name)
    if len(normalized) >= 2:
        return normalized[-2:] if len(normalized) >= 3 else normalized
    return normalized

def extract_time_of_day(line: str) -> tuple:
    """場所＋時間行から (location, dn) を抽出"""
    m = RE_LOCATION_TIME.match(line.strip())
    if m:
        loc = m.group("loc").strip()
        tod = m.group("tod").strip()
        dn = TIME_MAP.get(tod, "")
        return loc, dn
    return None, None

def extract_delimiter_number(line: str) -> int:
    """区切り行からシーン番号を抽出"""
    m = RE_DELIMITER.search(line)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            pass
    return None

# ===== モード検出 =====
def detect_mode(lines: list) -> str:
    """
    台本フォーマットを判定してモードを返す
    - "L": Mode-L（区切り行モード）
    - "S": Mode-S（標準脚本モード）
    """
    has_delimiter = False
    has_title_card = False
    has_location_time = False
    title_card_line = -1
    
    for i, line in enumerate(lines[:600]):  # 先頭600行で判定
        stripped = line.strip()
        if not stripped:
            continue
        
        # 区切り行検出
        if RE_DELIMITER.search(stripped):
            has_delimiter = True
            break
        
        # タイトルカード検出
        if RE_TITLE_CARD.match(stripped):
            has_title_card = True
            title_card_line = i
        
        # 場所＋時間検出（タイトルカードの近傍）
        if has_title_card and i < title_card_line + 20:
            if RE_LOCATION_TIME.match(stripped):
                has_location_time = True
    
    # Mode-L判定
    if has_delimiter or (has_title_card and has_location_time):
        return "L"
    
    return "S"

# ===== キャスト抽出 =====
def extract_cast(lines: list, mode: str) -> tuple:
    """
    台本冒頭からキャスト一覧を抽出
    Returns: (full_names_set, alias_map)
    """
    full_names = set()
    alias_map = {}
    found_delimiter = False
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        
        # Mode-L: 区切り行が出たら終了
        if mode == "L":
            if RE_DELIMITER.search(stripped):
                found_delimiter = True
                break
        
        # Mode-S: 柱が出たら終了（簡易判定）
        if mode == "S":
            if is_standard_pillar(stripped):
                break
        
        # キャスト行抽出
        m = RE_CAST_LINE.match(stripped)
        if m:
            original = m.group(1).strip()
            normalized = normalize_character_name(original)
            if is_valid_character_name(normalized) and not is_blacklisted(normalized):
                full_names.add(normalized)
                given = extract_given_name(original)
                if given != normalized:
                    alias_map[given] = normalized
        # ◯/○行もチェック
        elif stripped.startswith("◯") or stripped.startswith("○") or stripped.startswith("〇"):
            m = RE_CAST_LINE.match(stripped[1:].strip())
            if m:
                original = m.group(1).strip()
                normalized = normalize_character_name(original)
                if is_valid_character_name(normalized) and not is_blacklisted(normalized):
                    full_names.add(normalized)
                    given = extract_given_name(original)
                    if given != normalized:
                        alias_map[given] = normalized
    
    return full_names, alias_map

def is_standard_pillar(line: str) -> bool:
    """標準形式の柱判定（Mode-S用）"""
    s = line.strip()
    if not s:
        return False
    # S#形式
    if re.match(r"^S#\s*\d+", s, re.IGNORECASE):
        return True
    # 数字開始
    if re.match(r"^[0-9０-９]+", s):
        return True
    # ○/◯/〇開始
    if s.startswith("○") or s.startswith("◯") or s.startswith("〇"):
        return True
    return False

# ===== 話者抽出 =====
def extract_speaker(line: str, valid_chars: set = None) -> str:
    """
    行から話者名を抽出
    - 行頭のみチェック
    - 『…』は絶対に話者とみなさない
    """
    # 行頭スペースチェック
    if re.match(r"^[ \t\u3000]", line):
        return None
    
    # スキップパターン
    if should_skip_line(line):
        return None
    
    stripped = line.strip()
    
    # 「…」のみを話者判定に使用（『…』は地の文の引用なので除外）
    m = re.match(r"^([^「]{1,15})「", stripped)
    if not m:
        return None
    
    speaker_raw = m.group(1).strip()
    
    # 括弧内注記を削除
    speaker_raw = re.sub(r"[（(].*?[）)]", "", speaker_raw).strip()
    if not speaker_raw:
        return None
    
    # 末尾のNA等を削除
    speaker_raw = re.sub(r"(NA|ＮＡ|N\.A|Ｎ．Ａ|ナレ)$", "", speaker_raw, flags=re.IGNORECASE).strip()
    
    normalized = normalize_character_name(speaker_raw)
    if not is_valid_character_name(normalized):
        return None
    
    # ブラックリストチェック
    if is_blacklisted(normalized):
        return None
    
    # valid_charsでフィルタ（指定がある場合）
    if valid_chars is not None:
        if normalized in valid_chars:
            return normalized
        for valid in valid_chars:
            if normalized in valid or valid in normalized:
                return valid
        return None
    
    return normalized

def check_action_line(line: str, char: str) -> bool:
    """地の文行の先頭にキャラ名があるかチェック"""
    stripped = line.strip()
    if not stripped:
        return False
    
    # 「があればセリフ行なので除外
    if re.match(r"^.{1,15}「", stripped):
        return False
    
    # ブラックリスト語で始まる行は除外
    if should_skip_line(line):
        return False
    
    # 行頭がキャラ名 + 助詞/記号
    pattern = rf"^{re.escape(char)}(は|が|も|を|に|へ|で|、|（|)"
    if re.match(pattern, stripped):
        return True
    
    return False

# ===== メインパース関数 =====
def parse_lines(lines: list) -> tuple:
    """
    テキスト行配列をパースして香盤データを返す
    Returns: (rows, characters, debug_info)
    """
    # モード検出
    mode = detect_mode(lines)
    
    # キャスト抽出
    full_names, alias_map = extract_cast(lines, mode)
    valid_chars = full_names.copy()
    
    # デバッグ情報
    debug_info = {
        "mode": mode,
        "cast_count": len(full_names),
        "blacklist_hits": 0,
        "quote_guard_hits": 0,
        "delimiter_numbers": [],
        "scene_numbers": [],
    }
    
    scenes = []
    current_scene = None
    last_location = ""
    last_dn = ""
    in_prologue = True  # Mode-L: 区切り行が出るまでプロローグ
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        
        # ===== Mode-L処理 =====
        if mode == "L":
            # 区切り行検出 → 新シーン開始
            scene_no = extract_delimiter_number(stripped)
            if scene_no:
                in_prologue = False
                debug_info["delimiter_numbers"].append(scene_no)
                
                # 前のシーンを確定
                if current_scene:
                    scenes.append(current_scene)
                
                # シーン作成（last_location/last_dnは後で更新される）
                current_scene = {
                    "scene_no": scene_no,
                    "location": "",  # 後で更新
                    "DN": "",  # 後で更新
                    "lines": [],
                    "speakers": set(),
                }
                # すでにlast_location/last_dnがあれば設定（連続シーンの場合）
                if last_location:
                    current_scene["location"] = last_location
                if last_dn:
                    current_scene["DN"] = last_dn
                continue
            
            # プロローグ中はシーン生成しない
            if in_prologue:
                continue
            
            # タイトルカードはスキップ
            if RE_TITLE_CARD.match(stripped):
                continue
            
            # 場所＋時間行 → メタ情報更新
            loc, dn = extract_time_of_day(stripped)
            if loc:
                last_location = loc
                if dn:
                    last_dn = dn
                if current_scene:
                    current_scene["location"] = loc
                    current_scene["DN"] = dn
                continue
            
            # 「同...」行 → サブロケーションとして記録
            if stripped.startswith("同") or stripped.startswith("同・"):
                if current_scene:
                    current_scene["lines"].append(f"[sub: {stripped}]")
                continue
        
        # ===== Mode-S処理 =====
        else:
            # 標準柱判定
            if is_standard_pillar(stripped):
                if current_scene:
                    scenes.append(current_scene)
                
                # 場所/時間抽出（簡易）
                loc = stripped
                dn = ""
                for key in TIME_KEYS:
                    if key in stripped:
                        parts = stripped.split(key)
                        if len(parts) > 1:
                            loc = parts[0].strip(" 0123456789０-９○◯〇")
                            dn = TIME_MAP.get(key, "")
                            break
                
                current_scene = {
                    "scene_no": len(scenes) + 1,
                    "location": loc,
                    "DN": dn,
                    "lines": [],
                    "speakers": set(),
                }
                continue
        
        # ===== シーン本文処理（共通） =====
        if current_scene is None:
            continue
        
        current_scene["lines"].append(stripped)
        
        # 話者抽出
        speaker = extract_speaker(stripped, valid_chars if mode == "L" else None)
        if speaker:
            current_scene["speakers"].add(speaker)
            valid_chars.add(speaker)
        
        # 地の文行頭チェック（Mode-Lのみ）
        if mode == "L" and not speaker:
            for char in full_names:
                if check_action_line(stripped, char):
                    current_scene["speakers"].add(char)
                    break
            for given, full in alias_map.items():
                if check_action_line(stripped, given):
                    current_scene["speakers"].add(full)
                    break
    
    # 最後のシーンを確定
    if current_scene:
        scenes.append(current_scene)
    
    # デバッグ情報更新
    debug_info["scene_numbers"] = [s["scene_no"] for s in scenes]
    
    # 登場人物リスト
    all_chars = sorted(list(valid_chars))
    
    # CSV行生成
    rows = []
    for scene in scenes:
        summary = summarize_scene(scene["lines"])
        row = {
            "scene_no": scene["scene_no"],
            "location": scene["location"],
            "DN": scene["DN"],
            "summary": summary,
        }
        for char in all_chars:
            row[char] = "◯" if char in scene["speakers"] else ""
        row["props_art"] = ""
        row["notes"] = ""
        rows.append(row)
    
    return rows, all_chars, debug_info

def summarize_scene(lines: list) -> str:
    """シーン要約を生成"""
    # セリフ行以外を要約
    non_dialogue = []
    for line in lines:
        if not re.match(r"^.{1,15}「", line.strip()):
            non_dialogue.append(line.strip())
    
    base = "".join(non_dialogue[:3])
    base = re.sub(r"\[sub: .*?\]", "", base)  # サブロケーション記録を除去
    
    if not base:
        base = lines[0] if lines else ""
    
    if len(base) <= 40:
        return base
    return base[:39] + "…"

def convert_to_csv(rows: list, chars: list) -> str:
    """CSV文字列に変換"""
    output = io.StringIO()
    header = ["scene_no", "location", "D/N", "summary"] + chars + ["props_art", "notes"]
    w = csv.writer(output)
    w.writerow(header)
    
    for row in rows:
        csv_row = [
            row.get("scene_no", ""),
            row.get("location", ""),
            row.get("D/N", ""),
            row.get("summary", ""),
        ]
        for char in chars:
            csv_row.append(row.get(char, ""))
        csv_row.append(row.get("props_art", ""))
        csv_row.append(row.get("notes", ""))
        w.writerow(csv_row)
    
    return output.getvalue()

def parse_docx(docx_path: str) -> tuple:
    """docxをパース（メインエントリポイント）"""
    doc = Document(docx_path)
    lines = [p.text for p in doc.paragraphs if p.text]
    return parse_lines(lines)

def parse_docx_to_csv(docx_path: str) -> str:
    """docxをパースしてCSV文字列を返す（互換性）"""
    rows, chars, _ = parse_docx(docx_path)
    return convert_to_csv(rows, chars)


# ===== テスト =====
if __name__ == "__main__":
    # Mode-L fixture
    test_l = [
        "稲本 澪（２８）",
        "和希（３０）",
        "芳佳（２５）",
        "",
        "◯タイトル「ロス：タイム：ライフ」",
        "",
        "会員制バー 夜",
        "",
        "実況（声）「私たちは...」",
        "テロップ：会員制バー【ウェンブリー】",
        "澪NA「私たちは...」",
        "澪もテキーラを一気に飲み干し。",
        "‖‖（２）‖‖‖‖‖‖",
        "",
        "路地裏 夜",
        "和希「...」",
        "芳佳「...」",
    ]
    
    print("=== Mode-L テスト ===")
    rows, chars, debug = parse_lines(test_l)
    print(f"モード: {debug['mode']}")
    print(f"キャスト: {chars}")
    print(f"シーン数: {len(rows)}")
    print(f"区切り番号: {debug['delimiter_numbers']}")
    print(f"生成番号: {debug['scene_numbers']}")
    for row in rows:
        present = [c for c in chars if row.get(c) == "◯"]
        dn = row.get('DN', 'EMPTY')
        print(f"  シーン{row['scene_no']}: {row['location']} (DN={dn}) - {present}")
    
    # Mode-S fixture（回帰テスト）
    test_s = [
        "【人物表】",
        "・金沢 和希（２７）",
        "・中田 ハル（２５）",
        "",
        "1. カフェ 朝",
        "和希「おはよう」",
        "ハルはコーヒーを飲む。",
        "ハル「おはようございます」",
        "",
        "2. 公園 昼",
        "和希「いい天気だ」",
    ]
    
    print("\n=== Mode-S テスト ===")
    rows, chars, debug = parse_lines(test_s)
    print(f"モード: {debug['mode']}")
    print(f"キャスト: {chars}")
    print(f"シーン数: {len(rows)}")
    for row in rows:
        present = [c for c in chars if row.get(c) == "◯"]
        dn = row.get('DN', '')
        print(f"  シーン{row['scene_no']}: {row['location']} ({dn}) - {present}")
