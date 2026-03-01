"""
香盤ジェネレーター - パーサーモジュール
v0.9.2 - ロスタイム形式対応版
"""
import re
import csv
import io
from docx import Document

# ===== バージョン =====
PARSER_VERSION = "v0.9.2"

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
    "夕刻": "E",
    "黄昏": "E",
    "夜": "N",
    "深夜": "N",
    "翌朝": "M",
    "M": "M",
    "D": "D",
    "E": "E",
    "N": "N",
}
TIME_KEYS = list(TIME_MAP.keys())

# ===== 拡張ブラックリスト（ロスタイム形式対応） =====
SPEAKER_BLACKLIST = {
    "テロップ", "文字", "SE", "BGM", "効果音", "タイトル", "サブタイトル",
    "提供", "画", "ナレーション", "ナレーター", "実況", "解説", "音声",
    "音楽", "CM", "広告", "注釈", "注記", "キャプション", "字幕", "映像",
    "NA", "ＮＡ", "N.A", "Ｎ．Ａ", "ナレ", "ナレータ",
}

# ===== 行頭スキップパターン（人物抽出対象外） =====
SKIP_PREFIX_PATTERNS = [
    r"^【", r"^〔", r"^［", r"^〈", r"^《",
    r"^◆", r"^■", r"^▼", r"^●",
    r"^（", r"^\(",
    r"^[\-―＝]+\s*$",
    r"^\s*#",
    r"^テロップ", r"^解説", r"^実況", r"^映像", r"^ナレ",
]

# ===== ロスタイム形式検出用パターン =====
LOSTIME_SCENE_DELIMITER = re.compile(r"[|‖]{2,}\s*（(\d+)）\s*[|‖]{2,}")
LOSTIME_TITLE_CARD = re.compile(r"^[◯○〇]\s*タイトル")
LOSTIME_LOCATION_TIME = re.compile(r"^(?P<loc>.+?)[\s\u3000]+(?P<tod>朝|昼|夕|夕方|夕刻|夜|深夜|早朝|明け方|翌朝|日中)$")

def normalize_line(line: str) -> str:
    """行の正規化（判定用）"""
    normalized = line.replace("\u3000", " ")
    normalized = re.sub(r"[ \t]+", " ", normalized)
    normalized = normalized.strip()
    normalized = normalized.replace("◯", "○").replace("〇", "○")
    return normalized

def should_skip_speaker_extraction(line: str) -> bool:
    """行頭が特定パターンなら人物抽出をスキップ"""
    stripped = line.strip()
    for pattern in SKIP_PREFIX_PATTERNS:
        if re.match(pattern, stripped):
            return True
    return False

def is_blacklisted_speaker(name: str) -> bool:
    """ブラックリストに含まれる名前かチェック（部分一致も含む）"""
    normalized = normalize_character_name(name)
    # 完全一致
    if normalized in SPEAKER_BLACKLIST:
        return True
    # 部分一致（ブラックリスト語で始まる/含む）
    for black in SPEAKER_BLACKLIST:
        if normalized.startswith(black) or black in normalized:
            return True
    return False

def normalize_numbers(text: str) -> str:
    """全角数字を半角へ"""
    return text.translate(str.maketrans("０１２３４５６７８９", "0123456789"))

def clean_spaces(text: str) -> str:
    """全種類の空白を削除"""
    return re.sub(r"[ \t\u3000]+", "", text)

def normalize_character_name(name: str) -> str:
    """人物名の正規化（括弧注釈削除、NA除去など）"""
    if not name:
        return ""
    
    normalized = clean_spaces(name.strip())
    
    # 括弧内の注記を削除（声）、（NA）、（ＮＡ）、（山本）など
    normalized = re.sub(r"[（(][^）)]+[）)]", "", normalized)
    
    # 末尾の制作付記を削除
    normalized = re.sub(r"(off|OFF|声|OS|ナレーション|ナレ|NA|ＮＡ|M|N)$", "", normalized, flags=re.IGNORECASE)
    
    # 英語名は小文字に統一
    if re.match(r"^[a-zA-Z0-9]+$", normalized):
        normalized = normalized.lower()
    
    return normalized.strip()

def is_valid_character_name(name: str) -> bool:
    """人物名として妥当か（1文字は除外）"""
    normalized = normalize_character_name(name)
    return len(normalized) > 1

def extract_given_name(full_name: str) -> str:
    """フルネームから下の名前（given name）を抽出
    例：金沢和希 → 和希、稲本 澪 → 澪（スペース区切り対応）
    """
    # 正規化前にスペース区切りをチェック
    if " " in full_name or "\u3000" in full_name:
        parts = re.split(r"[ \t\u3000]+", full_name.strip())
        given = parts[-1]
        # 括弧や注記を削除
        given = re.split(r"[（(]", given)[0].strip()
        return clean_spaces(given)
    
    normalized = normalize_character_name(full_name)
    if len(normalized) == 2:
        return normalized
    elif len(normalized) >= 3:
        return normalized[-2:]
    return normalized

def detect_lostime_format(lines: list) -> bool:
    """
    ロスタイム形式かどうかを検出
    - ‖‖（数字）‖‖ パターンの存在
    - または ◯タイトル + 場所＋時間 の組み合わせ
    """
    has_delimiter = False
    has_title_card = False
    has_location_time = False
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        
        # ‖‖（数字）‖‖ パターン検出
        if LOSTIME_SCENE_DELIMITER.search(stripped):
            has_delimiter = True
            break
        
        # ◯タイトル検出
        if LOSTIME_TITLE_CARD.match(stripped):
            has_title_card = True
        
        # 場所＋時間パターン検出
        if LOSTIME_LOCATION_TIME.match(stripped):
            has_location_time = True
        
        # 両方見つかったらロスタイム形式と判定
        if has_title_card and has_location_time:
            return True
    
    return has_delimiter

def extract_scene_number_from_delimiter(line: str) -> int:
    """‖‖（数字）‖‖ からシーン番号を抽出"""
    m = LOSTIME_SCENE_DELIMITER.search(line)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            pass
    return None

def extract_location_time(line: str) -> tuple:
    """
    「会員制バー 夜」から (location, time_of_day) を抽出
    Returns: (location, dn_code) or (None, None)
    """
    m = LOSTIME_LOCATION_TIME.match(line.strip())
    if m:
        loc = m.group("loc").strip()
        tod = m.group("tod").strip()
        dn = TIME_MAP.get(tod, "")
        return loc, dn
    return None, None

def is_lostime_title_card(line: str) -> bool:
    """タイトルカード行かどうか"""
    return LOSTIME_TITLE_CARD.match(line.strip()) is not None

def is_scene_delimiter(line: str) -> bool:
    """シーン区切り行（‖‖（n）‖‖）かどうか"""
    return LOSTIME_SCENE_DELIMITER.search(line.strip()) is not None

def extract_character_from_cast_line(line: str) -> tuple:
    """
    キャスト紹介行から人物名を抽出
    例: "稲本 澪（２８）" → ("稲本澪", "稲本 澪")  # (正規化名, 元の名前)
        "中田ハル（20）" → ("中田ハル", "中田ハル")
    Returns: (normalized_name, original_name) or (None, None)
    """
    stripped = line.strip()
    # 「名前（年齢）」パターン
    m = re.match(r"^(.+?)[（(][0-9０-９]+[）)]", stripped)
    if m:
        original_name = m.group(1).strip()
        normalized = normalize_character_name(original_name)
        if is_valid_character_name(normalized) and not is_blacklisted_speaker(normalized):
            return normalized, original_name
    return None, None

def extract_characters_from_lostime_header(lines: list) -> tuple:
    """
    ロスタイム形式の冒頭からキャラクター一覧を抽出
    Returns: (full_names_set, alias_map)
    - full_names_set: フルネームのセット
    - alias_map: {別名: フルネーム} の辞書
    """
    full_names = set()
    alias_map = {}
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        
        # 「場所 時間」パターンが出たら終了
        if LOSTIME_LOCATION_TIME.match(stripped):
            break
        
        # キャスト紹介行から抽出
        normalized, original = extract_character_from_cast_line(stripped)
        if normalized:
            full_names.add(normalized)
            # given_name（下の名前）を別名として登録（元の名前から抽出）
            given = extract_given_name(original)
            if given != normalized:
                alias_map[given] = normalized
        # ◯名前（年齢）形式もチェック
        elif stripped.startswith("◯") or stripped.startswith("○") or stripped.startswith("〇"):
            normalized, original = extract_character_from_cast_line(stripped[1:].strip())
            if normalized:
                full_names.add(normalized)
                given = extract_given_name(original)
                if given != normalized:
                    alias_map[given] = normalized
    
    return full_names, alias_map

def is_dialogue_line(line: str) -> bool:
    """セリフ行判定（話者名「...」形式）"""
    stripped = line.strip()
    if not stripped:
        return False
    m = re.match(r"^(.{1,15})[「『]", stripped)
    return m is not None

def extract_speaker_from_line(line: str, valid_characters: set = None) -> str:
    """
    セリフ行から話者名を抽出
    - ブラックリスト適用
    - 括弧内注記削除
    - valid_charactersがあれば、その中に存在する場合のみ返す
    """
    # 行頭スペースチェック
    if re.match(r"^[ \t\u3000]", line):
        return None
    
    # スキップパターン
    if should_skip_speaker_extraction(line):
        return None
    
    stripped = line.strip()
    m = re.match(r"^([^「『]{1,15})[「『]", stripped)
    if not m:
        return None
    
    speaker_raw = m.group(1).strip()
    
    # 括弧内注記を削除（実況（声）→ 実況）
    speaker_raw = re.sub(r"[（(].*?[）)]", "", speaker_raw).strip()
    if not speaker_raw:
        return None
    
    # 末尾のNA/ＮＡなどを削除（澪NA → 澪）
    speaker_raw = re.sub(r"(NA|ＮＡ|N\.A|Ｎ．Ａ|ナレ)$", "", speaker_raw, flags=re.IGNORECASE).strip()
    
    # 正規化
    normalized = normalize_character_name(speaker_raw)
    if not is_valid_character_name(normalized):
        return None
    
    # ブラックリストチェック（完全一致＋部分一致）
    if is_blacklisted_speaker(normalized):
        return None
    
    # 有効キャラクターリストでフィルタ（指定がある場合）
    if valid_characters is not None:
        # 完全一致
        if normalized in valid_characters:
            return normalized
        # 部分一致（短い名前が長い名前に含まれるか）
        for valid in valid_characters:
            if normalized in valid or valid in normalized:
                return valid
        return None
    
    return normalized

def check_character_in_action_line(line: str, character: str) -> bool:
    """
    地の文行の先頭にキャラクター名があるかチェック
    例: "澪もテキーラを一気に飲み干し" → True (澪が行頭にある)
    """
    stripped = line.strip()
    if not stripped:
        return False
    
    # セリフ行は除外
    if is_dialogue_line(line):
        return False
    
    # ブラックリスト語で始まる行は除外
    if should_skip_speaker_extraction(line):
        return False
    
    # 行頭がキャラクター名 + 助詞/記号 かどうか
    pattern = rf"^{re.escape(character)}(は|が|も|を|に|へ|で|、|（|)"
    if re.match(pattern, stripped):
        return True
    
    return False

def summarize_scene(scene_lines: list, max_chars: int = 40) -> str:
    """シーンの要約を生成"""
    non_dialogue = [l.strip() for l in scene_lines if l.strip() and not is_dialogue_line(l)]
    base = "".join(non_dialogue[:3]).replace("\n", "").strip()
    if not base:
        base = "".join([l.strip() for l in scene_lines[:1] if l.strip()])
    if len(base) <= max_chars:
        return base
    return base[:max_chars-1] + "…"

def parse_lines_to_kouban(lines: list) -> tuple:
    """
    テキスト行配列をパースして香盤データを返す
    Returns: (rows, character_list)
    - rows: CSV行のリスト（辞書形式）
    - character_list: 登場人物名のリスト
    """
    # ロスタイム形式判定
    is_lostime = detect_lostime_format(lines)
    
    # キャラクター一覧抽出
    if is_lostime:
        full_names, alias_map = extract_characters_from_lostime_header(lines)
        valid_characters = full_names.copy()
    else:
        full_names = set()
        alias_map = {}
        valid_characters = set()
    
    scenes = []
    current_scene = None
    current_scene_no = 1
    pending_scene_no = None  # ‖‖（n）‖‖で検出した番号
    found_first_location = False  # ロスタイム形式：最初の場所＋時間検出フラグ
    
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        
        # --- ロスタイム形式固有の処理 ---
        if is_lostime:
            # シーン区切り行（‖‖（n）‖‖）
            if is_scene_delimiter(stripped):
                new_no = extract_scene_number_from_delimiter(stripped)
                if new_no:
                    pending_scene_no = new_no
                continue
            
            # タイトルカードはスキップ
            if is_lostime_title_card(stripped):
                continue
            
            # 場所＋時間パターンを柱として検出
            loc, dn = extract_location_time(stripped)
            if loc:
                # 前のシーンを確定
                if current_scene:
                    scenes.append(current_scene)
                
                # シーン番号決定
                if pending_scene_no:
                    current_scene_no = pending_scene_no
                    pending_scene_no = None
                else:
                    current_scene_no = len(scenes) + 1
                
                found_first_location = True
                current_scene = {
                    "scene_no": current_scene_no,
                    "location": loc,
                    "DN": dn,
                    "lines": [],
                    "speakers": set(),
                }
                continue
            
            # 最初の場所＋時間が出るまでは人物紹介部分としてスキップ
            if not found_first_location:
                continue
        
        # --- 既存形式（S#/数字/○柱）の処理 ---
        if not is_lostime:
            # 従来の柱判定
            if is_traditional_scene_heading(stripped):
                if current_scene:
                    scenes.append(current_scene)
                
                loc = extract_location_traditional(stripped)
                dn = extract_dn_traditional(stripped)
                current_scene = {
                    "scene_no": len(scenes) + 1,
                    "location": loc,
                    "DN": dn,
                    "lines": [],
                    "speakers": set(),
                }
                continue
        
        # --- シーン本文の処理 ---
        if current_scene is None:
            continue
        
        current_scene["lines"].append(stripped)
        
        # 話者抽出（ロスタイム形式はvalid_charactersでフィルタ、既存形式はフィルタなし）
        speaker = extract_speaker_from_line(stripped, valid_characters if is_lostime else None)
        if speaker:
            current_scene["speakers"].add(speaker)
            valid_characters.add(speaker)
        
        # ロスタイム形式：地の文行の先頭にキャラ名があるかチェック
        if is_lostime and not speaker:
            # フルネームでチェック
            for char in full_names:
                if check_character_in_action_line(stripped, char):
                    current_scene["speakers"].add(char)
                    break
            # given_name（別名）でチェック
            for given, full in alias_map.items():
                if check_character_in_action_line(stripped, given):
                    current_scene["speakers"].add(full)
                    break
    
    # 最後のシーンを確定
    if current_scene:
        scenes.append(current_scene)
    
    # 登場人物リストの作成
    all_characters = sorted(list(valid_characters))
    
    # CSV行の生成
    rows = []
    for scene in scenes:
        summary = summarize_scene(scene["lines"])
        row = {
            "scene_no": scene["scene_no"],
            "location": scene["location"],
            "D/N": scene["DN"],
            "summary": summary,
        }
        # 各キャラクターの登場有無
        for char in all_characters:
            row[char] = "◯" if char in scene["speakers"] else ""
        row["props_art"] = ""
        row["notes"] = ""
        rows.append(row)
    
    return rows, all_characters

def is_traditional_scene_heading(line: str) -> bool:
    """従来形式の柱判定（S#、数字、○）"""
    s = line.strip()
    if not s:
        return False
    
    # S#形式
    if re.match(r"^S#\s*\d+", s, re.IGNORECASE):
        return True
    
    # ○/◯/〇で開始
    s_match = normalize_line(s)
    if s_match.startswith("○"):
        return True
    
    # 数字で開始
    s2 = normalize_numbers(s)
    if re.match(r"^[0-9]+", s2):
        return True
    
    return False

def extract_location_traditional(line: str) -> str:
    """従来形式の柱から場所を抽出"""
    s = line.strip()
    
    # S#形式
    m = re.match(r"^S#\s*\d+\s*(.+)", s, re.IGNORECASE)
    if m:
        rest = m.group(1)
        parts = rest.split("・")
        if parts[-1].strip() in TIME_MAP:
            parts = parts[:-1]
        return "・".join(parts).strip()
    
    # ○/◯/〇で開始
    s_match = normalize_line(s)
    if s_match.startswith("○"):
        return s_match[1:].strip()
    
    # 数字で開始
    s2 = normalize_numbers(s)
    m = re.match(r"^[0-9]+\s*(.+)", s2)
    if m:
        rest = m.group(1).strip()
        for key in TIME_KEYS:
            if key in rest:
                return rest.split(key)[0].strip("（）() 	")
        return rest
    
    return s

def extract_dn_traditional(line: str) -> str:
    """従来形式の柱からD/Nを抽出"""
    s = line.strip()
    
    # S#形式
    if re.match(r"^S#", s, re.IGNORECASE):
        parts = s.split("・")
        if len(parts) > 1:
            last_part = parts[-1].strip()
            if last_part in TIME_MAP:
                return TIME_MAP[last_part]
    
    # 時間語検索
    for key in TIME_KEYS:
        if key in s:
            return TIME_MAP[key]
    
    return ""

def convert_to_csv_string(rows: list, characters: list) -> str:
    """行データとキャラクターリストをCSV文字列に変換"""
    output = io.StringIO()
    
    # ヘッダー
    header = ["scene_no", "location", "D/N", "summary"] + characters + ["props_art", "notes"]
    w = csv.writer(output)
    w.writerow(header)
    
    # データ行
    for row in rows:
        csv_row = [
            row.get("scene_no", ""),
            row.get("location", ""),
            row.get("D/N", ""),
            row.get("summary", ""),
        ]
        for char in characters:
            csv_row.append(row.get(char, ""))
        csv_row.append(row.get("props_art", ""))
        csv_row.append(row.get("notes", ""))
        w.writerow(csv_row)
    
    return output.getvalue()

def parse_docx_to_csv(docx_path: str) -> str:
    """docxをパースしてCSV文字列を返す（メインエントリポイント）"""
    doc = Document(docx_path)
    lines = [p.text for p in doc.paragraphs if p.text]
    
    rows, characters = parse_lines_to_kouban(lines)
    return convert_to_csv_string(rows, characters)


# ===== テスト用ユーティリティ =====
if __name__ == "__main__":
    # テストフィクスチャ
    test_lines = [
        "稲本 澪（２８）",
        "和希（３０）",
        "芳佳（２５）",
        "",
        "◯タイトル「ロス：タイム：ライフ」",
        "",
        "会員制バー 夜",
        "",
        "実況（声）「私たちは、賞味期限のある存在だ」",
        "テロップ：会員制バー【ウェンブリー】VIP個室",
        "澪NA「私たちは、賞味期限のある存在だ」",
        "澪もテキーラを一気に飲み干し。",
        "‖‖（２）‖‖‖‖‖‖",
        "",
        "路地裏 夜",
        "和希「…」",
        "芳佳「…」",
    ]
    
    print("=== テスト実行 ===")
    print(f"ロスタイム形式検出: {detect_lostime_format(test_lines)}")
    
    # デバッグ：キャラクター抽出確認
    fn, am = extract_characters_from_lostime_header(test_lines)
    print(f"フルネーム: {fn}")
    print(f"別名マップ: {am}")
    
    rows, chars = parse_lines_to_kouban(test_lines)
    print(f"\n抽出キャラクター: {chars}")
    print(f"\nシーン数: {len(rows)}")
    
    for row in rows:
        print(f"\nシーン{row['scene_no']}: {row['location']} ({row['D/N']})")
        print(f"  要約: {row['summary'][:40]}...")
        present = [c for c in chars if row.get(c) == "◯"]
        print(f"  登場: {present}")
    
    print("\n=== CSV出力 ===")
    csv_str = convert_to_csv_string(rows, chars)
    print(csv_str[:500])
