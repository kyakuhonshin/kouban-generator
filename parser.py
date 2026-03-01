import re
import csv
import io
from docx import Document

# ===== バージョン =====
PARSER_VERSION = "v0.9.1"

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

# ===== 人物抽出ブラックリスト =====
SPEAKER_BLACKLIST = {
    "テロップ", "文字", "SE", "BGM", "効果音", "タイトル", "サブタイトル",
    "提供", "画", "ナレーション", "ナレーター", "実況", "解説", "音声",
    "音楽", "CM", "広告", "注釈", "注記", "キャプション", "字幕"
}

# ===== 行頭スキップパターン（人物抽出対象外） =====
SKIP_PREFIX_PATTERNS = [
    r"^【", r"^〔", r"^［", r"^〈", r"^《",  # 括弧類
    r"^◆", r"^■", r"^▼", r"^●",           # 記号類
    r"^（", r"^\(",                         # 始め括弧
    r"^[\-―＝]+\s*$",                       # 装飾線
    r"^\s*#",                               # コメント風
]

def normalize_line(line: str) -> str:
    """
    行の正規化（判定用）
    - 全角スペース→半角スペース
    - 連続スペースを1つに
    - 行頭/行末の空白をstrip
    - ◯〇→○に統一（判定用）
    """
    # 全角スペース→半角
    normalized = line.replace("\u3000", " ")
    # 連続スペースを1つに
    normalized = re.sub(r"[ \t]+", " ", normalized)
    # strip
    normalized = normalized.strip()
    # ◯〇を○に統一（判定用）
    normalized = normalized.replace("◯", "○").replace("〇", "○")
    return normalized

def should_skip_speaker_extraction(line: str) -> bool:
    """
    行頭が特定パターンなら人物抽出をスキップ
    """
    stripped = line.strip()
    for pattern in SKIP_PREFIX_PATTERNS:
        if re.match(pattern, stripped):
            return True
    return False

def is_blacklisted_speaker(name: str) -> bool:
    """
    ブラックリストに含まれる名前かチェック
    """
    normalized = normalize_character_name(name)
    return normalized in SPEAKER_BLACKLIST

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

def extract_given_name(full_name: str) -> str:
    """
    フルネームから下の名前（given name）を抽出
    例：金沢和希 → 和希
    """
    normalized = normalize_character_name(full_name)
    # 日本語名の場合、姓と名の区切りを推測
    # 2文字の場合はそのまま、3文字以上の場合は後半2文字をgiven nameとする
    if len(normalized) == 2:
        return normalized
    elif len(normalized) >= 3:
        # 後半2文字をgiven nameと仮定
        return normalized[-2:]
    return normalized

def is_snumber_heading(line: str) -> bool:
    """S#形式の柱かどうか"""
    s = line.strip()
    if not s:
        return False
    return re.match(r"^S#\s*\d+", s, re.IGNORECASE) is not None

def is_scene_heading(line: str, snumber_mode: bool = False) -> bool:
    """
    柱判定
    - snumber_mode=True: S#のみを柱とみなす
    - snumber_mode=False: S#, 数字, ○/◯/〇対応
    """
    s = line.strip()
    if not s:
        return False
    
    # S#形式（共通）
    if re.match(r"^S#\s*\d+", s, re.IGNORECASE):
        return True
    
    # S#モード時はS#のみを柱とみなす（回帰防止）
    if snumber_mode:
        return False
    
    # 判定用に正規化（◯〇→○）
    s_match = normalize_line(s)
    
    # 従来形式 + ◯/〇対応（S#モードでない時のみ）
    # ○/◯/〇で開始
    if s_match.startswith("○"):
        return True
    
    # 数字で開始（従来の判定を維持）
    s2 = normalize_numbers(s)
    if re.match(r"^[0-9]+", s2):
        return True
    
    return False

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

def extract_character_from_intro_line(line: str) -> str:
    """
    キャラ紹介行から人物名を抽出
    例: "◯中田ハル（２０）" → "中田ハル"
         "○山田太郎（30）" → "山田太郎"
    """
    # 正規化（◯〇→○）
    normalized = normalize_line(line)
    
    # ○で始まり、直後に名前＋（年齢）のパターン
    m = re.match(r'^○\s*([^（\s]{2,20})[（(]', normalized)
    if m:
        name = m.group(1).strip()
        normalized_name = normalize_character_name(name)
        
        # ブラックリストチェック
        if is_blacklisted_speaker(normalized_name):
            return None
        
        if is_valid_character_name(normalized_name):
            return normalized_name
    
    return None

def extract_character_table(lines: list) -> tuple:
    """
    【人物表】セクションから人物リストと別名辞書を抽出
    ・金沢 和希（27）・・・ のような形式から「金沢和希」を抽出
    ◯/○から始まるキャラ紹介行も抽出対象に追加
    
    Returns:
        (full_names, alias_map)
        - full_names: フルネームのリスト
        - alias_map: {given_name: full_name} の辞書
    """
    full_names = []
    alias_map = {}
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
            
            full_normalized = normalize_character_name(name)
            if is_valid_character_name(full_normalized):
                full_names.append(full_normalized)
                
                # 別名（given name）を登録
                given = extract_given_name(full_normalized)
                if given != full_normalized and given not in alias_map:
                    alias_map[given] = full_normalized
        
        # キャラ紹介行（◯/○ + 名前 + （年齢））からも抽出
        intro_name = extract_character_from_intro_line(stripped)
        if intro_name and intro_name not in full_names:
            full_names.append(intro_name)
            
            # 別名（given name）を登録
            given = extract_given_name(intro_name)
            if given != intro_name and given not in alias_map:
                alias_map[given] = intro_name
    
    return full_names, alias_map

def is_dialogue_line(line: str) -> bool:
    """
    セリフ行判定
    - 行に「」または『』があり
    - quote前が短い（<=12文字）
    """
    stripped = line.strip()
    if not stripped:
        return False
    
    # 「または『を探す
    m = re.match(r"^(.{1,12})[「『]", stripped)
    if m:
        return True
    return False

def is_valid_speaker(speaker_raw: str) -> bool:
    """
    話者として妥当か厳格に判定
    - 長さ 1〜12文字程度
    - 「、」「。」「…」を含まない
    - INT / EXT 等の柱語ではない
    """
    if not speaker_raw:
        return False
    
    # 長さチェック
    if len(speaker_raw) < 1 or len(speaker_raw) > 12:
        return False
    
    # 句読点・記号を含む場合はト書きの可能性
    if re.search(r"[、。…，．！？]", speaker_raw):
        return False
    
    # 柱語を含む場合は除外
    if re.search(r"INT|EXT|S#|○|場面|シーン", speaker_raw, re.IGNORECASE):
        return False
    
    # 助詞っぽい文字列を含む場合は文章の可能性
    if re.search(r"[はがをにてでと]$", speaker_raw):
        return False
    
    return True

def extract_speaker_from_line(line: str, full_names: list = None, alias_map: dict = None) -> str:
    """
    行から話者名を抽出（超厳格化版）
    mito(off)「〜」や 和希（声）「〜」に対応
    ト書き内の『』を誤抽出しない
    
    追加条件：
    1. 必ず行頭からquoteまでがspeaker_raw（インデントなし）
    2. 行頭に全角/半角スペースがない
    3. 特定の行頭パターンはスキップ（【〔など）
    4. ブラックリストに含まれる名前は無視
    5. 括弧付き話者は括弧前の名前を使用
    6. 人物マスターに存在しない場合は無視（新規人物追加禁止）
    """
    # 条件2: 行頭にスペースがある場合は話者抽出しない（インデント行はト書き）
    if re.match(r"^[ \t\u3000]", line):
        return None
    
    # 条件3: 行頭スキップパターンをチェック
    if should_skip_speaker_extraction(line):
        return None
    
    # 条件1: 行頭からquoteまでがspeaker_raw（strip前の行頭チェック）
    m = re.match(r"^([^「『]+)[「『]", line)
    if not m:
        return None
    
    speaker_raw = m.group(1).strip()
    
    # 条件5: 括弧付き話者の正規化（「実況（山本）」→「実況」）
    speaker_raw = re.sub(r'[（(].*?[）)]', '', speaker_raw).strip()
    if not speaker_raw:
        return None
    
    # 厳格な話者判定
    if not is_valid_speaker(speaker_raw):
        return None
    
    # 正規化
    normalized = normalize_character_name(speaker_raw)
    if not is_valid_character_name(normalized):
        return None
    
    # 条件4: ブラックリストチェック
    if is_blacklisted_speaker(normalized):
        return None
    
    # 条件6: 人物マスターに存在しない場合は無視（新規人物追加禁止）
    if full_names and alias_map:
        # 完全一致チェック
        if normalized in full_names:
            return normalized
        # 別名辞書で変換
        if normalized in alias_map:
            return alias_map[normalized]
        # 部分一致チェック
        for full_name in full_names:
            if normalized in full_name or full_name in normalized:
                return full_name
        # マスターに存在しない場合は無視
        return None
    
    return normalized

def is_character_in_text(name: str, text: str) -> bool:
    """
    テキスト中に人物名が含まれるか判定
    - 1文字の名前は前後が漢字でない条件を付ける（誤爆防止）
    """
    if not name or not text:
        return False
    
    normalized_name = normalize_character_name(name)
    normalized_text = clean_spaces(text)
    
    if len(normalized_name) == 1:
        # 1文字の場合：前後が漢字でない単独出現に限定
        # (?<![一-龯]) name (?![一-龯])
        pattern = rf"(?<![一-龯ぁ-んァ-ンa-zA-Z0-9]){re.escape(normalized_name)}(?![一-龯ぁ-んァ-ンa-zA-Z0-9])"
        return re.search(pattern, normalized_text) is not None
    else:
        # 2文字以上：通常の部分一致
        return normalized_name in normalized_text

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
    """docxをパースしてCSV文字列を返す（S#モード対応・本文検索版）"""
    doc = Document(docx_path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    
    # 【人物表】から人物マスターと別名辞書を抽出
    full_names, alias_map = extract_character_table(lines)
    
    # S#モード判定（S#柱が1つでも存在するか）
    snumber_mode = any(is_snumber_heading(line) for line in lines)
    
    scenes = []
    current = None
    auto_no = 1
    found_first_snumber = False  # 最初のS#柱検出フラグ（S#モード用）
    
    for line in lines:
        if is_scene_heading(line, snumber_mode):
            # S#モード時：最初のS#柱検出チェック
            if snumber_mode and is_snumber_heading(line):
                if not found_first_snumber:
                    found_first_snumber = True
                    # 最初のS#柱検出前のシーンは破棄
                    if current:
                        current = None
            
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
                # 超厳格化された話者抽出（人物マスター必須）
                speaker = extract_speaker_from_line(line, full_names, alias_map)
                if speaker:
                    current["dialogue_characters"].add(speaker)
    
    if current:
        scenes.append(current)
    
    # 全人物一覧（人物表ベース > セリフベース）
    all_chars = full_names.copy() if full_names else []
    
    # セリフから抽出した人物を追加（マッピング済み）
    for s in scenes:
        for c in s["dialogue_characters"]:
            if c not in all_chars and is_valid_character_name(c):
                all_chars.append(c)
    
    # 登場◯判定（話者のみ - 本文検索は誤爆防止のため廃止）
    # B-4: 「セリフ本文に名前が出ただけ」で◯が付かない保証
    for s in scenes:
        # speakerとして抽出された人物のみを登場とみなす
        s["all_characters"] = set(s["dialogue_characters"])
    
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
