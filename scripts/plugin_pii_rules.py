"""配布プラグインの PII スキャンの検出ルール（plugin_pii_scan.py から使う）。

名前そのもの（顧客名簿）はリポジトリに入れず、実行時に名簿ファイルから読む。ここに置くのは
よくある姓の辞書・敬称の除外語など、公開されても困らない一般的な材料だけ。
"""
import re

KANJI = r'[一-龥々〆ヶ]'
SURNAMES = (   # よくある姓（区切りは / ）
    '佐藤/鈴木/高橋/田中/伊藤/渡辺/山本/中村/小林/加藤/吉田/山田/佐々木/山口/松本/井上/木村/斎藤/清水/'
    '山崎/池田/橋本/阿部/石川/山下/中島/石井/小川/前田/岡田/長谷川/藤田/後藤/近藤/村上/遠藤/青木/坂本/'
    '斉藤/福田/太田/西村/藤井/金子/岡本/藤原/中野/三浦/原田/中川/松田/竹内/小野/田村/中山/和田/石田/'
    '森田/上田/内田/柴田/酒井/宮崎/横山/高木/安藤/宮本/大野/小島/谷口/今井/工藤/高田/増田/丸山/杉山/'
    '村田/大塚/新井/小山/平野/藤本/河野/上野/野口/武田/松井/千葉/岩崎/菅原/木下/久保/佐野/野村/松尾/'
    '菊地/杉本/市川/古川/大西/島田/水野/桜井/高野/渡部/吉川/山内/西田/飯田/菊池/西川/小松/北村/安田/'
    '五十嵐/川口/平田/中田/久保田/服部/岩田/土屋/川崎/福島/本田/樋口/秋山/田口/永井/山中/中西/吉村/'
    '川上/石原/大橋/松岡/馬場/浅野/荒木/大久保/野田/小池/松下/浅井/岩本/片山'
).split('/')
DUMMY_GIVEN = {'太郎', '花子', '一郎', '次郎', '三郎', '四郎', '五郎', '良子', '和子', '一子'}
HONORIFIC_EXEMPT = {
    'お客', '客', '皆', '顧客', '担当者', '担当', '先方', '事務局', '社長', '経営者', '代表者', '代表', '営業',
    '窓口', '相手', '本人', '各', '同', '貴', '管理者', '取引先', '利用者', '会員', '申請者', '受給者',
    '確認者', '面談者', '関係者', '外部', '社外', '社内', '役員', '従業員', '社員', '上司', '部下', '後任',
    '前任', '新任', '玄人', '初心者', '主催者', '依頼者', '記入者', '作成者', '閲覧者', '編集者', '所有者',
    '回答者', '専門家', '税理士', '社労士', '事業主', '協力者', '発注者', '受注者', '甲野', '乙野', '丙野',
    '神', '仏', '王', '様々', '同様', '模様', '仕様', '多様', '異様', '一様', '有様', '左様', '文様', '皆様',
    'お疲れ', 'ご苦労', 'お世話', 'お陰', 'おかげ', 'ご馳走', 'お互い', '奥', '殿', '坊', '嬢', '両様',
    '職人', '業者', '大工', '店員', '店長', '先生', '医師', '患者', '生徒', '学生', '子供', '子ども', 'お子',
}


KANA_ONLY = re.compile(r'^[ァ-ヶー・]+$')
KANA_BOUNDARY_MAX_LEN = 4   # これ以下の長さのカタカナ名は、長いカタカナ語の一部に含まれるだけなら拾わない


def _roster_alt(name: str) -> str:
    esc = re.escape(name)
    if KANA_ONLY.match(name) and len(name) <= KANA_BOUNDARY_MAX_LEN:
        return rf'(?<![ァ-ヶー]){esc}(?![ァ-ヶー])'
    if name.isascii():   # 英字の名前は単語として現れたときだけ（ハッシュ値や英単語の一部は拾わない）
        return rf'(?<![A-Za-z0-9]){esc}(?![A-Za-z0-9])'
    return esc


def build_rules(roster: list[str], extra: list[tuple[str, re.Pattern]]) -> list[tuple[str, re.Pattern]]:
    surname = '|'.join(sorted(SURNAMES, key=len, reverse=True))
    rules = [
        ('roster', re.compile('|'.join(_roster_alt(n) for n in roster), re.IGNORECASE)),
        ('email-internal', re.compile(r'[A-Za-z0-9._%+-]+@colorfulbox\.(?:co\.)?jp', re.I)),
        ('email', re.compile(r'[A-Za-z0-9._%+-]+@(?!example\.(?:com|org|jp)\b|anthropic\.com\b)'
                             r'[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+')),
        ('google-url', re.compile(r'https?://(?:drive|docs|script|sheets)\.google\.com/[^\s)>\]"\']+')),
        ('long-id', re.compile(r'(?<![A-Za-z0-9_\-])(?=[A-Za-z0-9_\-]*\d)(?=[A-Za-z0-9_\-]*[a-z])'
                               r'(?=[A-Za-z0-9_\-]*[A-Z])[A-Za-z0-9_\-]{25,64}(?![A-Za-z0-9_\-])')),
        ('name-spaced', re.compile(rf'(?:{surname})[ 　]{KANJI}{{1,3}}(?!{KANJI})')),
        ('name-honorific', re.compile(rf'({KANJI}{{1,4}})(?:さん|様(?![式々相子態変])|氏(?!名)|くん|ちゃん)')),
        ('name-gender', re.compile(rf'{KANJI}{{2,5}}[\s ]*[(（][男女][)）]')),
        ('amount', re.compile(r'(?<![A-Za-z0-9_\-,.#{])[1-9]\d{0,2}(?:,\d{3}){2,}(?![A-Za-z0-9_,])'
                              r'|(?<![A-Za-z0-9_\-,.#{])[1-9]\d{6,}(?![A-Za-z0-9_])')),
    ]
    rules.extend(extra)
    return rules


DUMMY_CONTEXT = ('サンプル', 'ダミー', '見本', '架空', '〇〇', '○○', '◯◯', '××', 'XX')
DUMMY_CONTEXT_WIDTH = 6


def _honorific_exempt(m: re.Match, line: str) -> bool:
    word, full = m.group(1), m.group(0)
    if any(word.endswith(e) or full.endswith(e) for e in HONORIFIC_EXEMPT):
        return True
    if any(word.endswith(g) for g in DUMMY_GIVEN):   # 「花子さん」のような見本の名前
        return True
    before = line[max(0, m.start() - DUMMY_CONTEXT_WIDTH):m.start()]
    return any(d in before for d in DUMMY_CONTEXT)


AMOUNT_ROUND_UNIT = 10_000              # 万円単位の丸い数は例示・上限値として扱う
AMOUNT_SKIP_LINE = re.compile(r'content_length|バイト数|\bbytes\b', re.I)   # 「アルバイト」は対象外にしない
SEQUENTIAL_DIGITS = '1234567890' * 2     # 12,345,678 のような見本値
URL_PLACEHOLDER = re.compile(r'<|\.\.\.|…|XXX|\{')
BASE64_VALUE = re.compile(r'^[A-Za-z0-9+/]{40,}={0,2}$')   # シート保護のハッシュ値など（人の情報ではない）
GUID = re.compile(r'[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}')   # 図形・検証ルールの識別子


def exempt(rule: str, m: re.Match, line: str) -> bool:
    text = m.group(0)
    if rule == 'name-honorific':
        return _honorific_exempt(m, line)
    if rule == 'name-spaced':
        return re.split(r'[ 　]', text)[-1] in DUMMY_GIVEN
    if rule == 'amount':
        digits = text.replace(',', '')
        return (int(digits) % AMOUNT_ROUND_UNIT == 0 or len(set(digits)) == 1
                or digits in SEQUENTIAL_DIGITS or bool(AMOUNT_SKIP_LINE.search(line)))
    if rule == 'google-url':
        return bool(URL_PLACEHOLDER.search(text))
    if rule == 'long-id':
        return bool(BASE64_VALUE.match(line.strip()) or GUID.fullmatch(text))
    return False
