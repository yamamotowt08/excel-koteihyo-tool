import sys, os, datetime as dt
import xlsxwriter
sys.path.insert(0, os.path.dirname(__file__))
from vbabin import build_vba_project

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT_DIR = sys.argv[1] if len(sys.argv) > 1 else ROOT
src = open(os.path.join(ROOT, 'src', 'GanttMaker.bas'), encoding='utf-8').read()

# 取り込み用 .bas（Shift-JIS / CRLF）
with open(os.path.join(OUT_DIR, '工程表マクロ.bas'), 'wb') as f:
    f.write(src.replace('\n', '\r\n').encode('cp932'))

vba_bin = build_vba_project([('GanttMaker', src)],
                            ['Sheet1', 'Sheet2', 'Sheet3'])
import tempfile
bin_path = os.path.join(tempfile.gettempdir(), 'vbaProject.bin')
open(bin_path, 'wb').write(vba_bin)

D = dt.date
wb = xlsxwriter.Workbook(os.path.join(OUT_DIR, '工程表自動作成ツール.xlsm'))
wb.set_vba_name('ThisWorkbook')
wb.add_vba_project(bin_path)

FONT = 'Meiryo UI'
def F(**k):
    base = {'font_name': FONT, 'font_size': 10, 'valign': 'vcenter'}
    base.update(k); return wb.add_format(base)
INPUT = '#FFF9DB'
f_title = F(bold=True, font_size=16)
f_note = F(font_color='#666666', font_size=9)
f_sec = F(bold=True, font_size=11, font_color='#FFFFFF', bg_color='#305496')
f_label = F(bg_color='#E7ECF5', border=1, border_color='#A6A6A6')
f_in = F(bg_color=INPUT, border=1, border_color='#A6A6A6')
f_in_date = F(bg_color=INPUT, border=1, border_color='#A6A6A6', num_format='[$-411]yyyy/m/d(aaa)', align='left')
f_in_num = F(bg_color=INPUT, border=1, border_color='#A6A6A6', align='center')
f_head = F(bold=True, bg_color='#305496', font_color='#FFFFFF', border=1, border_color='#A6A6A6', align='center')
f_head_res = F(bold=True, bg_color='#808080', font_color='#FFFFFF', border=1, border_color='#A6A6A6', align='center')
f_row = F(bg_color=INPUT, border=1, border_color='#BFBFBF')
f_row_date = F(bg_color=INPUT, border=1, border_color='#BFBFBF', num_format='[$-411]yyyy/m/d(aaa)', align='center')
f_row_num = F(bg_color=INPUT, border=1, border_color='#BFBFBF', align='center')
f_res_date = F(font_color='#595959', border=1, border_color='#D9D9D9', num_format='[$-411]yyyy/m/d(aaa)', align='center', bg_color='#F2F2F2')
f_res_num = F(font_color='#595959', border=1, border_color='#D9D9D9', align='center', bg_color='#F2F2F2')
f_bar_def = F(bg_color='#000000', border=1, border_color='#A6A6A6')
f_row_color = F(border=1, border_color='#BFBFBF')
f_blue = F(bg_color='#2F5597', border=1, border_color='#BFBFBF')
f_plain = F()

# ================= 入力 =================
ws = wb.add_worksheet('入力'); ws.set_vba_name('Sheet1')
ws.activate()
ws.hide_gridlines(2)
ws.set_column('A:A', 40, f_plain); ws.set_column('B:C', 14, f_plain); ws.set_column('D:D', 7, f_plain)
ws.set_column('E:F', 14, f_plain); ws.set_column('G:G', 7, f_plain); ws.set_column('H:H', 8, f_plain)
ws.set_column('I:I', 24, f_plain); ws.set_column('J:J', 2, f_plain)
ws.set_column('K:L', 13, f_plain); ws.set_column('M:M', 6, f_plain)
ws.set_column('N:O', 13, f_plain); ws.set_column('P:P', 6, f_plain)
ws.set_row(0, 40)
ws.write('A1', '工程表 自動作成ツール', f_title)
ws.insert_button('B1', {'macro': 'MakeSchedule', 'caption': '工程表を作成',
                        'width': 190, 'height': 44, 'x_offset': 4, 'y_offset': 6})
ws.write('A2', '黄色のセルに入力して［工程表を作成］ボタンを押すと、「工程表」シートに工程表ができます（何度でも作り直せます）。日付が変わったときは、オレンジの「変更後」の欄に入れると、変更前のバーの下に変更後のバーが追加されます。', f_note)

CHG = '#FCE4D6'
f_chg_date = F(bg_color=CHG, border=1, border_color='#BFBFBF', num_format='[$-411]yyyy/m/d(aaa)', align='center')
f_chg_num = F(bg_color=CHG, border=1, border_color='#BFBFBF', align='center')
f_head_chg = F(bold=True, bg_color='#C65911', font_color='#FFFFFF', border=1, border_color='#A6A6A6', align='center')
f_bar_chg = F(bg_color='#C00000', border=1, border_color='#A6A6A6')

def section(r, text, last=15):
    ws.merge_range(r, 0, r, last, text, f_sec); ws.set_row(r, 20)

section(3, '■ 工程表の見出し')
header_items = [
    ('cfg_FormNo', '様式番号（左上）', '第２号様式'),
    ('cfg_Title', '表題', '工　　　程　　　表'),
    ('cfg_Date', '作成日', D(2026, 2, 26)),
    ('cfg_To', '宛先', '○○○○事務所長　殿'),
    ('cfg_Preface', '前文', '　　次のとおり業務の工程を設定して表示します。'),
    ('cfg_Project', '業務名', '令和８年度　○○○○整備業務（その１）○○施設検討業務一式'),
    ('cfg_Address', '作成者　所在地', '○○県○○市○○町１丁目２番３号'),
    ('cfg_Company', '作成者　会社名', '株式会社○○○○'),
    ('cfg_Name', '作成者　名前', '代表取締役　○○　○○'),
    ('cfg_Manager', '管理者', '部　長　○○　○○'),
    ('cfg_Staff', '担当者', '係　長　○○　○○'),
    ('cfg_Contact', '確認先（電話）', '０００－０００－００００'),
]
r = 4
for name, label, val in header_items:
    ws.write(r, 0, label, f_label)
    if isinstance(val, D):
        ws.merge_range(r, 1, r, 6, '', f_in_date); ws.write_datetime(r, 1, dt.datetime(val.year, val.month, val.day), f_in_date)
    else:
        ws.merge_range(r, 1, r, 6, val, f_in)
    wb.define_name(name, "='入力'!$B$%d" % (r + 1))
    r += 1

r += 1
section(r, '■ 設定'); r += 1
def setting(r, name, label, note):
    ws.write(r, 0, label, f_label); ws.write(r, 2, note, f_note)
    wb.define_name(name, "='入力'!$B$%d" % (r + 1))
setting(r, 'cfg_DayMode', '日数の数え方', '暦日＝カレンダーどおり／土日祝除く＝土日と「祝日」シートの日を数えない')
ws.write(r, 1, '暦日', f_in_num)
ws.data_validation(r, 1, r, 1, {'validate': 'list', 'source': ['暦日', '土日祝除く']}); r += 1
setting(r, 'cfg_BarColor', 'バーの標準色', 'このセルの塗りつぶし色がバーの色になります（作業ごとの色は下の「バー色」列で指定）')
ws.write_blank(r, 1, None, f_bar_def); r += 1
setting(r, 'cfg_BarHeight', 'バーの太さ（行の高さに対する％）', '10～90')
ws.write(r, 1, 35, f_in_num); r += 1
setting(r, 'cfg_ChangeColor', '変更後バーの色', 'このセルの塗りつぶし色が「変更後」のバーの色になります')
ws.write_blank(r, 1, None, f_bar_chg); r += 1
setting(r, 'cfg_OldStyle', '変更前バーの表示', '変更があった作業の、上段（変更前）のバーの描き方：塗りつぶし／点線／白抜き')
ws.write(r, 1, '塗りつぶし', f_in_num)
ws.data_validation(r, 1, r, 1, {'validate': 'list', 'source': ['塗りつぶし', '点線', '白抜き']}); r += 2

section(r, '■ 作業項目　（開始日・終了日・日数のうち２つ、または日数だけを入力。作業項目を空欄にすると上の作業の追加期間になります）'); r += 1
ws.merge_range(r, 1, r, 3, '変更前（当初の予定）', f_head)
ws.merge_range(r, 4, r, 6, '変更後（日付が変わったときだけ入力）', f_head_chg)
ws.merge_range(r, 10, r, 12, '計算結果：変更前', f_head_res)
ws.merge_range(r, 13, r, 15, '計算結果：変更後', f_head_res)
r += 1
hdr_row = r
for c, t in enumerate(['作業項目', '開始日', '終了日', '日数']):
    ws.write(r, c, t, f_head)
for c, t in zip((4, 5, 6), ['開始日', '終了日', '日数']):
    ws.write(r, c, t, f_head_chg)
ws.write(r, 7, 'バー色※', f_head); ws.write(r, 8, '備考', f_head)
for c, t in zip(range(10, 16), ['開始日', '終了日', '日数'] * 2):
    ws.write(r, c, t, f_head_res)
wb.define_name('tbl_Header', "='入力'!$A$%d" % (r + 1))
ws.set_row(r, 20)
r += 1

blue = 'blue'
# (作業項目, 開始日, 終了日, 日数, 色, 備考, 変更後開始日, 変更後終了日, 変更後日数)
tasks = [
    ('1.詳細検討',),
    ('1.1.計画準備', D(2026, 3, 2), None, 20),
    ('1.2.現地踏査', None, None, 15, None, '日数だけ入力（前の作業の翌日から）'),
    ('1.3.設計条件の検討', D(2026, 4, 1), D(2026, 4, 30)),
    ('1.4.基本検討', D(2026, 5, 1), D(2026, 5, 31), None, None, '終了日を延長', None, D(2026, 6, 10)),
    ('1.5.関連施設の検討', D(2026, 6, 11), D(2026, 6, 30)),
    ('1.6.施工計画', D(2026, 7, 1), D(2026, 7, 20)),
    ('1.7.図面作成', D(2026, 7, 21), D(2026, 8, 10), None, None, '開始日をずらす', D(2026, 7, 27)),
    ('1.8.数量計算', None, D(2026, 8, 31), 21),
    ('1.9.照査', D(2026, 3, 9), None, None, None, '3回'),
    ('', D(2026, 7, 6)),
    ('', D(2026, 9, 24)),
    ('1.10.報告書作成', D(2026, 9, 1), D(2026, 9, 30)),
    ('2.周辺対策検討',),
    ('2.1.周辺対策の検討', D(2026, 5, 1), D(2026, 7, 10), None, blue),
    ('2.2.前後区間への影響検討', D(2026, 7, 11), D(2026, 8, 10), None, blue),
    ('2.3.対策工の検討', D(2026, 6, 1), D(2026, 6, 20), None, blue),
    ('3.打合せ協議（中間５回・計７回）', D(2026, 3, 16), None, 2),
    ('', D(2026, 4, 13), None, 2),
    ('', D(2026, 4, 27), None, 2),
    ('', D(2026, 5, 18), None, 2, None, '第4回を延期', D(2026, 5, 25)),
    ('', D(2026, 7, 13), None, 2),
    ('', D(2026, 8, 17), None, 2),
    ('', D(2026, 9, 28), None, 2),
]

def put_date(r, c, v, fmt):
    if v: ws.write_datetime(r, c, dt.datetime(v.year, v.month, v.day), fmt)
    else: ws.write_blank(r, c, None, fmt)
def put_num(r, c, v, fmt):
    if v: ws.write_number(r, c, v, fmt)
    else: ws.write_blank(r, c, None, fmt)
def write_row(r, t):
    t = list(t) + [None] * (9 - len(t))
    name, s, e, d, color, note, cs, ce, cd = t
    ws.write(r, 0, name or '', f_row)
    put_date(r, 1, s, f_row_date); put_date(r, 2, e, f_row_date); put_num(r, 3, d, f_row_num)
    put_date(r, 4, cs, f_chg_date); put_date(r, 5, ce, f_chg_date); put_num(r, 6, cd, f_chg_num)
    ws.write_blank(r, 7, None, f_blue if color == blue else f_row_color)
    ws.write(r, 8, note or '', f_row)
    for c in (10, 11, 13, 14): ws.write_blank(r, c, None, f_res_date)
    for c in (12, 15): ws.write_blank(r, c, None, f_res_num)

for t in tasks:
    write_row(r, t); r += 1
for _ in range(60):   # 追加入力用の空行
    write_row(r, ()); r += 1

# ================= 使い方 =================
wh = wb.add_worksheet('使い方'); wh.set_vba_name('Sheet2')
wh.hide_gridlines(2); wh.set_column('A:A', 3); wh.set_column('B:B', 110, F(text_wrap=True))
f_h = F(bold=True, font_size=14); f_h2 = F(bold=True, font_size=11, font_color='#305496')
f_t = F(text_wrap=True)
lines = [
    ('h', '工程表 自動作成ツールの使い方'),
    ('', ''),
    ('h2', '１．基本の流れ'),
    ('', '①「入力」シートの黄色いセルに、見出し（業務名・作成者など）と作業項目を入力します。'),
    ('', '②［工程表を作成］ボタンを押します。'),
    ('', '③「工程表」シートに、月を上旬・中旬・下旬に分けた工程表と、日数・塗りつぶしバーが作られます。'),
    ('', '　入力を直したら、もう一度ボタンを押せば作り直されます（工程表シートは毎回まっさらにしてから作ります）。'),
    ('', ''),
    ('h2', '２．作業項目の入れ方'),
    ('', '・開始日＋終了日 …… 日数を自動計算'),
    ('', '・開始日＋日数 ………… 終了日を自動計算'),
    ('', '・終了日＋日数 ………… 開始日を自動計算'),
    ('', '・日数だけ …………… １つ上の期間が終わった翌日から始まる'),
    ('', '・開始日だけ ………… その日１日だけの作業（照査・打合せなど）'),
    ('', '・日付なし …………… 見出し行（「1.詳細検討」など）としてバーなしで表示'),
    ('', '・作業項目を空欄にして日付を入れた行は、上の作業の「追加期間」になり、同じ行にバーが増えます。'),
    ('', '　（打合せを何回も行う場合など。工程表の日数欄には合計日数が出ます）'),
    ('', '・計算した開始日・終了日・日数は、入力シート右側の灰色の欄（計算結果）にも表示されます。'),
    ('', ''),
    ('h2', '３．日付が変わったとき（変更後のバー）'),
    ('', '・その作業の行のオレンジの「変更後」の欄に、新しい日付を入れます。変更前の欄はそのまま残してください。'),
    ('', '・工程表では、その作業の行が上下２段になり、上段に変更前、下段に変更後のバーが描かれます。日数欄も上が変更前、下が変更後です。'),
    ('', '・変更後 開始日だけ …… 日数は変更前と同じで、日程だけずらす'),
    ('', '・変更後 終了日だけ …… 開始日は変更前のままで、延長・短縮'),
    ('', '・変更後 日数だけ ……… 開始日は変更前のままで、日数を変える'),
    ('', '・変更後 開始日＋終了日（または日数）…… そのとおりの期間'),
    ('', '・変更後のバーの色は「変更後バーの色」、変更前のバーを点線や白抜きにしたいときは「変更前バーの表示」で選べます。'),
    ('', '・工程表の「終了予定年月日」は、変更後の予定で表示されます。'),
    ('', ''),
    ('h2', '４．バーの色・太さ'),
    ('', '・「バーの標準色」セルの塗りつぶし色が、全部のバーの色になります。'),
    ('', '・作業ごとに色を変えたいときは、その行の「バー色」セルを塗りつぶしてください。'),
    ('', '・「バーの太さ」は行の高さに対する％です（10～90）。'),
    ('', ''),
    ('h2', '５．日数の数え方'),
    ('', '・暦日 ……………… 土日も含めてカレンダーどおりに数えます。'),
    ('', '・土日祝除く ……… 土日と「祝日」シートに書いた日を数えません（稼働日）。年末年始休暇なども「祝日」シートに追加できます。'),
    ('', '　※「祝日」シートの日付は2026・2027年分を入れてあります。使う前に内容を確認してください。'),
    ('', ''),
    ('h2', '６．注意'),
    ('', '・このファイルはマクロ有効ブック（.xlsm）です。開いたときに「コンテンツの有効化」を押してください。'),
    ('', '　インターネットやメールから取得したファイルで「ブロックされました」と出る場合は、ファイルを右クリック →［プロパティ］→「許可する」にチェックしてから開いてください。'),
    ('', '・行や列を挿入しても、見出しの欄は名前（cfg_～）で読み取るので動きます。作業表の見出し行（作業項目・開始日…）の列の並びは変えないでください。'),
    ('', '・工程表シートは毎回作り直すので、工程表シートで手直しした内容は消えます。手直しした工程表を残したいときは、シートをコピーしてから使ってください。'),
]
for i, (kind, text) in enumerate(lines):
    wh.write(i + 1, 1, text, f_h if kind == 'h' else f_h2 if kind == 'h2' else f_t)

# ================= 祝日 =================
wd = wb.add_worksheet('祝日'); wd.set_vba_name('Sheet3')
wd.set_column('A:A', 16, F(num_format='yyyy/m/d(aaa)', align='center')); wd.set_column('B:B', 22, f_plain)
wd.set_column('D:D', 60, f_note)
wd.write(0, 0, '日付', f_head); wd.write(0, 1, '名称', f_head)
wd.write(0, 3, 'A列に日付を並べてください（空行を入れずに）。「土日祝除く」のときに休みとして扱います。', f_note)
hol = [
    ('2026-01-01', '元日'), ('2026-01-12', '成人の日'), ('2026-02-11', '建国記念の日'), ('2026-02-23', '天皇誕生日'),
    ('2026-03-20', '春分の日'), ('2026-04-29', '昭和の日'), ('2026-05-03', '憲法記念日'), ('2026-05-04', 'みどりの日'),
    ('2026-05-05', 'こどもの日'), ('2026-05-06', '振替休日'), ('2026-07-20', '海の日'), ('2026-08-11', '山の日'),
    ('2026-09-21', '敬老の日'), ('2026-09-22', '国民の休日'), ('2026-09-23', '秋分の日'), ('2026-10-12', 'スポーツの日'),
    ('2026-11-03', '文化の日'), ('2026-11-23', '勤労感謝の日'),
    ('2027-01-01', '元日'), ('2027-01-11', '成人の日'), ('2027-02-11', '建国記念の日'), ('2027-02-23', '天皇誕生日'),
    ('2027-03-21', '春分の日'), ('2027-03-22', '振替休日'), ('2027-04-29', '昭和の日'), ('2027-05-03', '憲法記念日'),
    ('2027-05-04', 'みどりの日'), ('2027-05-05', 'こどもの日'), ('2027-07-19', '海の日'), ('2027-08-11', '山の日'),
    ('2027-09-20', '敬老の日'), ('2027-09-23', '秋分の日'), ('2027-10-11', 'スポーツの日'), ('2027-11-03', '文化の日'),
    ('2027-11-23', '勤労感謝の日'),
]
f_hd = F(num_format='yyyy/m/d(aaa)', align='center')
for i, (d, n) in enumerate(hol):
    wd.write_datetime(i + 1, 0, dt.datetime.fromisoformat(d), f_hd); wd.write(i + 1, 1, n, f_plain)

wb.close()
print('built')
