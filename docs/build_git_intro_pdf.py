from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, PageBreak, KeepTogether)
from reportlab.graphics.shapes import Drawing, Rect, String, Line, Circle, Polygon
import sys

pdfmetrics.registerFont(TTFont("JP", "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"))
pdfmetrics.registerFont(TTFont("JPP", "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"))

NAVY = colors.HexColor("#1F3A5F")
ACCENT = colors.HexColor("#E8663D")
LIGHT = colors.HexColor("#EEF2F7")
GRID = colors.HexColor("#B8C4D4")
GIT = colors.HexColor("#F05033")
GH = colors.HexColor("#24292F")

base = dict(fontName="JPP", wordWrap="CJK")
sTitle = ParagraphStyle("t", fontSize=24, leading=32, textColor=NAVY, spaceAfter=6, **base)
sSub = ParagraphStyle("s", fontSize=12, leading=18, textColor=colors.HexColor("#555555"), **base)
sH1 = ParagraphStyle("h1", fontSize=16, leading=22, textColor=colors.white, backColor=NAVY,
                     borderPadding=(5, 6, 5, 6), spaceBefore=10, spaceAfter=10, **base)
sH2 = ParagraphStyle("h2", fontSize=12.5, leading=18, textColor=NAVY, spaceBefore=8, spaceAfter=4, **base)
sBody = ParagraphStyle("b", fontSize=10, leading=16, **base)
sBul = ParagraphStyle("bu", parent=sBody, leftIndent=12, bulletIndent=2)
sCell = ParagraphStyle("c", fontSize=9, leading=13, **base)
sCellH = ParagraphStyle("ch", fontSize=9.5, leading=13, textColor=colors.white, **base)
sCode = ParagraphStyle("code", fontName="JP", fontSize=9, leading=13, wordWrap="CJK")
sNote = ParagraphStyle("n", fontSize=9, leading=14, textColor=colors.HexColor("#333333"),
                       backColor=colors.HexColor("#FFF4E5"), borderColor=ACCENT, borderWidth=0.8,
                       borderPadding=6, spaceBefore=6, spaceAfter=8, **base)

P = lambda t, s=sBody: Paragraph(t, s)
def bullets(items):
    return [Paragraph(i, sBul, bulletText="•") for i in items]

def table(rows, widths, header=True, first_col_bold=True):
    data = []
    for r, row in enumerate(rows):
        cells = []
        for c, v in enumerate(row):
            st = sCellH if (header and r == 0) else sCell
            if c == 0 and first_col_bold and not (header and r == 0):
                v = f"<b>{v}</b>"
            cells.append(Paragraph(v, st))
        data.append(cells)
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    style = [
        ("GRID", (0, 0), (-1, -1), 0.5, GRID),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), NAVY))
    for i in range(1 if header else 0, len(rows)):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
    t.setStyle(TableStyle(style))
    return t

def arrow(d, x1, y1, x2, y2, color=NAVY, label=None, above=True):
    d.add(Line(x1, y1, x2, y2, strokeColor=color, strokeWidth=1.4))
    import math
    a = math.atan2(y2 - y1, x2 - x1); L = 6
    d.add(Polygon([x2, y2, x2 - L*math.cos(a-0.4), y2 - L*math.sin(a-0.4),
                   x2 - L*math.cos(a+0.4), y2 - L*math.sin(a+0.4)], fillColor=color, strokeColor=color))
    if label:
        d.add(String((x1+x2)/2, (y1+y2)/2 + (4 if above else -11), label, fontName="JP",
                     fontSize=8.5, fillColor=color, textAnchor="middle"))

def box(d, x, y, w, h, title, sub, fill):
    d.add(Rect(x, y, w, h, rx=6, ry=6, fillColor=fill, strokeColor=NAVY, strokeWidth=1))
    d.add(String(x + w/2, y + h - 17, title, fontName="JP", fontSize=10, fillColor=NAVY, textAnchor="middle"))
    for i, s in enumerate(sub):
        d.add(String(x + w/2, y + h - 32 - i*12, s, fontName="JP", fontSize=8, fillColor=colors.HexColor("#444444"), textAnchor="middle"))

def areas_diagram():
    W = 170*mm; d = Drawing(W, 190)
    bw, bh, y = 100, 62, 72
    xs = [2, 128, 254, W - bw - 2]
    box(d, xs[0], y, bw, bh, "作業ツリー", ["Working Tree", "実際に編集する", "ファイル群"], colors.HexColor("#FFFFFF"))
    box(d, xs[1], y, bw, bh, "ステージング", ["Staging Area", "次のコミットに", "含める変更"], colors.HexColor("#FFF4E5"))
    box(d, xs[2], y, bw, bh, "ローカルリポジトリ", ["Local Repository", "コミット履歴", "（自分のPC）"], colors.HexColor("#E6F0FA"))
    box(d, xs[3], y, bw, bh, "リモートリポジトリ", ["Remote Repository", "GitHub上の共有先", "（チーム共有）"], colors.HexColor("#E9E9EF"))
    cx = [x + bw/2 for x in xs]
    top = y + bh + 14
    for i, lab in enumerate(["git add", "git commit", "git push"]):
        d.add(Line(cx[i], y + bh, cx[i], top, strokeColor=NAVY, strokeWidth=1.4))
        arrow(d, cx[i], top, cx[i+1], top, label=lab)
        d.add(Line(cx[i+1], top, cx[i+1], y + bh + 7, strokeColor=NAVY, strokeWidth=1.4))
        d.add(Polygon([cx[i+1], y+bh, cx[i+1]-3.5, y+bh+7, cx[i+1]+3.5, y+bh+7], fillColor=NAVY, strokeColor=NAVY))
    def back(x_from, x_to, yy, lab):
        d.add(Line(x_from, y, x_from, yy, strokeColor=ACCENT, strokeWidth=1.4))
        arrow(d, x_from, yy, x_to, yy, color=ACCENT, label=lab, above=False)
        d.add(Line(x_to, yy, x_to, y - 7, strokeColor=ACCENT, strokeWidth=1.4))
        d.add(Polygon([x_to, y, x_to-3.5, y-7, x_to+3.5, y-7], fillColor=ACCENT, strokeColor=ACCENT))
    back(cx[3], cx[2] + 20, y - 16, "git fetch / pull")
    back(cx[2] - 20, cx[0], y - 44, "git switch / restore（履歴の内容を作業ツリーへ反映）")
    d.add(Rect(-3, 4, xs[2] + bw + 10, 184, fillColor=None, strokeColor=GIT, strokeDashArray=[3, 3]))
    d.add(String(4, 176, "自分のPC（Git が管理）", fontName="JP", fontSize=9, fillColor=GIT))
    d.add(String(cx[3], 176, "GitHub（クラウド）", fontName="JP", fontSize=9, fillColor=GH, textAnchor="middle"))
    return d

def branch_diagram():
    W = 170*mm; d = Drawing(W, 120)
    ym, yf = 85, 35
    d.add(String(0, ym-3, "main", fontName="JP", fontSize=9.5, fillColor=NAVY))
    d.add(String(0, yf-3, "feature/契約書テンプレ修正", fontName="JP", fontSize=8.5, fillColor=ACCENT))
    mx = [60, 120, 180, 380, 440]
    d.add(Line(mx[0], ym, mx[-1], ym, strokeColor=NAVY, strokeWidth=2))
    fx = [230, 290, 340]
    d.add(Line(mx[2], ym, fx[0], yf, strokeColor=ACCENT, strokeWidth=2))
    d.add(Line(fx[0], yf, fx[-1], yf, strokeColor=ACCENT, strokeWidth=2))
    d.add(Line(fx[-1], yf, mx[3], ym, strokeColor=ACCENT, strokeWidth=2))
    for i, x in enumerate(mx):
        d.add(Circle(x, ym, 7, fillColor=colors.white if i != 3 else NAVY, strokeColor=NAVY, strokeWidth=2))
    for x in fx:
        d.add(Circle(x, yf, 7, fillColor=colors.white, strokeColor=ACCENT, strokeWidth=2))
    d.add(String(mx[2], ym+13, "ブランチ作成", fontName="JP", fontSize=8, textAnchor="middle", fillColor=NAVY))
    d.add(String(mx[3], ym+13, "マージ（統合）", fontName="JP", fontSize=8, textAnchor="middle", fillColor=NAVY))
    d.add(String(sum(fx)/3, yf-20, "main に影響を与えずに作業・レビュー", fontName="JP", fontSize=8, textAnchor="middle", fillColor=ACCENT))
    d.add(String(mx[0], ym-20, "○ = コミット（変更の記録）", fontName="JP", fontSize=8, fillColor=colors.HexColor("#555555")))
    return d

def on_page(c, doc):
    c.saveState()
    c.setFont("JP", 8); c.setFillColor(colors.HexColor("#888888"))
    c.drawString(20*mm, 10*mm, "Git 入門 ― 基本概念イントロダクション（GitHub Copilot 試験導入プロジェクト 立上げ資料）")
    c.drawRightString(190*mm, 10*mm, f"{doc.page}")
    c.setStrokeColor(NAVY); c.setLineWidth(2); c.line(20*mm, 287*mm, 190*mm, 287*mm)
    c.restoreState()

out = sys.argv[1]
doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=20*mm, rightMargin=20*mm, topMargin=16*mm,
                        bottomMargin=18*mm, title="Git 入門 ― 基本概念イントロダクション",
                        author="開発部門 GitHub Copilot 試験導入チーム", subject="Git と GitHub の違い／Git の基本概念")
S = []
W = 170*mm

# ---- 表紙・はじめに ----
S += [Spacer(1, 20), P("Git 入門", sTitle),
      P("基本概念イントロダクション ― Git と GitHub の違いから、日々のバージョン管理まで", sSub),
      Spacer(1, 4), P("対象：GitHub Copilot 試験導入プロジェクト 立上げチーム／開発部門の新任メンバー", sSub),
      Spacer(1, 14)]
S.append(P("はじめに", sH1))
S.append(P("本資料は、社内システム開発部門で「開発ファイルのバージョン管理」を担う方が、最初に押さえておくべき "
           "Git の考え方をまとめたイントロダクションです。GitHub Copilot の試験導入では、Copilot が提案したコードも "
           "最終的には Git で履歴管理され、GitHub 上でレビュー・承認されます。つまり <b>Git／GitHub の理解は、"
           "Copilot を安全に運用するための土台</b>になります。"))
S.append(Spacer(1, 6))
S.append(P("法務・AIリテラシー業務の経験から見ると、Git は次のように捉えると理解しやすくなります。", sBody))
S.append(table([
    ["法務・文書管理の感覚", "Git での対応概念"],
    ["契約書の版管理（第1版→第2版→最終版）", "コミット（いつ・誰が・何を・なぜ変えたかの記録）"],
    ["「変更履歴の記録」機能・赤入れ", "差分（diff）：行単位で変更前後を比較"],
    ["ドラフトを別ファイルで並行検討", "ブランチ：本流に影響を与えない作業用の分岐"],
    ["決裁・稟議（承認後に正本へ反映）", "プルリクエスト＋レビュー → マージ"],
    ["正本の保管庫・キャビネット", "リポジトリ（ローカル／リモート）"],
    ["監査証跡", "コミット履歴・レビュー記録（改ざんが検知可能な構造）"],
], [75*mm, 95*mm]))
S.append(PageBreak())

# ---- Git vs GitHub ----
S.append(P("1. Git と GitHub の違い（比較表）", sH1))
S.append(P("一言でいうと、<b>Git は「バージョン管理の仕組み（ツール）」</b>、<b>GitHub は「Git リポジトリを預けて"
           "チームで共同作業するためのクラウドサービス」</b>です。GitHub がなくても Git は使えますが、"
           "GitHub は Git を前提にしたサービスです。"))
S.append(Spacer(1, 6))
S.append(table([
    ["比較項目", "Git", "GitHub"],
    ["正体", "分散型バージョン管理システム（ソフトウェア）", "Git リポジトリのホスティング＋共同開発プラットフォーム（Web サービス）"],
    ["提供元", "オープンソース（2005年 Linus Torvalds 氏が開発）", "GitHub, Inc.（2018年より Microsoft 傘下）"],
    ["動作場所", "各自の PC（ローカル）", "クラウド（github.com）※GitHub Enterprise Server は自社環境に設置可"],
    ["ネット接続", "不要（オフラインでコミット・履歴参照が可能）", "必要"],
    ["主な役割", "変更履歴の記録、差分比較、ブランチ・マージ、過去版への復元", "リポジトリの共有、アクセス権管理、レビュー、課題管理、自動化"],
    ["操作方法", "コマンドライン（git コマンド）、各種 GUI ツール、IDE", "Web ブラウザ、GitHub Desktop、GitHub CLI（gh）、IDE 連携"],
    ["代表的な機能", "commit / branch / merge / diff / log / clone / push / pull", "Pull Request、Issues、Actions（CI/CD）、Projects、Wiki、セキュリティスキャン"],
    ["AI 機能", "なし", "GitHub Copilot（コード補完・チャット・PR 要約・コードレビュー支援など）"],
    ["費用", "無料（GPLv2 ライセンス）", "無料プランあり。組織利用は Team／Enterprise など有料プラン、Copilot は別途ライセンス"],
    ["権限・統制", "基本的に持たない（ファイルにアクセスできれば操作可能）", "組織・チーム単位の権限、ブランチ保護、必須レビュー、監査ログ、SSO 等"],
    ["代替・類似", "Subversion（SVN）、Mercurial など", "GitLab、Bitbucket、Azure Repos など"],
    ["例え", "文書の「版管理の仕組み」そのもの", "版管理された文書を保管・共有・決裁する「共有キャビネット＋ワークフロー」"],
], [28*mm, 66*mm, 76*mm]))
S.append(P("<b>ポイント：</b>Copilot 試験導入で議論する「誰がどのリポジトリにアクセスできるか」「レビュー必須にするか」"
           "「利用ログをどう残すか」といった統制は、主に <b>GitHub 側の設定</b>で実現します。一方、日々の変更記録そのものは "
           "<b>Git の役割</b>です。", sNote))
S.append(PageBreak())

# ---- 基本概念 ----
S.append(P("2. バージョン管理とは", sH1))
S.append(P("バージョン管理とは、ファイルの変更履歴を記録し、<b>「いつ・誰が・何を・なぜ」</b>変更したかを後から確認したり、"
           "過去の状態に戻したりできるようにする仕組みです。「仕様書_最終版_修正2_本当に最終.docx」のような"
           "ファイル名による管理から卒業し、1つのファイルの中で履歴を一元管理できます。"))
S += bullets([
    "<b>追跡性</b>：変更点と理由（コミットメッセージ）がすべて残る",
    "<b>復元性</b>：誤った変更をしても、任意の時点に戻せる",
    "<b>並行作業</b>：複数人が同じファイル群を同時に安全に編集できる",
    "<b>説明責任</b>：レビューと承認の記録が残り、監査に対応しやすい",
])
S.append(P("集中型と分散型", sH2))
S.append(P("Subversion などの<b>集中型</b>は、履歴を中央サーバーだけが持ちます。Git は<b>分散型</b>で、"
           "各メンバーの PC に履歴を含むリポジトリの完全な複製があります。そのため、オフラインでも作業でき、"
           "サーバー障害時にも履歴が失われにくいという特長があります。"))

S.append(P("3. Git の基本構造：4つの場所", sH1))
S.append(P("Git を理解するうえで最も重要なのが、ファイルの変更が次の4つの場所を順に移動するというイメージです。"))
S.append(Spacer(1, 4))
S.append(areas_diagram())
S.append(table([
    ["場所", "説明"],
    ["作業ツリー", "普段エディタで編集している、目に見えるファイルそのもの。"],
    ["ステージング（インデックス）", "「次のコミットに含める変更」を選んで置いておく準備エリア。変更の一部だけを記録したいときに役立つ。"],
    ["ローカルリポジトリ", "コミットした履歴が保存される場所（.git フォルダ）。自分の PC 内にある。"],
    ["リモートリポジトリ", "GitHub などに置かれた共有用リポジトリ。push で送り、fetch／pull で取り込む。"],
], [45*mm, 125*mm]))
S.append(PageBreak())

S.append(P("4. 主要な概念", sH1))
S.append(P("リポジトリ（Repository）", sH2))
S.append(P("プロジェクトのファイルと、その全変更履歴をまとめて保管する入れ物です。<font name='JP'>git init</font> で新規作成するか、"
           "<font name='JP'>git clone</font> で GitHub から複製して始めます。"))
S.append(P("コミット（Commit）", sH2))
S.append(P("ある時点のファイルの状態を記録した「スナップショット」です。各コミットには固有の ID（ハッシュ値。例：<font name='JP'>58ad506</font>）、"
           "作成者、日時、メッセージが付きます。ハッシュ値は内容から計算されるため、過去の履歴を書き換えると ID が変わり、"
           "改ざんが検知可能です。<b>コミットメッセージは「何を・なぜ」を簡潔に書く</b>のが基本です。"))
S.append(P("ブランチ（Branch）とマージ（Merge）", sH2))
S.append(P("ブランチは、本流（通常 <font name='JP'>main</font>）から分岐して作業するための仕組みです。新機能や修正は"
           "ブランチで行い、確認が済んだら本流へマージ（統合）します。"))
S.append(branch_diagram())
S.append(P("同じ箇所を複数人が別々に変更すると<b>コンフリクト（競合）</b>が発生します。Git は自動で決められないため、"
           "人が内容を確認してどちらを採用するか決めます。"))
S.append(P("リモート（Remote）と同期", sH2))
S += bullets([
    "<b>clone</b>：リモートリポジトリを丸ごと手元に複製する（最初の1回）",
    "<b>push</b>：自分のコミットをリモートへ送る",
    "<b>fetch</b>：リモートの最新履歴を取得する（作業ツリーは変えない）",
    "<b>pull</b>：fetch ＋ 取り込み（マージ）をまとめて行う",
])
S.append(P("プルリクエスト（Pull Request, PR）※GitHub の機能", sH2))
S.append(P("「このブランチの変更を main に取り込んでください」という依頼です。差分の確認、コメント、承認、自動テストの結果が"
           "1か所に集約され、<b>社内の決裁フローに相当する記録</b>になります。Copilot が生成したコードも、PR でのレビューを"
           "経てから本流に入れる運用が基本です。"))
S.append(PageBreak())

S.append(P("5. よく使う Git コマンド", sH1))
S.append(table([
    ["目的", "コマンド", "説明"],
    ["初期設定", "git config --global user.name \"氏名\"<br/>git config --global user.email \"メール\"", "コミットに記録される作成者情報を設定（最初に1回）"],
    ["開始", "git init／git clone &lt;URL&gt;", "新規リポジトリ作成／GitHub から複製"],
    ["状態確認", "git status", "変更・ステージング状況を表示（迷ったらまずこれ）"],
    ["差分確認", "git diff", "変更内容を行単位で表示"],
    ["ステージ", "git add &lt;ファイル&gt;", "次のコミットに含める変更を選ぶ"],
    ["記録", "git commit -m \"メッセージ\"", "ステージした変更を履歴に記録"],
    ["履歴", "git log --oneline", "コミット履歴を一覧表示"],
    ["ブランチ", "git branch／git switch -c &lt;名前&gt;", "ブランチ一覧／新規作成して切替"],
    ["統合", "git merge &lt;ブランチ&gt;", "指定ブランチの変更を現在のブランチへ統合"],
    ["同期", "git push／git pull／git fetch", "リモートへ送信／取得して統合／取得のみ"],
    ["取り消し", "git restore &lt;ファイル&gt;<br/>git revert &lt;コミットID&gt;", "作業中の変更を破棄／過去のコミットを打ち消す新コミットを作成"],
], [22*mm, 70*mm, 78*mm]))
S.append(P("<b>注意：</b><font name='JP'>git push --force</font> や <font name='JP'>git reset --hard</font> は履歴や作業内容を消す可能性がある操作です。"
           "共有ブランチでは使用しない、または GitHub の「ブランチ保護ルール」で禁止するのが一般的です。", sNote))

S.append(P("6. 基本的な作業の流れ（GitHub Flow）", sH1))
S.append(table([
    ["手順", "操作", "ポイント"],
    ["1", "main を最新化（git switch main → git pull）", "古い状態から作業を始めない"],
    ["2", "作業ブランチを作成（git switch -c feature/xxx）", "1つの目的に1ブランチ"],
    ["3", "編集 → git add → git commit を繰り返す", "小さく、意味のある単位でコミット"],
    ["4", "git push でブランチを GitHub へ送る", "バックアップ兼共有"],
    ["5", "GitHub でプルリクエストを作成", "目的・変更点・確認方法を記載"],
    ["6", "レビュー・自動テスト → 修正", "Copilot 生成コードも人が必ず確認"],
    ["7", "承認後に main へマージ、ブランチを削除", "記録は PR に残る"],
], [14*mm, 80*mm, 76*mm]))
S.append(PageBreak())

S.append(P("7. GitHub Copilot 試験導入に向けた留意点", sH1))
S.append(P("これまでの法務・AI リテラシーの知見が特に活きる領域です。導入前に、以下を Git／GitHub の運用ルールとして整理しておくことを推奨します。"))
S.append(table([
    ["観点", "確認・検討事項"],
    ["機密情報の管理", "パスワード・API キー・個人情報をリポジトリにコミットしない。<font name='JP'>.gitignore</font> で除外し、GitHub のシークレットスキャンを有効化する。Git の履歴は残り続けるため、一度コミットした秘密情報は「削除」ではなく無効化（キーの再発行）が必要。"],
    ["Copilot の参照範囲", "Copilot Business／Enterprise の「コンテンツ除外」設定で、機密性の高いファイルやリポジトリを Copilot の参照対象から外せるか検討する。"],
    ["ライセンス・著作権", "公開コードと一致する提案をブロックする設定（パブリックコードとの一致の抑制）の有効化を検討。利用する OSS のライセンス表記の扱いをルール化する。"],
    ["データの取扱い", "プロンプトや提案データの保存・学習利用の有無は、契約プランと GitHub の最新の利用規約・プライバシー文書で確認する。"],
    ["レビュー体制", "ブランチ保護で main への直接 push を禁止し、PR レビュー承認を必須化。「AI が書いたコードも人が責任を持って承認する」原則を明文化する。"],
    ["アクセス権", "組織・チーム単位で最小権限を付与。退職・異動時の権限削除手順を定める。"],
    ["記録・監査", "コミット履歴・PR・監査ログにより「誰が何を承認したか」を追跡可能にする。Copilot の利用状況は組織の管理画面・監査ログで確認できる範囲を把握する。"],
], [35*mm, 135*mm]))
S.append(P("※ Copilot の機能名・設定項目・プランの内容は更新が頻繁なため、導入判断の際は GitHub 公式ドキュメントと契約条件の最新版を必ず確認してください。", sNote))

S.append(PageBreak())
S.append(P("8. 用語集", sH1))
S.append(table([
    ["用語", "意味"],
    ["リポジトリ", "ファイルと変更履歴の保管場所"],
    ["コミット", "変更の記録（スナップショット）"],
    ["ブランチ", "本流から分岐した作業ライン"],
    ["マージ", "ブランチの変更を統合すること"],
    ["コンフリクト", "同じ箇所への異なる変更が衝突した状態"],
    ["クローン", "リモートリポジトリの複製を手元に作ること"],
    ["プッシュ／プル", "リモートへの送信／リモートからの取得と統合"],
    ["プルリクエスト", "変更の取り込みを依頼し、レビューを受ける GitHub の機能"],
    ["HEAD", "現在作業している位置（コミット）を指す目印"],
    [".gitignore", "Git の管理対象から除外するファイルを指定する設定ファイル"],
], [35*mm, 135*mm]))

S.append(P("9. 次のステップ", sH2))
S += bullets([
    "テスト用リポジトリを作成し、clone → 編集 → commit → push → PR → マージを一通り体験する",
    "ブランチ保護・必須レビューなどの GitHub 組織設定を、試験導入チームで確認する",
    "Copilot 利用ガイドライン（機密情報・ライセンス・レビュー責任）の草案を作成する",
    "参考：Pro Git（日本語版・無料公開）https://git-scm.com/book/ja/v2 ／ GitHub Docs https://docs.github.com/ja",
])

doc.build(S, onFirstPage=on_page, onLaterPages=on_page)
print("ok", out)
