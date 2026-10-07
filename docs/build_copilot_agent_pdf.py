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

def loop_diagram():
    W = 170*mm; d = Drawing(W, 200)
    bw, bh = 108, 44
    # 中央のループ（計画 → 実行 → 確認 → 修正）
    pos = {"plan": (187, 140), "act": (345, 82), "check": (187, 24), "fix": (29, 82)}
    labels = {"plan": ("① 計画", "タスクを分解"), "act": ("② 実行", "編集・コマンド・テスト"),
              "check": ("③ 確認", "結果・エラーを観察"), "fix": ("④ 修正", "必要なら計画を直す")}
    fills = {"plan": "#E6F0FA", "act": "#FFF4E5", "check": "#E6F0FA", "fix": "#FFF4E5"}
    for k, (x, y) in pos.items():
        box(d, x, y, bw, bh, labels[k][0], [labels[k][1]], colors.HexColor(fills[k]))
    c = {k: (x + bw/2, y + bh/2) for k, (x, y) in pos.items()}
    arrow(d, pos["plan"][0] + bw, c["plan"][1], c["act"][0], pos["act"][1] + bh)
    arrow(d, c["act"][0], pos["act"][1], pos["check"][0] + bw, c["check"][1])
    arrow(d, pos["check"][0], c["check"][1], c["fix"][0], pos["fix"][1])
    arrow(d, c["fix"][0], pos["fix"][1] + bh, pos["plan"][0], c["plan"][1])
    d.add(String(W/2, 100, "完了まで自律的に繰り返す", fontName="JP", fontSize=9,
                 fillColor=ACCENT, textAnchor="middle"))
    d.add(String(0, 186, "入力：Issue／チャットの指示", fontName="JP", fontSize=9, fillColor=NAVY))
    d.add(String(W, 186, "出力：変更差分・プルリクエスト → 人がレビュー", fontName="JP", fontSize=9,
                 fillColor=NAVY, textAnchor="end"))
    return d

def flow_diagram():
    W = 170*mm; d = Drawing(W, 110)
    steps = [("Issue を", "Copilot に割当"), ("copilot/ ブランチ", "を作成"), ("クラウドで作業", "（コミット）"),
             ("ドラフト PR", "を作成"), ("人がレビュー", "修正をコメント"), ("人が承認", "→ マージ")]
    n = len(steps); bw = 66; gap = (W - n*bw) / (n - 1); y = 40
    for i, (a, b) in enumerate(steps):
        x = i * (bw + gap)
        human = i >= 4 or i == 0
        d.add(Rect(x, y, bw, 46, rx=6, ry=6, strokeColor=NAVY, strokeWidth=1,
                   fillColor=colors.HexColor("#FFF4E5" if human else "#E6F0FA")))
        d.add(String(x + bw/2, y + 28, a, fontName="JP", fontSize=7.8, fillColor=NAVY, textAnchor="middle"))
        d.add(String(x + bw/2, y + 14, b, fontName="JP", fontSize=7.8, fillColor=NAVY, textAnchor="middle"))
        if i < n - 1:
            arrow(d, x + bw + 1, y + 23, x + bw + gap - 1, y + 23)
    d.add(Rect(0, 8, 9, 9, fillColor=colors.HexColor("#FFF4E5"), strokeColor=NAVY, strokeWidth=0.6))
    d.add(String(13, 9, "人の作業", fontName="JP", fontSize=8, fillColor=NAVY))
    d.add(Rect(70, 8, 9, 9, fillColor=colors.HexColor("#E6F0FA"), strokeColor=NAVY, strokeWidth=0.6))
    d.add(String(83, 9, "Copilot の作業（GitHub Actions 上の隔離環境）", fontName="JP", fontSize=8, fillColor=NAVY))
    # レビューコメント → 再作業 の戻り矢印
    x_rev = 4*(bw + gap) + bw/2; x_work = 2*(bw + gap) + bw/2
    d.add(Line(x_rev, y + 46, x_rev, y + 62, strokeColor=ACCENT, strokeWidth=1.3))
    arrow(d, x_rev, y + 62, x_work, y + 62, color=ACCENT, label="@copilot で修正依頼 → 追加コミット")
    d.add(Line(x_work, y + 62, x_work, y + 53, strokeColor=ACCENT, strokeWidth=1.3))
    d.add(Polygon([x_work, y + 46, x_work - 3.5, y + 53, x_work + 3.5, y + 53], fillColor=ACCENT, strokeColor=ACCENT))
    return d

def on_page(c, doc):
    c.saveState()
    c.setFont("JP", 8); c.setFillColor(colors.HexColor("#888888"))
    c.drawString(20*mm, 10*mm, "GitHub Copilot 入門 ― AI エージェント編（GitHub Copilot 試験導入プロジェクト 立上げ資料）")
    c.drawRightString(190*mm, 10*mm, f"{doc.page}")
    c.setStrokeColor(NAVY); c.setLineWidth(2); c.line(20*mm, 287*mm, 190*mm, 287*mm)
    c.restoreState()

out = sys.argv[1]
doc = SimpleDocTemplate(out, pagesize=A4, leftMargin=20*mm, rightMargin=20*mm, topMargin=16*mm,
                        bottomMargin=18*mm, title="GitHub Copilot 入門 ― AI エージェント編",
                        author="開発部門 GitHub Copilot 試験導入チーム",
                        subject="GitHub Copilot のエージェント機能・仕組み・ガバナンス")
S = []
C = lambda t: f"<font name='JP'>{t}</font>"

# ---- 表紙・はじめに ----
S += [Spacer(1, 20), P("GitHub Copilot 入門", sTitle),
      P("AI エージェント編 ― 「コード補完」から「仕事を任せる AI」へ", sSub),
      Spacer(1, 4), P("対象：GitHub Copilot 試験導入プロジェクト 立上げチーム／開発部門の新任メンバー", sSub),
      P("前提資料：「Git 入門 ― 基本概念イントロダクション」", sSub), Spacer(1, 14)]
S.append(P("はじめに", sH1))
S.append(P("GitHub Copilot は、GitHub が提供する AI によるコーディング支援サービスです。当初は「入力中のコードの続きを"
           "提案する」補完ツールでしたが、現在は<b>指示を受けて自ら計画し、ファイルを編集し、テストを実行し、"
           "プルリクエストまで作成する「AI エージェント」</b>としての機能を備えています。"))
S.append(Spacer(1, 4))
S.append(P("本資料では、エージェントとしての Copilot の仕組みと、試験導入にあたって組織として決めるべきルールを整理します。"
           "AI エージェントは「作業を任せられる」分、<b>任せる範囲・確認の方法・責任の所在</b>をあらかじめ設計しておくことが"
           "重要です。これは AI リテラシーや法務の観点が最も活きる領域でもあります。"))
S.append(Spacer(1, 6))
S.append(table([
    ["社内業務での例え", "Copilot の機能"],
    ["文章作成ソフトの予測変換", "コード補完（入力中に続きを提案）"],
    ["詳しい同僚への質問", "Copilot Chat（質問・説明・修正案の提示）"],
    ["隣の席で一緒に作業するアシスタント", "エージェントモード（IDE 内で複数ファイルを自律的に編集・実行）"],
    ["案件を任せる担当者（成果物は決裁に回る）", "Copilot coding agent（Issue を任せると PR を作成）"],
    ["起案文書のダブルチェック担当", "Copilot code review（PR へのレビューコメント）"],
], [75*mm, 95*mm]))
S.append(PageBreak())

# ---- 1. 機能の全体像 ----
S.append(P("1. GitHub Copilot の機能の全体像", sH1))
S.append(P("Copilot の機能は、<b>人がどこまで主導するか</b>によって段階的に整理できます。下にいくほど AI の自律性が高くなります。"))
S.append(Spacer(1, 4))
S.append(table([
    ["機能", "動作場所", "概要", "人の関与"],
    ["コード補完", "IDE（VS Code、Visual Studio、JetBrains など）", "入力中のコードの続きを灰色の文字で提案。Tab キーで採用。", "1行ごとに人が採否を判断"],
    ["Copilot Chat", "IDE、github.com、モバイル", "自然言語で質問・説明依頼・修正案の作成。コードの解説やテスト作成の相談に。", "提案を人が読んで反映"],
    ["エージェントモード", "IDE", "目的を伝えると、関連ファイルを探して複数ファイルを編集し、ターミナルでコマンドやテストを実行。エラーが出れば自ら修正を繰り返す。", "コマンド実行の承認、最終的な変更の確認"],
    ["Copilot coding agent", "GitHub（クラウド）", "Issue を Copilot に割り当てると、GitHub Actions 上の隔離環境で作業し、ブランチ・コミット・ドラフト PR を作成。", "PR のレビューと承認（マージ）"],
    ["Copilot code review", "GitHub（PR）", "PR の差分を読み、バグの可能性や改善点をレビューコメントとして指摘。", "指摘の採否を人が判断"],
    ["Copilot CLI ほか", "ターミナル", "コマンドラインからエージェントに作業を依頼。", "実行内容の承認"],
], [30*mm, 34*mm, 70*mm, 36*mm]))
S.append(P("<b>注意：</b>機能名・提供範囲（プレビュー／一般提供）・対象プランは頻繁に更新されます。試験導入で使う機能は、"
           "GitHub 公式ドキュメントと自社の契約プランで最新状況を必ず確認してください。", sNote))

S.append(P("2. 「AI アシスタント」と「AI エージェント」の違い", sH1))
S.append(table([
    ["比較項目", "AI アシスタント（補完・チャット）", "AI エージェント（エージェントモード・coding agent）"],
    ["動き方", "1回の質問に1回答える（一問一答）", "目標に向けて、計画 → 実行 → 確認 → 修正を自律的に繰り返す"],
    ["扱う範囲", "主に開いているファイル・選択範囲", "リポジトリ全体（複数ファイルの横断的な変更）"],
    ["使う道具", "なし（テキストを生成するだけ）", "ファイル編集、ターミナル実行、テスト、検索、外部ツール（MCP）"],
    ["成果物", "提案（人がコピー・採用する）", "実際の変更差分・コミット・プルリクエスト"],
    ["人の役割", "書き手（AI は補助）", "依頼者・レビュー担当（AI は作業者）"],
    ["主なリスク", "誤った提案をそのまま採用する", "意図しない広範囲の変更、不要なコマンド実行、外部からの不正な指示（プロンプトインジェクション）"],
    ["必要な統制", "提案内容の確認", "権限の範囲設定、実行環境の隔離、PR レビュー必須、ログ・記録の保持"],
], [26*mm, 62*mm, 82*mm]))
S.append(PageBreak())

# ---- 3. 仕組み ----
S.append(P("3. AI エージェントの仕組み", sH1))
S.append(P("AI エージェントは、大規模言語モデル（LLM）に<b>「道具（ツール）」と「繰り返し（ループ）」</b>を組み合わせたものです。"
           "LLM が次に何をすべきかを判断し、ファイル編集やコマンド実行などの道具を使い、その結果を見てまた次の行動を決めます。"))
S.append(Spacer(1, 6))
S.append(loop_diagram())
S.append(Spacer(1, 4))
S.append(table([
    ["構成要素", "説明", "Copilot での例"],
    ["モデル（LLM）", "状況を理解し、次の行動を決める「頭脳」", "複数の AI モデルから選択可能（プラン・管理者設定による）"],
    ["コンテキスト", "判断材料として与えられる情報", "リポジトリのコード、Issue の本文、指示ファイル、会話履歴"],
    ["ツール", "実際に作業を行う手段", "ファイルの読み書き、ターミナル、テスト実行、Web 検索、MCP サーバー"],
    ["指示ファイル", "プロジェクト固有のルールを AI に伝える文書", C(".github/copilot-instructions.md") + " など（コーディング規約・ビルド方法・禁止事項を記載）"],
    ["MCP", "Model Context Protocol。外部ツールやデータを AI に接続する共通規格", "課題管理ツール、社内 DB、ドキュメントなどとの連携"],
], [28*mm, 62*mm, 80*mm]))
S.append(P("<b>ポイント：</b>エージェントの出来は「モデルの賢さ」だけでなく、<b>与える情報（コンテキスト）の質</b>で大きく変わります。"
           "Issue に目的・完了条件・制約を明確に書くことは、業務委託で仕様書を明確にすることと同じです。", sNote))
S.append(PageBreak())

# ---- 4. coding agent と Git/GitHub ----
S.append(P("4. Copilot coding agent と Git／GitHub の関係", sH1))
S.append(P("Copilot coding agent は、前回の資料で学んだ <b>Git／GitHub の仕組みの上で</b>動きます。AI の作業も通常の開発と同じく"
           "「ブランチ → コミット → プルリクエスト → レビュー → マージ」の流れに乗るため、<b>既存のレビュー・承認フローで統制できる</b>のが特長です。"))
S.append(Spacer(1, 6))
S.append(flow_diagram())
S.append(Spacer(1, 4))
S.append(table([
    ["手順", "内容", "Git／GitHub の概念"],
    ["1. 依頼", "Issue に目的・完了条件を書き、担当者（Assignee）に Copilot を指定。チャット画面などから依頼することも可能。", "Issue"],
    ["2. 作業", "Copilot が専用ブランチを作成し、GitHub Actions 上の隔離された環境でコードを調査・編集・テスト。", "ブランチ、Actions"],
    ["3. 提出", "作業内容をコミットし、ドラフトのプルリクエストを作成。進捗やセッションのログを確認できる。", "コミット、PR"],
    ["4. 修正", "人が PR にレビューコメント（@copilot を付けて指示）すると、Copilot が追加のコミットで対応。", "レビュー"],
    ["5. 承認", "人が内容を確認し、承認・マージ。Copilot 自身はマージできない。", "マージ、ブランチ保護"],
], [20*mm, 115*mm, 35*mm]))
S.append(P("向いているタスクと向いていないタスク", sH2))
S.append(table([
    ["任せやすい（範囲が明確・検証しやすい）", "任せにくい（判断・合意が必要）"],
    ["テストコードの追加、テストカバレッジの改善", "要件があいまいな新機能、業務仕様の決定"],
    ["小さなバグ修正、エラーメッセージの改善", "アーキテクチャの大幅な変更"],
    ["ドキュメント・コメントの整備", "セキュリティ・認証・個人情報処理の中核部分"],
    ["定型的なリファクタリング、ライブラリの更新", "複数部署の合意が必要な変更、本番データの操作"],
], [85*mm, 85*mm], first_col_bold=False))
S.append(PageBreak())

# ---- 5. ガバナンス ----
S.append(P("5. 安全策とガバナンス", sH1))
S.append(P("Copilot coding agent には、GitHub が用意している<b>組み込みの安全策</b>があります（主なもの。詳細は公式ドキュメントで確認）。"))
S += bullets([
    "作業できるのは自分が作成したブランチ（" + C("copilot/") + " で始まる名前）のみで、" + C("main") + " などへ直接 push できない",
    "Copilot が作成した PR は人のレビュー・承認が必要。Copilot に作業を依頼した本人は、その PR を承認できない",
    "Copilot が作成した PR で GitHub Actions のワークフローを実行するには、人の許可が必要",
    "作業環境からのインターネットアクセスはファイアウォールで制限できる",
    "作業のコミットには依頼者の情報が記録され、誰の依頼による変更かを追跡できる",
])
S.append(P("組織として決めること（管理者設定・運用ルール）", sH2))
S.append(table([
    ["項目", "検討内容"],
    ["機能の有効化範囲", "組織・リポジトリ単位で、どの機能（補完・チャット・エージェントモード・coding agent・code review）を誰に許可するか。試験導入は対象リポジトリと対象者を限定するのが安全。"],
    ["ブランチ保護", "main への直接 push 禁止、PR の承認必須、自動テスト合格を必須に設定（前回資料の内容がそのまま統制手段になる）。"],
    ["コンテンツ除外", "機密ファイルを Copilot の参照対象から外す設定。<b>エージェント機能では適用されない場合がある</b>ため、対象機能ごとに最新の対応状況を確認する。"],
    ["パブリックコードとの一致", "公開コードと一致する提案をブロックするか、参照情報付きで許可するかを決める。"],
    ["MCP・外部接続", "どの MCP サーバー（外部ツール連携）を許可するか。許可リスト方式を推奨。"],
    ["シークレット管理", "エージェントの作業環境に渡す認証情報は最小限にする。リポジトリへの秘密情報のコミット禁止を徹底。"],
    ["ログ・監査", "監査ログ、エージェントのセッションログ、PR 履歴を保存し、利用状況を定期的に確認する。"],
], [35*mm, 135*mm]))
S.append(PageBreak())

# ---- 6. 法務・リスク ----
S.append(P("6. 法務・リスクの観点からの確認事項", sH1))
S.append(table([
    ["観点", "リスク", "対応の方向性"],
    ["責任の所在", "AI が作成したコードの不具合の責任が曖昧になる", "「マージを承認した人（組織）が責任を負う」原則を明文化。AI は作業者であり、承認者ではない。"],
    ["著作権・ライセンス", "提案コードが既存の公開コードと一致する可能性", "パブリックコード一致の抑制設定、OSS ライセンス確認手順、提供元の補償条件（契約）の確認。"],
    ["機密情報", "プロンプトや参照ファイルを通じた機密情報の取扱い", "コンテンツ除外、対象リポジトリの限定、入力してはいけない情報のガイドライン化。"],
    ["プロンプトインジェクション", "Issue・Web ページ・外部データに紛れた悪意ある指示に AI が従う", "信頼できる人のみが依頼可能にする、外部アクセスの制限、PR レビューでの差分確認。"],
    ["データの取扱い", "入力・出力データの保存期間や学習利用", "契約プラン（Business／Enterprise 等）の利用規約・データ保護条項・DPA を確認。"],
    ["誤り・品質", "もっともらしいが誤ったコード（ハルシネーション）", "自動テスト必須、人によるレビュー、重要領域は AI 単独作業を禁止。"],
    ["説明責任・監査", "変更理由を後から説明できない", "Issue・PR・コミット・セッションログで「誰が・何を依頼し・誰が承認したか」を記録。"],
], [27*mm, 58*mm, 85*mm]))
S.append(P("<b>ポイント：</b>AI エージェントの統制は、新しい仕組みを一から作るよりも、<b>「依頼（Issue）→ 作業（ブランチ）→ 決裁（PR レビュー）→ 記録（履歴）」</b>"
           "という既存の Git／GitHub のフローに乗せることで実現しやすくなります。", sNote))

S.append(P("7. 試験導入（PoC）の進め方の例", sH1))
S.append(table([
    ["フェーズ", "内容", "成果物"],
    ["1. 準備", "目的・評価指標の設定、対象リポジトリ・参加者の選定、管理者設定とブランチ保護の確認", "導入計画書、利用ガイドライン草案"],
    ["2. 補助機能から開始", "コード補完と Chat を利用し、操作に慣れる", "利用者アンケート"],
    ["3. エージェント機能", "テスト追加など任せやすいタスクで、エージェントモード・coding agent を試行", "タスクごとの評価記録"],
    ["4. 評価・判断", "効果とリスクを評価し、本格導入の可否と運用ルールを決定", "評価報告書、正式ガイドライン"],
], [32*mm, 88*mm, 50*mm]))
S.append(P("評価指標の例", sH2))
S += bullets([
    "<b>効果</b>：作業時間の短縮、PR 作成からマージまでの期間、テストカバレッジの変化、利用者の満足度",
    "<b>品質</b>：Copilot 作成 PR の差し戻し率・修正回数、マージ後の不具合件数",
    "<b>統制</b>：ガイドライン違反の有無、機密情報の混入件数、監査ログの確認結果",
])
S.append(PageBreak())

S.append(P("8. 用語集", sH1))
S.append(table([
    ["用語", "意味"],
    ["AI エージェント", "目標に向けて、計画・ツール実行・確認を自律的に繰り返す AI"],
    ["LLM", "大規模言語モデル。文章やコードを理解・生成する AI の中核"],
    ["プロンプト", "AI への指示文。Issue の本文やチャットの入力など"],
    ["コンテキスト", "AI が判断に使う情報（コード、指示、会話履歴など）"],
    ["エージェントモード", "IDE 内で Copilot が複数ファイルの編集やコマンド実行を自律的に行うモード"],
    ["Copilot coding agent", "Issue を任せるとクラウド上で作業し、PR を作成する Copilot の機能"],
    ["MCP", "Model Context Protocol。AI と外部ツール・データをつなぐ共通規格"],
    ["カスタム指示", "リポジトリ固有のルールを Copilot に伝えるファイル（" + C(".github/copilot-instructions.md") + " など）"],
    ["プロンプトインジェクション", "外部データに埋め込まれた悪意ある指示で AI を誤動作させる攻撃"],
    ["ハルシネーション", "AI がもっともらしいが事実と異なる内容を生成すること"],
    ["Human-in-the-loop", "AI の作業の要所に人の確認・承認を組み込む考え方"],
], [42*mm, 128*mm]))
S.append(P("9. 次のステップ", sH2))
S += bullets([
    "試験導入チームで対象リポジトリ・参加者・評価指標を決める",
    "管理者設定（機能の有効化範囲、コンテンツ除外、パブリックコード一致、MCP）とブランチ保護を確認する",
    "Copilot 利用ガイドライン（任せてよいタスク、入力禁止情報、レビュー責任）を作成する",
    "テスト用リポジトリで、Issue を Copilot に割り当て → PR レビュー → マージを一通り体験する",
    "参考：GitHub Copilot ドキュメント https://docs.github.com/ja/copilot",
])
S.append(P("※ 本資料の機能・設定の説明は作成時点の公開情報に基づく概要です。導入判断の際は公式ドキュメントと契約条件の最新版を確認してください。", sNote))

doc.build(S, onFirstPage=on_page, onLaterPages=on_page)
print("ok", out)
