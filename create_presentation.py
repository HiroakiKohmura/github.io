#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# プレゼンテーションの作成
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# カラースキーム
TITLE_COLOR = RGBColor(102, 126, 234)  # #667eea
ACCENT_COLOR = RGBColor(118, 75, 162)  # #764ba2
TEXT_COLOR = RGBColor(44, 62, 80)      # #2c3e50
WHITE = RGBColor(255, 255, 255)

def add_title_slide(prs, title, subtitle):
    """タイトルスライドを追加"""
    slide_layout = prs.slide_layouts[6]  # 空白レイアウト
    slide = prs.slides.add_slide(slide_layout)

    # 背景色の設定
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = TITLE_COLOR

    # タイトル
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(60)
    title_para.font.bold = True
    title_para.font.color.rgb = WHITE
    title_para.alignment = PP_ALIGN.CENTER

    # サブタイトル
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = subtitle
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.font.size = Pt(32)
    subtitle_para.font.color.rgb = WHITE
    subtitle_para.alignment = PP_ALIGN.CENTER

def add_content_slide(prs, title, content_items):
    """コンテンツスライドを追加"""
    slide_layout = prs.slide_layouts[6]  # 空白レイアウト
    slide = prs.slides.add_slide(slide_layout)

    # タイトル
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR

    # タイトル下線
    line = slide.shapes.add_shape(
        1,  # 長方形
        Inches(0.5), Inches(1.4),
        Inches(9), Inches(0.05)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = TITLE_COLOR
    line.line.fill.background()

    # コンテンツ
    content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    text_frame = content_box.text_frame
    text_frame.word_wrap = True

    for item in content_items:
        p = text_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(18)
        p.font.color.rgb = TEXT_COLOR
        p.level = 0
        p.space_after = Pt(12)

# スライド1: タイトル
add_title_slide(prs, "ペロブスカイト太陽電池", "次世代の太陽光発電技術")

# スライド2: 目次
add_content_slide(prs, "目次", [
    "1. ペロブスカイト太陽電池とは",
    "2. 発展の歴史",
    "3. 基本構造",
    "4. 動作原理",
    "5. 主な利点",
    "6. 克服すべき課題",
    "7. 最新の研究動向",
    "8. タンデム太陽電池の可能性",
    "9. 将来展望",
    "10. 日本における取り組み",
    "11. まとめ"
])

# スライド3: ペロブスカイト太陽電池とは
add_content_slide(prs, "ペロブスカイト太陽電池とは", [
    "【定義】",
    "• ペロブスカイト構造を持つ材料を光吸収層に使用した太陽電池",
    "",
    "【ペロブスカイト構造】",
    "• 化学式: ABX₃",
    "  - A: 有機カチオン（CH₃NH₃⁺など）",
    "  - B: 金属カチオン（Pb²⁺、Sn²⁺など）",
    "  - X: ハロゲンイオン（I⁻、Br⁻、Cl⁻）",
    "",
    "【代表例】",
    "• CH₃NH₃PbI₃（メチルアンモニウム鉛ヨウ化物）"
])

# スライド4: 発展の歴史
add_content_slide(prs, "発展の歴史", [
    "2009年",
    "• 桐蔭横浜大学の宮坂力教授らが初めて開発",
    "• 変換効率: 3.8%",
    "",
    "2012年",
    "• 固体型ペロブスカイト太陽電池の開発",
    "• 変換効率: 10%超",
    "",
    "2015年",
    "• 変換効率: 20%突破",
    "",
    "2023年",
    "• 実験室レベルで変換効率: 26%以上達成",
    "• シリコン太陽電池に匹敵する性能"
])

# スライド5: 基本構造
add_content_slide(prs, "基本構造", [
    "【層構造（上から下へ）】",
    "1. 透明電極（ITO/FTO）",
    "2. 電子輸送層（TiO₂など）",
    "3. ペロブスカイト光吸収層 ★",
    "4. 正孔輸送層（Spiro-OMeTADなど）",
    "5. 金属電極（Au、Agなど）",
    "",
    "【特徴】",
    "• 薄膜構造（数百nm程度）",
    "• 低温プロセスで製造可能",
    "• フレキシブル基板にも対応可能"
])

# スライド6: 動作原理
add_content_slide(prs, "動作原理", [
    "【ステップ1】光吸収",
    "• ペロブスカイト層が太陽光を吸収",
    "• 励起子（電子-正孔対）を生成",
    "",
    "【ステップ2】電荷分離",
    "• 励起子が電子と正孔に分離",
    "",
    "【ステップ3】電荷輸送",
    "• 電子 → 電子輸送層へ移動",
    "• 正孔 → 正孔輸送層へ移動",
    "",
    "【ステップ4】電流生成",
    "• 電極で電荷を収集し、外部回路に電流を供給"
])

# スライド7: 主な利点（1/2）
add_content_slide(prs, "主な利点（1/2）", [
    "⚡ 高い変換効率",
    "• シリコン太陽電池に匹敵する26%以上の変換効率を達成",
    "",
    "💰 低コスト",
    "• 溶液プロセスで製造可能",
    "• 安価な材料を使用",
    "",
    "🌡️ 低温製造",
    "• 150℃以下の低温で製造可能",
    "• （シリコンは1000℃以上が必要）"
])

# スライド8: 主な利点（2/2）
add_content_slide(prs, "主な利点（2/2）", [
    "📱 柔軟性",
    "• フレキシブル基板への適用が可能",
    "• 曲げられる太陽電池の実現",
    "",
    "🎨 カラーバリエーション",
    "• 材料の組成変更により色の調整が可能",
    "• 建材一体型への応用に有利",
    "",
    "🔬 簡易な製造プロセス",
    "• 塗布やスピンコートなど簡単な方法で製造可能"
])

# スライド9: 克服すべき課題
add_content_slide(prs, "克服すべき課題", [
    "🔴 【重要課題1】耐久性・安定性",
    "• 湿度、熱、紫外線に弱い",
    "• 長期安定性が不十分（劣化が早い）",
    "• 実用化には20-25年の寿命が必要",
    "",
    "🔴 【重要課題2】鉛の使用",
    "• 高効率な材料には鉛が含まれる",
    "• 環境・健康への影響が懸念",
    "• 鉛フリー代替材料の開発が必要",
    "",
    "🟡 大面積化",
    "• 小型セルでは高効率だが、大型化で効率低下",
    "",
    "🟡 製造プロセスの最適化",
    "• 再現性の確保、量産技術の確立"
])

# スライド10: 最新の研究動向
add_content_slide(prs, "最新の研究動向", [
    "【材料開発】",
    "• 鉛フリーペロブスカイト（Snベース、Biベース）",
    "• 2D/3Dハイブリッド構造",
    "• 混合ハロゲン系ペロブスカイト",
    "",
    "【安定性向上】",
    "• 封止技術の改善、添加剤による安定化、界面工学",
    "",
    "【タンデム太陽電池】",
    "• シリコン/ペロブスカイトタンデム",
    "• 変換効率30%以上を目指す、実用化に最も近い技術",
    "",
    "【新応用分野】",
    "• 建材一体型（BIPV）、透明太陽電池、IoTデバイス用電源"
])

# スライド11: タンデム太陽電池の可能性
add_content_slide(prs, "タンデム太陽電池の可能性", [
    "【タンデム構造とは】",
    "• 異なるバンドギャップを持つ複数の太陽電池を積層",
    "• ペロブスカイト層（高エネルギー光吸収）",
    "  + シリコン層（低エネルギー光吸収）",
    "",
    "【メリット】",
    "• 太陽光スペクトルをより効率的に利用",
    "• 理論変換効率: 35%以上",
    "• 既存のシリコン技術と組み合わせ可能",
    "",
    "【最新記録】",
    "• 33.9%の変換効率達成（2024年）"
])

# スライド12: 将来展望
add_content_slide(prs, "将来展望", [
    "【短期（2025-2027年）】",
    "• タンデム太陽電池の商業化開始",
    "• 耐久性10年以上の製品登場、特殊用途での実用化",
    "",
    "【中期（2028-2032年）】",
    "• 鉛フリー材料の実用化",
    "• 大面積モジュールの量産開始、建材一体型製品の普及",
    "",
    "【長期（2033年以降）】",
    "• シリコン太陽電池市場の一部代替",
    "• 変換効率35%以上の達成",
    "• 新しいエネルギーインフラの一翼を担う",
    "",
    "【市場予測】2030年までに数十億ドル規模に成長"
])

# スライド13: 日本における取り組み
add_content_slide(prs, "日本における取り組み", [
    "【研究機関】",
    "• 桐蔭横浜大学: ペロブスカイト太陽電池の発祥地",
    "• 東京大学: 高効率化・安定化研究",
    "• NIMS: 材料開発",
    "• 産業技術総合研究所: 実用化研究",
    "",
    "【企業】",
    "• パナソニック: タンデム太陽電池開発",
    "• 積水化学: フィルム型太陽電池",
    "• 東芝: 大面積化技術",
    "• リコー: インクジェット塗布技術",
    "",
    "日本は基礎研究から実用化まで、世界をリード"
])

# スライド14: まとめ
add_content_slide(prs, "まとめ", [
    "【強み】",
    "✓ 高い変換効率（26%以上）",
    "✓ 低コスト製造",
    "✓ 柔軟な応用可能性",
    "✓ タンデム化による更なる高効率化",
    "",
    "【課題】",
    "✗ 長期安定性の確保",
    "✗ 鉛の代替材料開発",
    "✗ 大面積化技術、量産プロセスの確立",
    "",
    "【結論】",
    "ペロブスカイト太陽電池は、多くの課題を抱えながらも、",
    "次世代の太陽光発電技術として大きな可能性を秘めています。",
    "今後5-10年の研究開発が、実用化の鍵となります。"
])

# 最終スライド: 参考資料
add_content_slide(prs, "参考資料・さらに学ぶために", [
    "【主要な学術論文】",
    "• Nature Energy、Nature Photonicsなどの学術誌",
    "• NREL（米国再生可能エネルギー研究所）の効率チャート",
    "",
    "【日本の研究機関】",
    "• 桐蔭横浜大学 先端材料科学研究所",
    "• NEDO、JSTのプロジェクト情報",
    "",
    "【国際会議】",
    "• HOPV (Hybrid and Organic Photovoltaics)",
    "• MRS (Materials Research Society) Meeting",
    "",
    "",
    "ご清聴ありがとうございました"
])

# ファイルを保存
prs.save('ペロブスカイト太陽電池.pptx')
print("PowerPointファイルを作成しました: ペロブスカイト太陽電池.pptx")
