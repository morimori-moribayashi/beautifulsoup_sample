from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def extract_sections(soup):
    """
    HTMLからsectionタグの内容を抽出する関数
    
    Parameters:
    soup (BeautifulSoup): 解析対象のBeautifulSoupオブジェクト
    
    Returns:
    list: 各sectionの内容を辞書形式で格納したリスト
    """
    # すべてのsectionタグを取得
    sections = soup.find_all('section')
    
    # 結果を格納するリスト
    results = []
    
    # 各sectionタグの処理
    for section in sections:
        # 見出し（h2, h3等）の取得
        heading = section.find(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
        heading_text = heading.get_text().strip() if heading else "見出しなし"
        
        # 段落（p）の取得
        paragraphs = section.find_all('p')
        paragraph_texts = [p.get_text().strip() for p in paragraphs]
        
        # リスト（ul, ol）の取得
        lists = section.find_all(['ul', 'ol'])
        list_items = []
        for list_tag in lists:
            items = list_tag.find_all('li')
            list_items.extend([item.get_text().strip() for item in items])
        
        # テーブル（table）の取得
        tables = section.find_all('table')
        table_data = []
        for table in tables:
            rows = table.find_all('tr')
            table_rows = []
            for row in rows:
                headers = row.find_all('th')
                if headers:
                    table_rows.append([th.get_text().strip() for th in headers])
                cells = row.find_all('td')
                if cells:
                    table_rows.append([td.get_text().strip() for td in cells])
            table_data.append(table_rows)
        
        # 結果を辞書形式で格納
        section_data = {
            "title": heading_text,
            "paragraphs": paragraph_texts,
            "list_items": list_items,
            "tables": table_data
        }
        
        results.append(section_data)
    
    return results

def create_title_slide(prs, title_text):
    """タイトルスライドを作成する関数"""
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    title.text = title_text
    
    # タイトルテキストのフォーマット調整
    for paragraph in title.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.bold = True
            run.font.size = Pt(40)
    
    return slide

def add_text_to_textbox(textbox, text, font_size=Pt(20), level=0):
    """テキストボックスにテキストを追加する関数"""
    p = textbox.text_frame.add_paragraph()
    p.text = text
    p.font.size = font_size
    p.level = level
    return p

def create_content_slide(prs, section_data):
    """セクションデータからコンテンツスライドを作成する関数"""
    # タイトルとコンテンツのレイアウトを使用
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    # タイトル設定
    title = slide.shapes.title
    title.text = section_data["title"]
    
    # コンテンツ領域の取得
    content = slide.placeholders[1]
    tf = content.text_frame
    tf.clear()  # デフォルトテキストをクリア
    
    # 段落の追加
    for k,paragraph_text in enumerate(section_data["paragraphs"]):
        # 最初の段落はテキストボックスのテキストとして設定
        if k==0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        # テキストボックスのテキストを設定
        p.text = paragraph_text
        p.font.size = Pt(20)
    
    # リスト項目の追加
    if section_data["list_items"]:
        # テキストボックスの下にリストを配置するための位置計算
        for item in section_data["list_items"]:
            p = tf.add_paragraph()
            p.text = item
            p.font.size = Pt(20)
            p.level = 1  # 箇条書きのレベル
    
    # テーブルの追加
    if section_data["tables"]:
        # テキストボックスの下にテーブルを配置するための位置計算
        top = Inches(3.5)  # 適宜調整
        
        for table_rows in section_data["tables"]:
            if not table_rows:
                continue
                
            rows = len(table_rows)
            cols = len(table_rows[0]) if table_rows else 0
            
            if rows == 0 or cols == 0:
                continue
                
            # テーブルの追加
            left = Inches(1)
            width = Inches(8)
            height = Inches(rows * 0.5)
            
            # スライドにテーブル追加
            table = slide.shapes.add_table(rows, cols, left, top, width, height).table
            
            # テーブル内容の設定
            for i, row_cells in enumerate(table_rows):
                for j, cell_text in enumerate(row_cells):
                    if j < cols:  # 列数チェック
                        cell = table.cell(i, j)
                        cell.text = cell_text
                        
                        # セルテキストのフォーマット
                        for paragraph in cell.text_frame.paragraphs:
                            paragraph.font.size = Pt(16)
                            if i == 0:  # ヘッダー行
                                paragraph.font.bold = True
                                
            top += height + Inches(0.5)  # 次のテーブルのための位置調整

    return slide

def convert_html_to_ppt(html_content, output_filename):
    """HTMLからPowerPointプレゼンテーションを作成する関数"""
    try:
        # BeautifulSoupでHTML解析
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # プレゼンテーション作成
        prs = Presentation()
        
        # タイトル取得
        first_h1 = soup.find('h1')
        if first_h1:
            title_text = first_h1.text.strip()
        else:
            title_text = "プレゼンテーション"
            
        # タイトルスライド作成
        create_title_slide(prs, title_text)
        
        # セクションデータ抽出
        sections = extract_sections(soup)
        
        # 各セクションをスライドに変換
        for section in sections:
            create_content_slide(prs, section)
        
        # プレゼンテーション保存
        prs.save(output_filename)
        print(f"プレゼンテーションを '{output_filename}' に保存しました。")
        return True
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        return False

# HTMLコンテンツ
html = '''
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>2024年果汁飲料市場トレンド分析と戦略提案</title>
</head>
<body>
    <header>
        <h1>2024年果汁飲料市場トレンド分析と戦略提案</h1>
    </header>
    
    <main>
        <section>
            <h2>1. 市場概況</h2>
            <p>2024年の果汁飲料市場は、健康志向の高まりとサステナビリティへの関心増加を背景に、大きな変化を遂げています。市場規模は前年比6.5%増の9,320億円に拡大し、特にオーガニック製品（27%、+9ポイント）、低糖・無糖製品（45%、+13ポイント）、機能性表示食品（22%、+8ポイント）の比率が顕著に上昇しています。</p>
        </section>

        <section>
            <h2>2. 主要トレンド</h2>
            <p>果汁飲料市場における主要なトレンドは以下の3つに分類されます。</p>
        </section>

        <section>
            <h3>2.1 健康・機能性重視の傾向</h3>
            <p>消費者の健康意識の高まりにより、単なる果汁飲料から機能性を付加した製品への移行が進んでいます。特に以下のトレンドが顕著です：</p>
            <ul>
                <li>低糖・無糖志向：天然甘味料の活用、糖質オフ表示</li>
                <li>機能性付加：ビタミン、ミネラル、食物繊維の強化</li>
                <li>スーパーフード配合：アサイー、チアシード、マキベリーなど</li>
                <li>発酵果汁：プロバイオティクス、コンブチャなど</li>
            </ul>
        </section>

        <section>
            <h3>2.2 サステナビリティへの取り組み</h3>
            <p>環境配慮型の製品開発や包装が重要視されています。バイオマスPETボトルの導入率は前年の22%から38%へ、リサイクル素材使用は45%から63%へと大幅に上昇しています。地産地消原料調達も31%から47%へと拡大し、カーボンフットプリント表示も徐々に普及しています。</p>
        </section>

        <section>
            <h3>2.3 フレーバー・味覚の多様化</h3>
            <table border="1">
                <tr>
                    <th>カテゴリー</th>
                    <th>人気フレーバー</th>
                    <th>前年比成長率</th>
                </tr>
                <tr>
                    <td>エキゾチック</td>
                    <td>ドラゴンフルーツ、パッションフルーツ</td>
                    <td>+32%</td>
                </tr>
                <tr>
                    <td>和風フレーバー</td>
                    <td>柚子、桜、山椒</td>
                    <td>+28%</td>
                </tr>
                <tr>
                    <td>クラフト系</td>
                    <td>スモーキー、ハーブ調合</td>
                    <td>+45%</td>
                </tr>
                <tr>
                    <td>ハイブリッド</td>
                    <td>果物×野菜、果物×スパイス</td>
                    <td>+37%</td>
                </tr>
            </table>
        </section>

        <section>
            <h2>3. 消費者行動分析</h2>
            <p>消費シーンや購買動機が多様化しています。Z世代（18-25歳）はSNS映えや体験価値、サステナビリティを重視し、ミレニアル世代（26-41歳）は健康管理とワークライフバランスを意識した選択をする傾向があります。X世代（42-57歳）は機能性と品質、シニア層（58歳以上）は健康維持と原材料の安全性を重視しています。</p>
        </section>

        <section>
            <h2>4. 今後の戦略提案</h2>
            <p>これらのトレンドを踏まえ、以下の戦略展開を提案します。</p>
        </section>

        <section>
            <h3>4.1 製品開発戦略</h3>
            <ul>
                <li>機能性強化シリーズ：特定保健用食品・機能性表示食品の拡充</li>
                <li>低糖質プレミアムライン：自然由来甘味料使用、糖質50%オフ</li>
                <li>季節限定フレーバー：四季に合わせた限定商品の展開</li>
                <li>地域特産果実シリーズ：各地の特産品を活かした地域限定商品</li>
            </ul>
        </section>

        <section>
            <h3>4.2 サステナビリティ推進計画</h3>
            <p>2025年第2四半期までに100%リサイクルPETの導入、2025年通年でのCO2削減30%、2026年第1四半期までにフェアトレード原料100%の実現、2025年第3四半期からの水資源保全プロジェクト開始を計画しています。</p>
        </section>

        <section>
            <h3>4.3 マーケティング戦略</h3>
            <p>SNSを活用したインフルエンサーコラボやUGC促進、産地見学ツアーなどの体験型イベント、医師・栄養士との共同研究発表による機能性訴求、パーソナライズド推奨システムによるデジタルDMなど、各ターゲット層に合わせた施策を展開します。</p>
        </section>

        <section>
            <h2>5. 業績予測</h2>
            <p>2025年は売上高前年比+8%、営業利益率12%、2026年は売上高前年比+12%、営業利益率15%を見込んでいます。研究開発投資は売上高の6～7%、サステナビリティ投資は4～5%を計画しています。</p>
        </section>

        <section>
            <h2>6. まとめ</h2>
            <p>2024年の果汁飲料市場は、健康志向とサステナビリティへの関心が牽引する形で成長を続けています。当社としては、機能性とサステナビリティを両立させた製品開発を進めるとともに、消費者との新しい関係構築を目指したマーケティング活動を展開することで、業界のリーディングカンパニーとしての地位を確立していきます。</p>
        </section>
    </main>
</body>
</html>
'''

if __name__ == "__main__":
    convert_html_to_ppt(html,"sample.pptx")