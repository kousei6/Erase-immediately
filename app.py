import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# ページの設定
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄")

def extract_text_from_pdf(pdf_file):
    """
    PDFファイルからテキストを抽出し、ページごとのリストを返す
    """
    extracted_data = []
    with pdfplumber.open(pdf_file) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                extracted_data.append({
                    "Page": i + 1,
                    "Content": text
                })
    return extracted_data

def convert_to_excel(data):
    """
    抽出したデータをExcelファイル（バイトデータ）に変換する
    """
    df = pd.DataFrame(data)
    output = BytesIO()
    # Excelライブラリとしてopenpyxlを使用（pandasの依存関係）
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracted Text')
    return output.getvalue()

def main():
    st.title("📄 PDFテキスト抽出 & Excel変換ツール")
    st.write("PDFファイルをアップロードすると、テキストを抽出してExcelファイルとしてダウンロードできます。")

    # ファイルアップローダー
    uploaded_file = st.file_uploader("PDFファイルをアップロードしてください", type="pdf")

    if uploaded_file is not None:
        with st.spinner("テキストを抽出中..."):
            try:
                # テキスト抽出実行
                extracted_data = extract_text_from_pdf(uploaded_file)
                
                if not extracted_data:
                    st.warning("テキストを抽出できませんでした。スキャンされた画像形式のPDFの可能性があります。")
                    return

                # 抽出結果の表示（プレビュー）
                st.subheader("抽出テキストのプレビュー")
                for item in extracted_data[:3]:  # 最初の3ページ分のみ表示
                    with st.expander(f"ページ {item['Page']}"):
                        st.text(item['Content'])
                
                if len(extracted_data) > 3:
                    st.info(f"ほか {len(extracted_data) - 3} ページが抽出されました。")

                # Excel変換
                excel_data = convert_to_excel(extracted_data)

                # ダウンロードボタン
                st.success("抽出が完了しました！")
                st.download_button(
                    label="Excelファイルをダウンロード",
                    data=excel_data,
                    file_name=f"extracted_text_{uploaded_file.name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()
