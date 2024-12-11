import streamlit as st
import pandas as pd
import docx
import io
from converter import MessageConverter

def read_docx(file):
    """DOCXファイルを読み込んでテキスト内容を返す"""
    doc = docx.Document(file)
    content = "\n".join([paragraph.text for paragraph in doc.paragraphs])
    return {"document_content": content}

def main():
    st.title("文書変換アプリ")
    st.write("DOCXファイルを表形式データに変換します")
    
    # ファイルアップロード
    uploaded_files = st.file_uploader(
        "DOCXファイルをアップロードしてください",
        type="docx",
        accept_multiple_files=True
    )
    
    if uploaded_files:
        try:
            # ドキュメントの読み込み
            documents = [read_docx(file) for file in uploaded_files]
            
            # 変換処理
            converter = MessageConverter()
            df = converter.process_documents(documents, st)
            
            # 結果の表示
            st.subheader("変換結果")
            st.dataframe(df)
            
            # Excelファイルのダウンロード
            excel_buffer = io.BytesIO()
            df.to_excel(excel_buffer, index=False)
            excel_data = excel_buffer.getvalue()
            
            st.download_button(
                label="Excelファイルをダウンロード",
                data=excel_data,
                file_name="converted_data.xlsx",
                mime="application/vnd.ms-excel"
            )
            
        except Exception as e:
            st.error(f"エラーが発生しました: {str(e)}")

if __name__ == "__main__":
    main()
