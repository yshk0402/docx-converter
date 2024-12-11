import streamlit as st
import pandas as pd
import docx
import io
import openpyxl
from typing import List, Dict
from pathlib import Path
from converter import MessageConverter, ProcessingError
from config import Config
import time

class DocxConverterApp:
    def __init__(self):
        self.converter = MessageConverter()
        self.config = Config()
        self.setup_streamlit()

    def setup_streamlit(self):
        """Streamlit設定の初期化"""
        st.set_page_config(
            page_title="DOCX Converter",
            page_icon="📄",
            layout="wide"
        )
        if 'processed_df' not in st.session_state:
            st.session_state.processed_df = None
        if 'errors' not in st.session_state:
            st.session_state.errors = []

    def load_favorite_columns(self) -> List[str]:
        """お気に入りカラムの読み込み"""
        config = self.config.load_config()
        return config.get('favorite_columns', [])

    def save_favorite_columns(self, columns: List[str]):
        """お気に入りカラムの保存"""
        self.config.update_favorite_columns(columns)

    def read_docx(self, file) -> Dict:
        """DOCXファイルの読み込み"""
        doc = docx.Document(file)
        content = "\n".join([paragraph.text for paragraph in doc.paragraphs])
        return {"document_content": content}

    def create_excel_with_highlights(self, df: pd.DataFrame, errors: List[ProcessingError]) -> bytes:
        """エラーハイライト付きExcelの作成"""
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='データ')
            workbook = writer.book
            worksheet = writer.sheets['データ']
            
            # エラー行の背景色を設定
            error_rows = set(error.document_index for error in errors)
            for row_idx in error_rows:
                for cell in worksheet[row_idx + 2]:  # Excel行は1から始まり、ヘッダーがあるため+2
                    cell.fill = openpyxl.styles.PatternFill(
                        start_color='FFE7E6',
                        end_color='FFE7E6',
                        fill_type='solid'
                    )
        
        return output.getvalue()

    def display_preview(self, content: str):
        """文書のプレビュー表示"""
        st.subheader("プレビュー")
        with st.expander("文書内容", expanded=True):
            st.text(content[:500] + "..." if len(content) > 500 else content)

    def display_errors(self, errors: List[ProcessingError]):
        """エラーの表示"""
        if errors:
            st.error("処理中にエラーが発生しました")
            for error in errors:
                with st.expander(f"文書 {error.document_index + 1} のエラー"):
                    st.write(f"種類: {error.error_type}")
                    st.write(f"メッセージ: {error.message}")
                    if error.details:
                        st.write(f"詳細: {error.details}")

    def run(self):
        """アプリケーションのメイン処理"""
        st.title("文書変換アプリ")

        # ファイルアップロード
        uploaded_files = st.file_uploader(
            "DOCXファイルをアップロードしてください",
            type="docx",
            accept_multiple_files=True
        )

        if uploaded_files:
            # プログレスバーの表示
            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                # ドキュメントの読み込みと処理
                documents = []
                for i, file in enumerate(uploaded_files):
                    doc = self.read_docx(file)
                    documents.append(doc)
                    
                    # プレビュー表示（最初のファイルのみ）
                    if i == 0:
                        self.display_preview(doc['document_content'])
                    
                    # 進捗更新
                    progress = (i + 1) / len(uploaded_files)
                    progress_bar.progress(progress)
                    status_text.text(f"処理中... {i + 1}/{len(uploaded_files)} ファイル")
                    time.sleep(0.1)  # UIの更新のため

                # カラム選択
                favorite_columns = self.load_favorite_columns()
                all_columns = self.converter.extract_potential_columns(documents[0]['document_content'])
                
                col1, col2 = st.columns(2)
                with col1:
                    selected_columns = st.multiselect(
                        "使用するカラムを選択：",
                        list(all_columns),
                        default=favorite_columns if favorite_columns else ['番号', '所属', '名前', '原稿']
                    )
                
                with col2:
                    if st.button("カラム設定を保存"):
                        self.save_favorite_columns(selected_columns)
                        st.success("カラム設定を保存しました")

                self.converter.selected_columns = selected_columns

                # データ処理
                df, errors = self.converter.process_documents(documents)
                
                # エラー表示
                self.display_errors(errors)

                # 結果表示
                if not df.empty:
                    st.subheader("変換結果")
                    
                    # データフレームの編集機能
                    edited_df = st.data_editor(
                        df,
                        num_rows="dynamic",
                        use_container_width=True
                    )
                    
                    # エクセルファイルのダウンロード
                    excel_data = self.create_excel_with_highlights(edited_df, errors)
                    st.download_button(
                        label="Excelファイルをダウンロード",
                        data=excel_data,
                        file_name="converted_data.xlsx",
                        mime="application/vnd.ms-excel"
                    )

            except Exception as e:
                st.error(f"予期せぬエラーが発生しました: {str(e)}")
            
            finally:
                # プログレスバーを完了状態に
                progress_bar.progress(1.0)
                status_text.text("処理完了")

if __name__ == "__main__":
    app = DocxConverterApp()
    app.run()
