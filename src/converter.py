import pandas as pd
import re
from typing import Dict, List, Set
from datetime import datetime

class MessageConverter:
    def __init__(self):
        self.default_columns = ['番号']  # 番号は常に含める
        self.selected_columns = None
        
    def extract_potential_columns(self, content: str) -> Set[str]:
        """DOCXコンテンツから潜在的なカラム候補を抽出する"""
        columns = set()
        
        # 正規表現パターンでキーと値のペアを検索
        pattern = r'【(.+?)】|^(.+?)：'  # 【】で囲まれた文字列、または：で終わる文字列
        matches = re.finditer(pattern, content, re.MULTILINE)
        
        for match in matches:
            column_name = match.group(1) or match.group(2)
            if column_name:
                column_name = column_name.strip()
                columns.add(column_name)
                
        # デフォルトで含めたい一般的なカラムを追加
        common_columns = {'所属', '名前', '原稿', '勤続年数', '入社年', '部署'}
        columns.update(common_columns)
        
        return columns

    def prompt_column_selection(self, potential_columns: Set[str], st) -> List[str]:
        """Streamlit用にカラム選択機能を提供する"""
        st.subheader("カラム選択")
        
        column_list = sorted(list(potential_columns))
        selected_columns = st.multiselect(
            "使用するカラムを選択してください：",
            column_list,
            default=['所属', '名前', '原稿']
        )
        
        return self.default_columns + selected_columns

    def extract_info_from_docx(self, content: str, columns: List[str]) -> Dict:
        """DOCXコンテンツから必要な情報を抽出する"""
        info = {column: '' for column in columns}
        
        # 各行を解析
        lines = content.split('\n')
        for line in lines:
            line = line.strip()
            
            # キーと値のペアを検出
            for column in columns:
                patterns = [
                    f'{column}：\s*(.*?)(?=\n|$)',
                    f'【{column}】\s*(.*?)(?=\n|$)',
                    f'^{column}\s+(.*?)(?=\n|$)'
                ]
                
                for pattern in patterns:
                    match = re.search(pattern, line)
                    if match:
                        info[column] = match.group(1).strip()
                        break
        
        # 特別な処理が必要なカラム
        if '勤続年数' in columns and '入社年' in info:
            info['勤続年数'] = self._calculate_years_worked(info['入社年'])
            
        # 原稿の抽出（本文部分）
        if '原稿' in columns:
            message_lines = []
            message_started = False
            for line in lines:
                if re.search(r'\d+$', line.strip()) and not message_started:
                    message_started = True
                    continue
                if message_started and line.strip() and not re.search(r'^\d+$', line.strip()):
                    message_lines.append(line.strip())
            
            info['原稿'] = ' '.join(message_lines).strip()
            
        return info
    
    def _calculate_years_worked(self, entry_year: str) -> str:
        """入社年から勤続年数を計算する"""
        if not entry_year:
            return ''
            
        try:
            if '昭和' in entry_year:
                year = int(entry_year.replace('昭和', '')) + 1925
            elif '平成' in entry_year:
                year = int(entry_year.replace('平成', '')) + 1989
            else:
                year = int(entry_year)
                
            years_worked = datetime.now().year - year
            return f'{years_worked}年'
        except:
            return ''
    
    def create_table_row(self, idx: int, info: Dict) -> Dict:
        """1行分のテーブルデータを作成する"""
        row = {'番号': idx + 1}
        row.update(info)
        return row
    
    def process_documents(self, documents: List[Dict], st) -> pd.DataFrame:
        """複数のドキュメントを処理してDataFrameを作成する"""
        # 最初のドキュメントからカラム候補を抽出
        first_doc = documents[0]['document_content']
        potential_columns = self.extract_potential_columns(first_doc)
        
        # まだカラムが選択されていない場合、ユーザーに選択させる
        if self.selected_columns is None:
            self.selected_columns = self.prompt_column_selection(potential_columns, st)
        
        rows = []
        for idx, doc in enumerate(documents):
            content = doc['document_content']
            info = self.extract_info_from_docx(content, self.selected_columns)
            row = self.create_table_row(idx, info)
            rows.append(row)
            
        return pd.DataFrame(rows, columns=self.selected_columns)
