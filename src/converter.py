import pandas as pd
import re
from typing import Dict, List, Set, Tuple
from datetime import datetime
import markdown
from bs4 import BeautifulSoup
from dataclasses import dataclass

@dataclass
class ProcessingError:
    document_index: int
    error_type: str
    message: str
    details: str = ""

class DocumentProcessor:
    def __init__(self):
        self.errors: List[ProcessingError] = []
        
    def convert_to_markdown(self, content: str) -> str:
        """文書をMarkdown形式に変換"""
        try:
            # 特殊文字のエスケープ
            content = content.replace('*', '\*')
            # 見出しの変換
            content = re.sub(r'【(.+?)】', r'### \1', content)
            
            lines = []
            for line in content.split('\n'):
                line = line.strip()
                if line.startswith('+-') or line.startswith('|='):
                    continue
                if '：' in line and not line.startswith('|'):
                    key, value = line.split('：', 1)
                    lines.append(f'### {key.strip()}\n{value.strip()}')
                elif line:
                    lines.append(line)
            
            return '\n\n'.join(lines)
        except Exception as e:
            raise ValueError(f"Markdown変換エラー: {str(e)}")

    def extract_structure(self, markdown_content: str) -> Dict:
        """Markdownから構造を抽出"""
        try:
            html = markdown.markdown(markdown_content)
            soup = BeautifulSoup(html, 'html.parser')
            
            structure = {
                'h1': [], 'h2': [], 'h3': [], 'text': []
            }
            
            current_section = None
            for element in soup.find_all(['h1', 'h2', 'h3', 'p']):
                if element.name in ['h1', 'h2', 'h3']:
                    text = element.get_text().strip()
                    structure[element.name].append({
                        'text': text,
                        'content': []
                    })
                    current_section = structure[element.name][-1]
                elif element.name == 'p':
                    text = element.get_text().strip()
                    if current_section:
                        current_section['content'].append(text)
                    else:
                        structure['text'].append(text)
            
            return structure
        except Exception as e:
            raise ValueError(f"構造抽出エラー: {str(e)}")

class MessageConverter:
    def __init__(self):
        self.default_columns = ['番号']
        self.selected_columns = None
        self.processor = DocumentProcessor()
        self.validation_rules = {
            '原稿': lambda x: 150 <= len(x) <= 200 if x else False,
            '名前': lambda x: bool(x and x.strip()),
            '部署': lambda x: bool(x and x.strip()),
        }

    def _clean_text(self, text: str) -> str:
        """テキストのクリーニング"""
        if not text:
            return ""
        # 余分な空白を削除
        text = re.sub(r'\s+', ' ', text.strip())
        # 特殊文字を削除
        text = re.sub(r'[\r\n\t\f\v]', '', text)
        return text

    def extract_potential_columns(self, content: str) -> Set[str]:
        """DOCXコンテンツから実際に存在するカラムのみを抽出"""
        columns = set(['番号'])  # 基本カラム
        
        # テーブル内のカラム抽出（改善版）
        table_pattern = r'\|(.*?)\|'
        table_rows = re.findall(table_pattern, content, re.MULTILINE)
        
        for row in table_rows:
            # セル内のテキストを抽出
            cells = [cell.strip() for cell in row.split('|')]
            for cell in cells:
                # カラム名のパターンを検出
                if '：' in cell or ':' in cell:
                    key = cell.split('：')[0].split(':')[0].strip()
                    if key:
                        normalized = self._normalize_column_name(key)
                        if normalized:
                            columns.add(normalized)

        # 見出しからのカラム抽出
        header_patterns = [
            (r'【(.+?)】', 1),            # 【見出し】
            (r'■\s*(.+?)\s*[:：]', 1),    # ■見出し：
            (r'□\s*(.+?)\s*[:：]', 1),    # □見出し：
            (r'●\s*(.+?)\s*[:：]', 1),    # ●見出し：
            (r'^(?:■|□|●)?\s*(.+?)\s*[:：](?!\d)', 1),  # キー：（時刻を除外）
        ]
        
        for line in content.split('\n'):
            line = line.strip()
            for pattern, group in header_patterns:
                match = re.search(pattern, line)
                if match:
                    column_name = match.group(group).strip()
                    if column_name:
                        normalized = self._normalize_column_name(column_name)
                        if normalized:
                            columns.add(normalized)
        
        # 長文検出による原稿カラムの追加
        paragraphs = content.split('\n')
        main_content = self._extract_main_content(paragraphs)
        if main_content and len(main_content) >= 20:
            columns.add('原稿')
        
        # 不要なカラムの除外
        columns = {col for col in columns if not any(skip in col.lower() for skip in [
            'について', 'お願い', 'です', 'ます', 'した', 'から', '注意', '備考'
        ])}
        
        return columns

    def _normalize_column_name(self, name: str) -> str:
        """カラム名を正規化"""
        name = name.strip()
        
        # 特定のプレフィックスとサフィックスを削除
        prefixes = ['部署', 'お', '■', '□', '●', '※']
        suffixes = ['・事業所名', 'の内容', 'について', 'のお願い', '：', ':', '欄']
        
        for prefix in prefixes:
            if name.startswith(prefix):
                name = name[len(prefix):]
        
        for suffix in suffixes:
            if name.endswith(suffix):
                name = name[:-len(suffix)]
        
        # マッピング（同じ意味の異なる表現を統一）
        mapping = {
            '企画名': '企画',
            '締切日': '締切',
            '文字数': '文字制限',
            'メッセージ': '原稿',
            '本文': '原稿',
            '氏名': '名前',
            'フリガナ': '名前',
            '所属': '部署',
            '事業所': '部署'
        }
        
        name = name.strip()
        return mapping.get(name, name)

    def _extract_main_content(self, paragraphs: List[str]) -> str:
        """本文部分を抽出する（完全版）"""
        all_text_chunks = []
        current_chunk = []
        in_table = False
        table_content = []
        
        def process_chunk(chunk: List[str]) -> str:
            """テキストチャンクを処理"""
            text = ' '.join(chunk).strip()
            text = self._clean_text(text)
            
            # 無効なチャンクの判定
            if any([
                not text,
                len(text) < 20,  # 短すぎるテキスト
                re.match(r'^[\d\s]+$', text),  # 数字のみ
                re.match(r'^[-=＝]+$', text),  # 区切り線
                re.match(r'^((?:●|■|※|【|》).+?[:：]|.+?[:：].+?$)', text),  # メタデータ
            ]):
                return ""
            
            return text

        for line in paragraphs:
            line = self._clean_text(line)
            
            # テーブルの処理
            if '|' in line or '+-' in line:
                # 現在のチャンクを処理
                if current_chunk:
                    if processed := process_chunk(current_chunk):
                        all_text_chunks.append(processed)
                    current_chunk = []
                
                # テーブル行の処理
                if '|' in line:
                    cells = [cell.strip() for cell in line.split('|')]
                    for cell in cells:
                        if len(cell) >= 20:  # 長いセル内容は本文候補
                            table_content.append(cell)
                
                in_table = True
                continue
            
            # テーブル終了の検出
            if in_table and not ('|' in line or '+-' in line):
                in_table = False
                # テーブル内容の処理
                for content in table_content:
                    if processed := process_chunk([content]):
                        all_text_chunks.append(processed)
                table_content = []
            
            # 通常のテキスト処理
            if not in_table:
                if line:
                    current_chunk.append(line)
                else:
                    if current_chunk:
                        if processed := process_chunk(current_chunk):
                            all_text_chunks.append(processed)
                        current_chunk = []
        
        # 最後のチャンクを処理
        if current_chunk:
            if processed := process_chunk(current_chunk):
                all_text_chunks.append(processed)
        
        # 最適な本文を選択
        if not all_text_chunks:
            return ""
            
        # 最も長い本文を選択
        main_content = max(all_text_chunks, key=len)
        
        # 文字数制限のバリデーション
        if len(main_content) < 150:
            return f"{main_content}（要追記）"
        elif len(main_content) > 200:
            return f"{main_content[:200]}（文字数超過）"
        
        return main_content

    def extract_content_for_column(self, structure: Dict, column: str) -> str:
        """特定のカラムの内容を抽出"""
        content = []
        
        # 各ヘッダーレベルから内容を探索
        for level in ['h2', 'h3']:
            for item in structure[level]:
                if self._normalize_column_name(item['text']) == column:
                    content.extend(item['content'])
        
        # 原稿の場合は特別処理
        if column == '原稿' and not content:
            paragraphs = structure['text']
            content = [self._extract_main_content(paragraphs)]
        
        # 特定のカラム用の後処理
        result = ' '.join(content).strip()
        if column == '原稿':
            # 文字数チェックは_extract_main_contentで実施済み
            return result
        elif column in ['部署', '名前']:
            # 部署と名前は最初の有効な値のみを使用
            return result.split()[0] if result else ""
        
        return result

    def validate_data(self, data: Dict[str, str]) -> List[str]:
        """データのバリデーション"""
        errors = []
        for field, rule in self.validation_rules.items():
            if field in data:
                if not rule(data[field]):
                    if field == '原稿':
                        word_count = len(data[field])
                        errors.append(
                            f"{field}の文字数が不適切です（現在: {word_count}文字, 要件: 150-200文字）"
                        )
                    else:
                        errors.append(f"{field}が未入力です")
        return errors

    def process_document(self, doc: Dict, index: int) -> Tuple[Dict, List[ProcessingError]]:
        """単一文書の処理"""
        try:
            md_content = self.processor.convert_to_markdown(doc['document_content'])
            structure = self.processor.extract_structure(md_content)
            
            data = {'番号': index + 1}
            
            # 特別なカラムの処理
            if '企画' in self.selected_columns:
                data['企画'] = self.extract_content_for_column(structure, '企画名') or \
                            self.extract_content_for_column(structure, '企画')

            # その他のカラム処理
            for column in self.selected_columns:
                if column not in data:  # まだ処理されていないカラムのみ
                    content = self.extract_content_for_column(structure, column)
                    data[column] = content

            # バリデーション
            validation_errors = self.validate_data(data)
            if validation_errors:
                for error in validation_errors:
                    self.processor.errors.append(
                        ProcessingError(index, "検証エラー", error)
                    )

            return data, self.processor.errors

        except Exception as e:
            self.processor.errors.append(
                ProcessingError(index, "処理エラー", str(e))
            )
            return None, self.processor.errors

    def process_documents(self, documents: List[Dict]) -> Tuple[pd.DataFrame, List[ProcessingError]]:
        """複数文書の処理"""
        processed_docs = []
        self.processor.errors = []  # エラーリストのリセット

        for idx, doc in enumerate(documents):
            result, errors = self.process_document(doc, idx)
            if result:
                processed_docs.append(result)

        if not processed_docs:
            return pd.DataFrame(), self.processor.errors

        df = pd.DataFrame(processed_docs)
        return df, self.processor.errors
