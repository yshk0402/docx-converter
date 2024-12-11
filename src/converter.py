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
            content = re.sub(r'\*\*【(.+?)】\*\*', r'## \1', content)
            content = re.sub(r'【(.+?)】', r'### \1', content)
            
            lines = []
            for line in content.split('\n'):
                line = line.strip()
                if line.startswith('+-'):
                    continue
                if '：' in line:
                    key, value = line.split('：', 1)
                    lines.append(f'### {key.strip()}\n{value.strip()}')
                elif line and not line.startswith('|'):
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
                elif element.name == 'p' and current_section:
                    current_section['content'].append(element.get_text().strip())
                elif element.name == 'p':
                    structure['text'].append(element.get_text().strip())
            
            return structure
        except Exception as e:
            raise ValueError(f"構造抽出エラー: {str(e)}")

class MessageConverter:
    def __init__(self):
        self.default_columns = ['番号']
        self.selected_columns = None
        self.processor = DocumentProcessor()
        self.validation_rules = {
            '原稿': lambda x: len(x) <= 200 if x else True,
            '名前': lambda x: bool(x.strip()) if x else False
        }

    def extract_potential_columns(self, content: str) -> Set[str]:
        """DOCXコンテンツから実際に存在するカラムのみを抽出する"""
        columns = set(['番号'])
        
        # テーブルヘッダーからカラムを抽出
        table_headers = re.findall(r'\|\s*([^|]+?)\s*(?:：|\:)', content)
        for header in table_headers:
            normalized = self._normalize_column_name(header)
            if normalized:
                columns.add(normalized)
        
        # その他のパターンからカラムを抽出
        patterns = [
            (r'【(.+?)】', 1),
            (r'■\s*(.+?)\s*[:：]', 1),
            (r'□\s*(.+?)\s*[:：]', 1),
            (r'●\s*(.+?)\s*[:：]', 1),
        ]
        
        for line in content.split('\n'):
            line = line.strip()
            for pattern, group in patterns:
                match = re.search(pattern, line)
                if match:
                    column_name = match.group(group).strip()
                    if column_name:
                        normalized = self._normalize_column_name(column_name)
                        if normalized:
                            columns.add(normalized)
        
        # 20文字以上のテキストブロックが存在する場合は「原稿」カラムを追加
        paragraphs = content.split('\n')
        if self._extract_main_content(paragraphs).strip():
            columns.add('原稿')
        
        return columns

    def _normalize_column_name(self, name: str) -> str:
        """カラム名を正規化"""
        name = name.strip()
        
        # 特定のプレフィックスとサフィックスを削除
        prefixes = ['部署', 'お', '■', '□', '●', '※']
        suffixes = ['・事業所名', 'の内容', 'について', 'のお願い', '：', ':']
        
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
            '本文': '原稿'
        }
        
        name = name.strip()
        return mapping.get(name, name)

    def _extract_main_content(self, paragraphs: List[str]) -> str:
        """本文部分を抽出する（改善版）"""
        all_text_chunks = []
        current_chunk = []
        in_table = False
        
        def process_chunk(chunk):
            """テキストチャンクを処理して有効な本文かどうかを判断"""
            text = ' '.join(chunk).strip()
            # 数字のみの行を除外
            if re.match(r'^\d+$', text):
                return None
            # メタデータっぽい行を除外
            if re.match(r'^((?:●|■|※|【|》).+|.+[:：].+)$', text):
                return None
            # 区切り線を除外
            if re.match(r'^[-=＝]+$', text):
                return None
            # 最低文字数を満たさないものを除外
            if len(text) < 20:
                return None
            return text

        for line in paragraphs:
            line = line.strip()
            
            # テーブル処理
            if '+-' in line or '|' in line:
                if current_chunk:
                    if processed := process_chunk(current_chunk):
                        all_text_chunks.append(processed)
                    current_chunk = []
                
                in_table = True
                # テーブルの行から本文を抽出
                parts = [part.strip() for part in line.split('|')]
                text_parts = [part for part in parts if part and not part.startswith('+')]
                if text_parts:
                    for part in text_parts:
                        if len(part.strip()) >= 20:  # テーブル内の長いテキストは本文候補
                            all_text_chunks.append(part.strip())
                continue
                
            # テーブル終了判定
            if in_table and not ('+-' in line or '|' in line):
                in_table = False
            
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
        
        # 最も長いテキストチャンクを本文として採用
        if all_text_chunks:
            main_content = max(all_text_chunks, key=len)
            
            # 文字数制限のバリデーション（150〜200文字）
            if len(main_content) < 150:
                return f"{main_content}（要追記）"
            elif len(main_content) > 200:
                return f"{main_content[:200]}（文字数超過）"
            
            return main_content
        
        return ""

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
        
        return ' '.join(content).strip()

    def validate_data(self, data: Dict[str, str]) -> List[str]:
        """データのバリデーション"""
        errors = []
        for field, rule in self.validation_rules.items():
            if field in data:
                if not rule(data[field]):
                    if field == '原稿':
                        errors.append(f"{field}が200文字を超えています")
                    else:
                        errors.append(f"{field}が無効です")
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
