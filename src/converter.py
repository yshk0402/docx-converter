import re
import pandas as pd
import logging
from typing import Dict, List, Set, Tuple, Optional
from datetime import datetime
import markdown
from bs4 import BeautifulSoup
from dataclasses import dataclass
from pathlib import Path

@dataclass
class ProcessingError:
    """処理エラー情報を格納するデータクラス"""
    document_index: int
    error_type: str
    message: str
    details: str = ""
    timestamp: datetime = datetime.now()

    def to_dict(self) -> Dict:
        return {
            'index': self.document_index,
            'type': self.error_type,
            'message': self.message,
            'details': self.details,
            'timestamp': self.timestamp.isoformat()
        }

class TextNormalizer:
    """テキスト正規化を行うユーティリティクラス"""
    @staticmethod
    def clean_text(text: str) -> str:
        if not text:
            return ""
        # 空白の正規化
        text = re.sub(r'\s+', ' ', text.strip())
        # 特殊文字を削除
        text = re.sub(r'[\r\n\t\f\v]', '', text)
        return text

class DocumentProcessor:
    """文書処理を行うクラス"""
    def __init__(self):
        self.errors: List[ProcessingError] = []
        self._setup_logging()

    def _setup_logging(self):
        """ロギングの設定"""
        log_dir = Path.home() / '.docx_converter' / 'logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / 'document_processor.log'
        
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        self.logger = logging.getLogger(__name__)

    def convert_to_markdown(self, content: str) -> str:
        """文書をMarkdown形式に変換"""
        try:
            content = TextNormalizer.clean_text(content)
            content = re.sub(r'【(.+?)】', r'### \1', content)
            
            lines = []
            in_table = False
            table_content = []
            
            for line in content.split('\n'):
                line = line.strip()
                
                if '|' in line or line.startswith('+-'):
                    if not in_table:
                        in_table = True
                    table_content.append(line)
                    continue
                
                if in_table and not ('|' in line or line.startswith('+-')):
                    in_table = False
                    processed_lines = []
                    for table_line in table_content:
                        if not table_line.startswith('+-'):
                            cells = [cell.strip() for cell in table_line.split('|')]
                            for cell in cells:
                                if cell and not cell.isspace():
                                    processed_lines.append(cell)
                    lines.extend(processed_lines)
                    table_content = []
                    continue
                
                if not in_table and line:
                    if '：' in line:
                        key, value = line.split('：', 1)
                        lines.append(f'### {key.strip()}\n{value.strip()}')
                    else:
                        lines.append(line)
            
            # 最後のテーブルの処理
            if table_content:
                processed_lines = []
                for table_line in table_content:
                    if not table_line.startswith('+-'):
                        cells = [cell.strip() for cell in table_line.split('|')]
                        for cell in cells:
                            if cell and not cell.isspace():
                                processed_lines.append(cell)
                lines.extend(processed_lines)
            
            return '\n\n'.join(lines)
        
        except Exception as e:
            self.logger.error(f"Markdown変換エラー: {str(e)}")
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
            self.logger.error(f"構造抽出エラー: {str(e)}")
            raise ValueError(f"構造抽出エラー: {str(e)}")
class MessageConverter:
    """メッセージ変換の主要クラス"""
    def __init__(self):
        self.processor = DocumentProcessor()
        self.selected_columns: Optional[List[str]] = None
        self.logger = self.processor.logger
        
        self.validation_rules = {
            '原稿': lambda x: 150 <= len(x) <= 200 if x else False,
            '名前': lambda x: bool(x and x.strip()),
            '部署': lambda x: bool(x and x.strip()),
            '企画': lambda x: bool(x and x.strip()),
        }

    def _normalize_column_name(self, name: str) -> str:
        """カラム名を正規化"""
        name = name.strip()
        
        # プレフィックスとサフィックスを削除
        prefixes = ['部署', 'お', '■', '□', '●', '※']
        suffixes = ['・事業所名', 'の内容', 'について', 'のお願い', '：', ':', '欄']
        
        for prefix in prefixes:
            if name.startswith(prefix):
                name = name[len(prefix):]
        
        for suffix in suffixes:
            if name.endswith(suffix):
                name = name[:-len(suffix)]
        
        # マッピング
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

    def extract_potential_columns(self, content: str) -> Set[str]:
        """実際に存在するカラムのみを抽出"""
        columns = set(['番号'])
        
        # テーブル内のカラム抽出
        table_headers = re.findall(r'\|\s*([^|]+?)\s*(?:：|\:)', content)
        for header in table_headers:
            normalized = self._normalize_column_name(header)
            if normalized:
                columns.add(normalized)
        
        # 見出しからのカラム抽出
        header_patterns = [
            (r'【(.+?)】', 1),
            (r'■\s*(.+?)\s*[:：]', 1),
            (r'□\s*(.+?)\s*[:：]', 1),
            (r'●\s*(.+?)\s*[:：]', 1),
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
        if self._extract_main_content(paragraphs).strip():
            columns.add('原稿')
        
        # 不要なカラムの除外
        columns = {col for col in columns if not any(skip in col.lower() for skip in [
            'について', 'お願い', 'です', 'ます', 'した', 'から', '注意', '備考'
        ])}
        
        return columns

    def _extract_main_content(self, paragraphs: List[str]) -> str:
        """本文部分を抽出する"""
        all_text_chunks = []
        current_chunk = []
        in_table = False
        
        def process_chunk(text: str) -> Optional[str]:
            text = TextNormalizer.clean_text(text)
            if len(text) >= 20 and not re.match(r'^[\d\s]+$', text):
                return text
            return None
        
        for line in paragraphs:
            line = line.strip()
            
            # テーブル処理
            if '|' in line or line.startswith('+-'):
                if current_chunk:
                    if chunk_text := process_chunk(' '.join(current_chunk)):
                        all_text_chunks.append(chunk_text)
                    current_chunk = []
                
                if '|' in line:
                    cells = [cell.strip() for cell in line.split('|')]
                    for cell in cells:
                        if len(cell) >= 20:
                            if chunk_text := process_chunk(cell):
                                all_text_chunks.append(chunk_text)
                
                in_table = True
                continue
            
            # テーブル終了の検出
            if in_table and not ('|' in line or line.startswith('+-')):
                in_table = False
            
            # 通常のテキスト処理
            if not in_table and line:
                if not re.match(r'^[-=＝]+$', line):  # 区切り線を除外
                    current_chunk.append(line)
            elif current_chunk:
                if chunk_text := process_chunk(' '.join(current_chunk)):
                    all_text_chunks.append(chunk_text)
                current_chunk = []
        
        # 最後のチャンクを処理
        if current_chunk:
            if chunk_text := process_chunk(' '.join(current_chunk)):
                all_text_chunks.append(chunk_text)
        
        # 最適な本文を選択
        if not all_text_chunks:
            return ""
        
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
        
        # 結果の処理
        result = ' '.join(content).strip()
        if column == '原稿':
            return result
        elif column in ['部署', '名前']:
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
            if '企画' in self.selected_columns:
                data['企画'] = self.extract_content_for_column(structure, '企画名') or \
                            self.extract_content_for_column(structure, '企画')

            for column in self.selected_columns:
                if column not in data:
                    content = self.extract_content_for_column(structure, column)
                    data[column] = content

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
            try:
                result, errors = self.process_document(doc, idx)
                if result:
                    processed_docs.append(result)
            except Exception as e:
                self.processor.errors.append(
                    ProcessingError(
                        document_index=idx,
                        error_type="処理エラー",
                        message=str(e)
                    )
                )

        if not processed_docs:
            return pd.DataFrame(), self.processor.errors

        df = pd.DataFrame(processed_docs)
        return df, self.processor.errors
