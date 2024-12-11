import re
import pandas as pd
from typing import Dict, List, Set, Tuple, Optional
from datetime import datetime
import markdown
from bs4 import BeautifulSoup
from dataclasses import dataclass
import logging
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
        """エラー情報を辞書形式で返す"""
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
    def normalize_spaces(text: str) -> str:
        """空白の正規化"""
        return ' '.join(text.split())
    
    @staticmethod
    def normalize_japanese_numbers(text: str) -> str:
        """日本語数字の正規化"""
        mapping = str.maketrans('０１２３４５６７８９', '0123456789')
        return text.translate(mapping)
    
    @staticmethod
    def clean_text(text: str) -> str:
        """テキストの総合的なクリーニング"""
        if not text:
            return ""
        # 改行とタブを空白に変換
        text = re.sub(r'[\n\t\r]', ' ', text)
        # 複数の空白を1つに
        text = re.sub(r'\s+', ' ', text)
        # 全角数字を半角に
        text = TextNormalizer.normalize_japanese_numbers(text)
        return text.strip()

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
            # 文書のクリーニング
            content = TextNormalizer.clean_text(content)
            
            # セクション見出しの変換
            content = re.sub(r'【(.+?)】', r'### \1', content)
            
            lines = []
            in_table = False
            table_content = []
            
            for line in content.split('\n'):
                line = line.strip()
                
                # テーブル処理
                if '|' in line or line.startswith('+-'):
                    if not in_table:
                        if table_content:
                            lines.extend(table_content)
                            table_content = []
                        in_table = True
                    table_content.append(line)
                    continue
                
                # テーブル終了の検出
                if in_table and not ('|' in line or line.startswith('+-')):
                    in_table = False
                    if table_content:
                        # テーブル内容を処理してから追加
                        processed_table = self._process_table_content(table_content)
                        lines.extend(processed_table)
                        table_content = []
                
                # 通常のテキスト処理
                if not in_table:
                    if '：' in line:
                        key, value = line.split('：', 1)
                        lines.append(f'### {key.strip()}\n{value.strip()}')
                    elif line:
                        lines.append(line)
            
            # 最後のテーブルの処理
            if table_content:
                processed_table = self._process_table_content(table_content)
                lines.extend(processed_table)
            
            return '\n\n'.join(lines)
        except Exception as e:
            self.logger.error(f"Markdown変換エラー: {str(e)}")
            raise ValueError(f"Markdown変換エラー: {str(e)}")

    def _process_table_content(self, table_content: List[str]) -> List[str]:
        """テーブル内容の処理"""
        processed_lines = []
        for line in table_content:
            if line.startswith('+-'):
                continue
            
            # セル内容の抽出と処理
            cells = [cell.strip() for cell in line.split('|')]
            cells = [cell for cell in cells if cell]  # 空のセルを除去
            
            for cell in cells:
                if len(cell) >= 20:  # 長いテキストは本文候補
                    processed_lines.append(cell)
                elif '：' in cell:  # メタデータは見出しとして処理
                    key, value = cell.split('：', 1)
                    processed_lines.append(f'### {key.strip()}\n{value.strip()}')
        
        return processed_lines

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
                    text = TextNormalizer.clean_text(element.get_text())
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
        self.text_normalizer = TextNormalizer()
        self.selected_columns: Optional[List[str]] = None
        self.logger = self.processor.logger
        
        # バリデーションルール
        self.validation_rules = {
            '原稿': lambda x: 150 <= len(x) <= 200 if x else False,
            '名前': lambda x: bool(x and x.strip()),
            '部署': lambda x: bool(x and x.strip()),
            '企画': lambda x: bool(x and x.strip()),
        }

    def extract_potential_columns(self, content: str) -> Set[str]:
        """DOCXコンテンツから実際に存在するカラムのみを抽出"""
        columns = set(['番号'])  # 基本カラム
        
        # テーブル内のカラム抽出
        table_pattern = r'\|(.*?)\|'
        table_rows = re.findall(table_pattern, content, re.MULTILINE)
        
        for row in table_rows:
            cells = [cell.strip() for cell in row.split('|')]
            for cell in cells:
                if '：' in cell or ':' in cell:
                    key = cell.split('：')[0].split(':')[0].strip()
                    if key:
                        normalized = self._normalize_column_name(key)
                        if normalized:
                            columns.add(normalized)

        # 見出しからのカラム抽出
        header_patterns = [
            (r'【(.+?)】', 1),
            (r'■\s*(.+?)\s*[:：]', 1),
            (r'□\s*(.+?)\s*[:：]', 1),
            (r'●\s*(.+?)\s*[:：]', 1),
            (r'^(?:■|□|●)?\s*(.+?)\s*[:：](?!\d)', 1),
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
            text = self.text_normalizer.clean_text(text)
            
            # 無効なチャンクの判定
            if any([
                not text,
                len(text) < 20,
                re.match(r'^[\d\s]+$', text),
                re.match(r'^[-=＝]+$', text),
                re.match(r'^((?:●|■|※|【|》).+?[:：]|.+?[:：].+?$)', text),
            ]):
                return ""
            
            return text

        for line in paragraphs:
            line = self.text_normalizer.clean_text(line)
            
            # テーブル処理
            if '|' in line or '+-' in line:
                if current_chunk:
                    if processed := process_chunk(current_chunk):
                        all_text_chunks.append(processed)
                    current_chunk = []
                
                if '|' in line:
                    cells = [cell.strip() for cell in line.split('|')]
                    for cell in cells:
                        if len(cell) >= 20:
                            table_content.append(cell)
                
                in_table = True
                continue
            
            # テーブル終了の検出
            if in_table and not ('|' in line or '+-' in line):
                in_table = False
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
            return result
        elif column in ['部署', '名前']:
            return
