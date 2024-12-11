# src/converter.py

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
        columns = set(['番号'])  # 番号は常に必要なので残す
        
        # 文書内の実際のセクションを探す
        patterns = [
            (r'【(.+?)】', 1),           # 【】で囲まれた項目
            (r'■\s*(.+?)\s*[:：]', 1),   # ■マークで始まる項目
            (r'□\s*(.+?)\s*[:：]', 1),   # □マークで始まる項目
            (r'●\s*(.+?)\s*[:：]', 1),   # ●マークで始まる項目
            (r'^(.+?)[:：](?!\d)', 1),   # 行頭から：までの項目（時刻を除外）
        ]
        
        for line in content.split('\n'):
            line = line.strip()
            
            # 各パターンでマッチを試行
            for pattern, group in patterns:
                match = re.search(pattern, line)
                if match:
                    column_name = match.group(group).strip()
                    if column_name:
                        normalized_name = self._normalize_column_name(column_name)
                        if normalized_name:  # 空文字列でない場合のみ追加
                            columns.add(normalized_name)
                    break  # マッチしたら次の行へ
        
        # カラムの正規化とフィルタリング
        normalized_columns = set()
        for column in columns:
            # 明らかにカラムではないものを除外
            if not any(skip in column.lower() for skip in [
                'について', 'お願い', 'です', 'ます', 'した', 'から'
            ]):
                normalized_columns.add(column)
        
        # 原稿セクションの特別処理
        if any(keyword in content for keyword in ['原稿', 'メッセージ', '本文']):
            normalized_columns.add('原稿')
        
        return normalized_columns

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
        # 特定のセクションを示すキーワード
        section_keywords = ['企画', '締切', '文字数', '原稿', 'メッセージ']
        
        content = []
        current_section = None
        
        for line in paragraphs:
            line = line.strip()
            
            # セクション開始の検出
            section_match = re.search(r'【(.+?)】|■\s*(.+?)\s*[:：]', line)
            if section_match:
                section_name = section_match.group(1) or section_match.group(2)
                section_name = self._normalize_column_name(section_name)
                if section_name in section_keywords:
                    current_section = section_name
                    # セクション名の部分を除去
                    line = re.sub(r'【.+?】|■\s*.+?\s*[:：]', '', line).strip()
            
            # 区切り線をスキップ
            if re.match(r'^[-=＝]+$', line):
                continue
                
            # メタデータ行をスキップ
            if re.match(r'^\s*\|.*\|\s*$', line) or not line:
                continue
                
            # 本文として追加
            if current_section in ['原稿', 'メッセージ'] and line:
                content.append(line)
        
        # 内容を結合して返す
        result = ' '.join(content).strip()
        
        # 文字数制限のバリデーション（150〜200文字）
        if len(result) < 150:
            return f"{result}（要追記）"
        elif len(result) > 200:
            return f"{result[:200]}（文字数超過）"
        
        return result

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
