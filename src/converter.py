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
        """DOCXコンテンツから潜在的なカラム候補を動的に抽出する"""
        columns = set(['番号'])  # デフォルトカラム
        
        # ヘッダー形式のパターン
        patterns = [
            r'【(.+?)】',  # 【】で囲まれた項目
            r'^(.+?)[:：]',  # 行頭から：までの項目
            r'■\s*(.+?)\s*[:：]',  # ■マークで始まる項目
            r'□\s*(.+?)\s*[:：]',  # □マークで始まる項目
            r'●\s*(.+?)\s*[:：]',  # ●マークで始まる項目
        ]
        
        # デフォルトで含めたい列
        common_columns = {'所属', '名前', '原稿', '勤続年数', '入社年', '部署'}
        columns.update(common_columns)
        
        # パターンマッチングによる列名抽出
        for pattern in patterns:
            matches = re.finditer(pattern, content, re.MULTILINE)
            for match in matches:
                column_name = match.group(1).strip()
                if column_name:
                    columns.add(self._normalize_column_name(column_name))
        
        return columns

    def _normalize_column_name(self, name: str) -> str:
        """カラム名を正規化"""
        name = name.strip()
        
        # 特定のプレフィックスとサフィックスを削除
        prefixes = ['部署', 'お']
        suffixes = ['・事業所名', 'の内容', 'について']
        
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
            '文字数': '文字制限'
        }
        
        return mapping.get(name.strip(), name.strip())

    def _extract_main_content(self, paragraphs: List[str]) -> str:
        """本文部分を抽出する（汎用版）"""
        content_lines = []
        is_content = False
        current_line = []
        
        # メタデータパターン（このパターンに当てはまらない部分が本文の候補）
        metadata_patterns = [
            r'^[\s|]*[-=＝]+[\s|]*$',  # 区切り線
            r'^.*[:：].*$',            # キー：値形式
            r'^\s*\|.*\|\s*$',        # 表の行
            r'^\s*\d+\s*$',           # 数字のみの行
            r'^\s*$',                 # 空行
            r'【.*】'                 # 【】で囲まれた見出し
        ]
        
        # 最初のメタデータ以外のまとまったテキストを本文とみなす
        consecutive_text_lines = 0
        
        for line in paragraphs:
            line = line.strip()
            
            # メタデータパターンに一致するかチェック
            is_metadata = any(re.match(pattern, line) for pattern in metadata_patterns)
            
            if not is_metadata and line:
                # パイプと数字を除去
                cleaned_line = re.sub(r'\|\s*\d+\s*$', '', line)  # 右端の数字を除去
                cleaned_line = re.sub(r'^\|\s*', '', cleaned_line)  # 左端のパイプを除去
                cleaned_line = cleaned_line.strip()
                
                if cleaned_line:
                    consecutive_text_lines += 1
                    current_line.append(cleaned_line)
            else:
                if consecutive_text_lines >= 2:  # 複数行の通常テキストが続いた場合、本文とみなす
                    is_content = True
                
                if is_content and current_line:
                    content_lines.append(' '.join(current_line))
                consecutive_text_lines = 0
                current_line = []
        
        # 最後の行が残っている場合は追加
        if current_line and (is_content or consecutive_text_lines >= 2):
            content_lines.append(' '.join(current_line))
        
        # 全ての行を結合
        content = ' '.join(content_lines).strip()
        
        # 内容が極端に短い場合は本文として不適切と判断
        if len(content.split()) < 3:  # 3単語未満は不適切と判断
            return ""
        
        return content

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
            for column in self.selected_columns:
                if column != '番号':
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
