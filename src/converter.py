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

    # その他のメソッドは前回のコードと同じ
