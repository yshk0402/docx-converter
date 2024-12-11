import json
from pathlib import Path
from typing import List, Dict

class Config:
    def __init__(self):
        self.config_dir = Path.home() / '.docx_converter'
        self.config_file = self.config_dir / 'config.json'
        self.ensure_config_dir()

    def ensure_config_dir(self):
        """設定ディレクトリの作成"""
        self.config_dir.mkdir(parents=True, exist_ok=True)
        if not self.config_file.exists():
            self.save_config({'favorite_columns': []})

    def load_config(self) -> Dict:
        """設定の読み込み"""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {'favorite_columns': []}

    def save_config(self, config: Dict):
        """設定の保存"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

    def update_favorite_columns(self, columns: List[str]):
        """よく使うカラムの更新"""
        config = self.load_config()
        config['favorite_columns'] = columns
        self.save_config(config)
