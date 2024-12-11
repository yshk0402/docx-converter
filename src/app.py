import json
import os
from pathlib import Path
from typing import List, Dict, Optional, Any
import logging
from datetime import datetime
import shutil

class Config:
    def __init__(self):
        """設定管理クラスの初期化"""
        self.config_dir = Path.home() / '.docx_converter'
        self.config_file = self.config_dir / 'config.json'
        self.log_dir = self.config_dir / 'logs'
        self.backup_dir = self.config_dir / 'backups'
        
        # ロギングの設定
        self._setup_logging()
        
        # 設定ディレクトリの初期化
        self.ensure_config_dir()

    def _setup_logging(self):
        """ロギングの設定"""
        try:
            self.log_dir.mkdir(parents=True, exist_ok=True)
            log_file = self.log_dir / 'config.log'
            
            logging.basicConfig(
                filename=log_file,
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            self.logger = logging.getLogger(__name__)
        except Exception as e:
            print(f"ロギングの設定に失敗しました: {str(e)}")
            raise

    def ensure_config_dir(self):
        """設定ディレクトリの作成と初期化"""
        try:
            # 必要なディレクトリを作成
            self.config_dir.mkdir(parents=True, exist_ok=True)
            self.backup_dir.mkdir(parents=True, exist_ok=True)
            self.log_dir.mkdir(parents=True, exist_ok=True)
            
            # 設定ファイルが存在しない場合は作成
            if not self.config_file.exists():
                self.save_config(self._get_default_config())
                self.logger.info("デフォルト設定ファイルを作成しました")
        except Exception as e:
            self.logger.error(f"設定ディレクトリの初期化中にエラーが発生: {str(e)}")
            raise

    def _get_default_config(self) -> Dict[str, Any]:
        """デフォルト設定の取得"""
        return {
            'version': '1.0.0',
            'last_modified': datetime.now().isoformat(),
            'favorite_columns': ['番号', '原稿'],
            'settings': {
                'max_preview_length': 500,
                'default_columns': ['番号', '原稿'],
                'excel_settings': {
                    'default_sheet_name': 'データ',
                    'error_highlight_color': 'FFE7E6',
                    'min_column_width': 10,
                    'max_column_width': 50
                },
                'validation': {
                    'min_text_length': 150,
                    'max_text_length': 200
                }
            }
        }

    def _create_backup(self):
        """設定ファイルのバックアップを作成"""
        if self.config_file.exists():
            try:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_file = self.backup_dir / f'config_{timestamp}.json'
                
                # バックアップを作成
                shutil.copy2(self.config_file, backup_file)
                self.logger.info(f"設定ファイルのバックアップを作成: {backup_file}")
                
                # 古いバックアップの削除（最新の5つのみ保持）
                self._cleanup_old_backups()
            except Exception as e:
                self.logger.error(f"バックアップ作成中にエラーが発生: {str(e)}")

    def _cleanup_old_backups(self, keep_count: int = 5):
        """古いバックアップファイルの削除"""
        try:
            backup_files = sorted(self.backup_dir.glob('config_*.json'))
            if len(backup_files) > keep_count:
                for old_file in backup_files[:-keep_count]:
                    old_file.unlink()
                    self.logger.info(f"古いバックアップファイルを削除: {old_file}")
        except Exception as e:
            self.logger.error(f"バックアップクリーンアップ中にエラーが発生: {str(e)}")

    def _handle_corrupt_config(self):
        """破損した設定ファイルの処理"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            corrupt_file = self.config_dir / f'corrupt_config_{timestamp}.json'
            
            if self.config_file.exists():
                shutil.move(self.config_file, corrupt_file)
                self.logger.warning(f"破損した設定ファイルを移動: {corrupt_file}")
        except Exception as e:
            self.logger.error(f"破損ファイルの処理中にエラーが発生: {str(e)}")

    def load_config(self) -> Dict[str, Any]:
        """設定の読み込み"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.logger.debug("設定ファイルを読み込みました")
                return config
            else:
                self.logger.warning("設定ファイルが存在しないためデフォルト設定を使用します")
                return self._get_default_config()
        except json.JSONDecodeError as e:
            self.logger.error(f"設定ファイルの解析に失敗: {str(e)}")
            self._handle_corrupt_config()
            return self._get_default_config()
        except Exception as e:
            self.logger.error(f"設定ファイルの読み込み中にエラーが発生: {str(e)}")
            return self._get_default_config()

    def save_config(self, config: Dict[str, Any]):
        """設定の保存"""
        try:
            # バックアップの作成
            self._create_backup()
            
            # 設定の更新
            config['last_modified'] = datetime.now().isoformat()
            
            # 設定の保存
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            self.logger.info("設定ファイルを保存しました")
        except Exception as e:
            self.logger.error(f"設定ファイルの保存中にエラーが発生: {str(e)}")
            raise

    def update_favorite_columns(self, columns: List[str]):
        """よく使うカラムの更新"""
        try:
            config = self.load_config()
            
            # 重複を除去し、順序を保持
            columns = list(dict.fromkeys(columns))
            
            config['favorite_columns'] = columns
            self.save_config(config)
            
            self.logger.info(f"お気に入りカラムを更新しました: {columns}")
        except Exception as e:
            self.logger.error(f"お気に入りカラムの更新中にエラーが発生: {str(e)}")
            raise

    def get_setting(self, key: str, default: Any = None) -> Any:
        """特定の設定値の取得"""
        try:
            config = self.load_config()
            return config.get('settings', {}).get(key, default)
        except Exception as e:
            self.logger.error(f"設定値の取得中にエラーが発生: {str(e)}")
            return default

    def update_setting(self, key: str, value: Any):
        """特定の設定値の更新"""
        try:
            config = self.load_config()
            if 'settings' not in config:
                config['settings'] = {}
            config['settings'][key] = value
            self.save_config(config)
            self.logger.info(f"設定を更新しました: {key}={value}")
        except Exception as e:
            self.logger.error(f"設定の更新中にエラーが発生: {str(e)}")
            raise

    def reset_config(self):
        """設定を初期状態に戻す"""
        try:
            self._create_backup()
            self.save_config(self._get_default_config())
            self.logger.info("設定を初期化しました")
        except Exception as e:
            self.logger.error(f"設定の初期化中にエラーが発生: {str(e)}")
            raise
