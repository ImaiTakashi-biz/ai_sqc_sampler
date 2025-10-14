"""
メモリ管理モジュール
アプリケーションのメモリ使用量を最適化
"""

import gc
import threading
import time
import psutil
import os
from typing import Dict, Any, Optional
from collections import deque
import logging


class MemoryManager:
    """メモリ管理クラス"""
    
    def __init__(self, max_cache_size: int = 100, gc_threshold: int = 50):
        self.max_cache_size = max_cache_size
        self.gc_threshold = gc_threshold
        self._cache = {}
        self._cache_access_times = {}
        self._lock = threading.Lock()
        self._memory_usage_history = deque(maxlen=100)
        self._last_gc_time = 0
        self._gc_interval = 300  # 5分間隔
        self.logger = self._setup_logger()
        
    def _setup_logger(self):
        """メモリログの設定"""
        logger = logging.getLogger('memory_manager')
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.FileHandler('memory.log', encoding='utf-8')
            formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s - %(message)s'
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        
        return logger
    
    def get_memory_usage(self) -> Dict[str, float]:
        """現在のメモリ使用量を取得"""
        try:
            process = psutil.Process(os.getpid())
            memory_info = process.memory_info()
            
            return {
                'rss': memory_info.rss / 1024 / 1024,  # MB
                'vms': memory_info.vms / 1024 / 1024,  # MB
                'percent': process.memory_percent(),
                'available': psutil.virtual_memory().available / 1024 / 1024  # MB
            }
        except Exception as e:
            self.logger.error(f"メモリ使用量取得エラー: {e}")
            return {}
    
    def cache_data(self, key: str, data: Any, max_age: int = 3600) -> bool:
        """データをキャッシュに保存"""
        try:
            with self._lock:
                # キャッシュサイズの制限
                if len(self._cache) >= self.max_cache_size:
                    self._cleanup_old_cache()
                
                # データの保存
                self._cache[key] = {
                    'data': data,
                    'timestamp': time.time(),
                    'max_age': max_age
                }
                self._cache_access_times[key] = time.time()
                
                self.logger.debug(f"キャッシュに保存: {key}")
                return True
                
        except Exception as e:
            self.logger.error(f"キャッシュ保存エラー: {e}")
            return False
    
    def get_cached_data(self, key: str) -> Optional[Any]:
        """キャッシュからデータを取得"""
        try:
            with self._lock:
                if key not in self._cache:
                    return None
                
                cache_item = self._cache[key]
                current_time = time.time()
                
                # 有効期限のチェック
                if current_time - cache_item['timestamp'] > cache_item['max_age']:
                    del self._cache[key]
                    if key in self._cache_access_times:
                        del self._cache_access_times[key]
                    return None
                
                # アクセス時間の更新
                self._cache_access_times[key] = current_time
                
                self.logger.debug(f"キャッシュから取得: {key}")
                return cache_item['data']
                
        except Exception as e:
            self.logger.error(f"キャッシュ取得エラー: {e}")
            return None
    
    def _cleanup_old_cache(self):
        """古いキャッシュのクリーンアップ"""
        try:
            current_time = time.time()
            
            # 有効期限切れのキャッシュを削除
            expired_keys = []
            for key, cache_item in self._cache.items():
                if current_time - cache_item['timestamp'] > cache_item['max_age']:
                    expired_keys.append(key)
            
            for key in expired_keys:
                del self._cache[key]
                if key in self._cache_access_times:
                    del self._cache_access_times[key]
            
            # まだキャッシュが満杯の場合は、最も古いアクセス時間のものを削除
            if len(self._cache) >= self.max_cache_size:
                sorted_keys = sorted(
                    self._cache_access_times.items(),
                    key=lambda x: x[1]
                )
                
                keys_to_remove = sorted_keys[:len(self._cache) - self.max_cache_size + 1]
                for key, _ in keys_to_remove:
                    if key in self._cache:
                        del self._cache[key]
                    if key in self._cache_access_times:
                        del self._cache_access_times[key]
            
            self.logger.info(f"キャッシュクリーンアップ完了: {len(expired_keys)}件削除")
            
        except Exception as e:
            self.logger.error(f"キャッシュクリーンアップエラー: {e}")
    
    def clear_cache(self):
        """キャッシュをクリア"""
        try:
            with self._lock:
                self._cache.clear()
                self._cache_access_times.clear()
                self.logger.info("キャッシュをクリアしました")
        except Exception as e:
            self.logger.error(f"キャッシュクリアエラー: {e}")
    
    def optimize_memory(self):
        """メモリの最適化"""
        try:
            current_time = time.time()
            
            # 定期的なガベージコレクション
            if current_time - self._last_gc_time > self._gc_interval:
                self._perform_garbage_collection()
                self._last_gc_time = current_time
            
            # メモリ使用量の記録
            memory_usage = self.get_memory_usage()
            if memory_usage:
                self._memory_usage_history.append(memory_usage)
                
                # メモリ使用量が閾値を超えた場合の処理
                if memory_usage.get('percent', 0) > self.gc_threshold:
                    self._perform_aggressive_cleanup()
            
            self.logger.debug(f"メモリ最適化完了: {memory_usage}")
            
        except Exception as e:
            self.logger.error(f"メモリ最適化エラー: {e}")
    
    def _perform_garbage_collection(self):
        """ガベージコレクションの実行"""
        try:
            # 古いキャッシュのクリーンアップ
            self._cleanup_old_cache()
            
            # Pythonのガベージコレクション
            collected = gc.collect()
            
            self.logger.info(f"ガベージコレクション完了: {collected}オブジェクト回収")
            
        except Exception as e:
            self.logger.error(f"ガベージコレクションエラー: {e}")
    
    def _perform_aggressive_cleanup(self):
        """積極的なメモリクリーンアップ"""
        try:
            # キャッシュの大幅削減
            with self._lock:
                if len(self._cache) > self.max_cache_size // 2:
                    # 最も古いアクセス時間の半分を削除
                    sorted_keys = sorted(
                        self._cache_access_times.items(),
                        key=lambda x: x[1]
                    )
                    
                    keys_to_remove = sorted_keys[:len(self._cache) // 2]
                    for key, _ in keys_to_remove:
                        if key in self._cache:
                            del self._cache[key]
                        if key in self._cache_access_times:
                            del self._cache_access_times[key]
            
            # 強制的なガベージコレクション
            gc.collect()
            
            self.logger.info("積極的メモリクリーンアップ完了")
            
        except Exception as e:
            self.logger.error(f"積極的クリーンアップエラー: {e}")
    
    def get_memory_stats(self) -> Dict[str, Any]:
        """メモリ統計情報の取得"""
        try:
            memory_usage = self.get_memory_usage()
            
            return {
                'current_usage': memory_usage,
                'cache_size': len(self._cache),
                'max_cache_size': self.max_cache_size,
                'memory_history_count': len(self._memory_usage_history),
                'last_gc_time': self._last_gc_time
            }
            
        except Exception as e:
            self.logger.error(f"メモリ統計取得エラー: {e}")
            return {}


# グローバルメモリマネージャーインスタンス
memory_manager = MemoryManager()
