"""API 配置管理"""
import os
import json

# 项目根目录
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class Settings:
    """API 配置，支持环境变量覆盖"""

    def __init__(self):
        # API 认证
        self.api_key: str = os.environ.get("SCRAPER_API_KEY", "changeme-your-secret-key")

        # Chrome 配置
        self.chrome_debug_port: int = int(os.environ.get("SCRAPER_CHROME_PORT", "9222"))
        self.chrome_fallback_ports: list[int] = [9222, 9223, 9224, 9225]

        # 任务配置
        self.task_history_max: int = 100
        self.task_result_ttl: int = 3600

        # 文件输出目录
        self.output_dir: str = BASE_DIR

        # 通知默认配置（从 scraper_config.json 加载）
        self.wx_type: str = "server_chan"
        self.wx_key: str = ""
        self.email_from: str = ""
        self.email_pwd: str = ""
        self.email_to: str = ""
        self.own_shop: str = ""

        # 尝试从 scraper_config.json 加载默认值
        self._load_scraper_config()

    def _load_scraper_config(self):
        """从现有配置文件加载默认值"""
        cfg_path = os.path.join(BASE_DIR, "scraper_config.json")
        if not os.path.exists(cfg_path):
            return
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.wx_key = cfg.get("wx_key", self.wx_key)
            self.wx_type = cfg.get("wx_type", self.wx_type)
            self.email_from = cfg.get("email_from", self.email_from)
            self.email_pwd = cfg.get("email_pwd", self.email_pwd)
            self.email_to = cfg.get("email_to", self.email_to)
            self.own_shop = cfg.get("own_shop", self.own_shop)
        except Exception:
            pass


settings = Settings()
