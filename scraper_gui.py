# -*- coding: utf-8 -*-
"""Taobao Scraper GUI v5.2 - Multi-language (ZH/EN/KO)"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading, time, re, os, logging, traceback, smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import queue
import json as _json

# ── SMTP 多邮箱配置 ──────────────────────────
SMTP_SERVERS = {
    "qq.com":      ("smtp.qq.com",      465, True),
    "foxmail.com": ("smtp.qq.com",      465, True),
    "163.com":     ("smtp.163.com",     465, True),
    "126.com":     ("smtp.126.com",     465, True),
    "yeah.net":    ("smtp.yeah.net",    465, True),
    "gmail.com":   ("smtp.gmail.com",   465, True),
    "outlook.com": ("smtp-mail.outlook.com", 587, False),
    "hotmail.com": ("smtp-mail.outlook.com", 587, False),
    "sina.com":    ("smtp.sina.com",    465, True),
    "sohu.com":    ("smtp.sohu.com",    465, True),
}

def get_smtp_config(email_addr):
    """根据邮箱地址自动识别SMTP服务器、端口、是否SSL"""
    domain = email_addr.strip().split("@")[-1].lower()
    if domain in SMTP_SERVERS:
        return SMTP_SERVERS[domain]
    # 默认尝试 smtp.{domain}:465 SSL
    return (f"smtp.{domain}", 465, True)

def smtp_connect(host, port, use_ssl):
    """根据配置创建SMTP连接"""
    if use_ssl:
        return smtplib.SMTP_SSL(host, port)
    else:
        srv = smtplib.SMTP(host, port)
        srv.starttls()
        return srv

# ── Colors ───────────────────────────────────
BG="#1A1A2E"; PANEL="#22223A"; CARD="#1C1C30"; ACCENT="#E94560"
ACCENT2="#2A5298"; TEXT="#F0F0FF"; TEXT_DIM="#BBBBDD"
GREEN="#2ECC71"; YELLOW="#F39C12"; RED_C="#E74C3C"; BORDER="#3A3A5C"
INPUT_BG="#F0F4FF"; INPUT_FG="#1A1A2E"
FT=("Georgia",20,"bold"); FL=("Segoe UI",10); FS=("Segoe UI",9)
FLOG=("Consolas",9); FB=("Segoe UI",10,"bold"); FH=("Georgia",12,"bold")

KW_DEFAULT = "\u5929\u6c14\u4e39\u4e24\u4ef6\u5957"
OUT_DEFAULT = "\u5929\u6c14\u4e39\u4e24\u4ef6\u5957_data.xlsx"

# ── Translations ─────────────────────────────
LANG = {
    "zh": {
        "title":        "淘宝数据采集工具",
        "subtitle":     "v5.3  自动采集 + 邮件发送",
        "ready":        "待机",
        "running":      "运行中",
        "settings":     "采集配置",
        "hint":         "所有选项已设默认值，直接点开始即可",
        "keyword":      "关键词",
        "mode":         "搜索模式",
        "mode_fuzzy":   "模糊搜索（推荐）",
        "mode_exact":   "精确搜索",
        "pages":        "采集页数",
        "sort":         "排序方式",
        "sort1":        "销量从高到低",
        "sort2":        "综合排序",
        "sort3":        "价格从低到高",
        "sort4":        "价格从高到低",
        "shops":        "店铺范围",
        "shops_all":    "全部（淘宝+天猫）",
        "shops_tmall":  "仅天猫",
        "price":        "价格（元）",
        "price_any":    "（0=不限）",
        "repeat":       "采集间隔",
        "rep1":         "仅采集一次",
        "rep2":         "每1小时重复",
        "rep3":         "每2小时重复",
        "rep4":         "每6小时重复",
        "saveas":       "保存文件",
        "steps":        "使用步骤：\n1. 双击 START_GUI.bat\n2. 在Chrome里登录淘宝\n3. 点下方开始按钮",
        "email_title":  "邮件发送",
        "qq_email":     "发件邮箱",
        "auth_code":    "授权码",
        "send_to":      "收件邮箱",
        "auth_hint":    "支持QQ/163/Gmail/Outlook等邮箱\n需在邮箱设置中开启SMTP并获取授权码",
        "btn_send":     "发送Excel到邮箱",
        "sending":      "发送中...",
        "btn_start":    "开始采集",
        "btn_stop":     "停止",
        "collected":    "已采集：",
        "log_title":    "运行日志",
        "clear":        "清空",
        "login_title":  "确认登录",
        "login_msg":    "请先在Chrome里登录淘宝账号\n登录完成后点击下方按钮，程序全自动完成采集",
        "login_btn":    "已登录，开始采集！",
        "cancel":       "取消",
        "quit_msg":     "采集正在运行，确认退出？",
        "quit_title":   "退出",
        "warn_kw":      "请输入搜索关键词",
        "warn_kw_max":  "最多添加10个关键词，过多会增加风控风险",
        "warn_email":   "请填写发件邮箱、授权码和收件邮箱",
        "warn_file":    "未找到数据文件，请先运行采集",
        "email_ok":     "邮件发送成功！",
        "auto_send":    "采集完自动发邮件",
        "btn_report":   "生成日报并发送",
        "own_shop":     "自家店铺名",
        "report_ok":    "日报已生成并发送！",
        "report_fail":  "日报生成失败：",
        "no_data":      "暂无历史数据可对比，请先采集两天以上数据",
        "email_fail":   "邮件发送失败：",
        "wx_title":     "微信预警推送",
        "wx_key":       "Token/Key",
        "wx_enable":    "启用降价预警推送",
        "wx_threshold": "降价超过",
        "wx_pct":       "% 时推送",
        "wx_test":      "发送测试消息",
        "wx_hint_sc":   "sct.ftqq.com 免费获取",
        "wx_hint_pp":   "pushplus.plus 微信扫码获取Token",
        "wx_test_ok":   "测试消息已发送，请查看微信",
        "wx_type_sc":   "Server酱",
        "wx_type_pp":   "PushPlus",
        "wx_test_fail": "发送失败：",
        "wx_no_key":    "请先填写推送Token/Key",
        "rank_drop":    "排名下滑超过",
        "rank_pos":     "位时预警",
        "rank_enable":  "启用排名下滑预警",
        "sales_title":  "销量暴增监控",
        "sales_enable": "启用竞品销量暴增预警",
        "sales_drop":   "销量增长超过",
        "sales_pct":    "% 时预警",
        "multi_title":  "采集选项（勾选后与淘宝同步运行）",
        "multi_jd":     "同步采集京东",
        "multi_hw":     "采集热门关键词",
        "multi_review": "采集竞品评价分析",
        "multi_btn":    "单独触发多平台采集",
        "sched_title":  "定时自动采集",
        "sched_enable": "启用每日定时采集",
        "sched_time":   "采集时间",
        "sched_hint":   "（电脑开机时自动触发）",
        "lang_label":   "语言",
    },
    "en": {
        "title":        "Taobao Scraper",
        "subtitle":     "v5.3  Auto Collect + Email",
        "ready":        "Ready",
        "running":      "Running",
        "settings":     "Settings",
        "hint":         "Defaults are set — just click START",
        "keyword":      "Keyword",
        "mode":         "Search Mode",
        "mode_fuzzy":   "Fuzzy (Recommended)",
        "mode_exact":   "Exact Match",
        "pages":        "Pages",
        "sort":         "Sort By",
        "sort1":        "Sales High to Low",
        "sort2":        "Default",
        "sort3":        "Price Low to High",
        "sort4":        "Price High to Low",
        "shops":        "Shops",
        "shops_all":    "All (Taobao + Tmall)",
        "shops_tmall":  "Tmall Only",
        "price":        "Price (CNY)",
        "price_any":    "(0 = any)",
        "repeat":       "Repeat",
        "rep1":         "Run Once",
        "rep2":         "Every 1 Hour",
        "rep3":         "Every 2 Hours",
        "rep4":         "Every 6 Hours",
        "saveas":       "Save As",
        "steps":        "Steps:\n1. Run START_GUI.bat\n2. Login Taobao in Chrome\n3. Click START below",
        "email_title":  "Send by Email",
        "qq_email":     "Sender Email",
        "auth_code":    "Auth Code",
        "send_to":      "Send To",
        "auth_hint":    "Supports QQ/163/Gmail/Outlook etc.\nEnable SMTP in email settings to get auth code",
        "btn_send":     "Send Excel to Email",
        "sending":      "Sending...",
        "btn_start":    "START",
        "btn_stop":     "STOP",
        "collected":    "Collected: ",
        "log_title":    "Log",
        "clear":        "Clear",
        "login_title":  "Login Required",
        "login_msg":    "Please login to Taobao in Chrome first.\nThen click the button — the program does the rest.",
        "login_btn":    "Logged In — Start!",
        "cancel":       "Cancel",
        "quit_msg":     "Scraper is running. Quit anyway?",
        "quit_title":   "Quit",
        "warn_kw":      "Please enter a keyword",
        "warn_kw_max":  "Max 10 keywords allowed to reduce risk of anti-bot detection",
        "warn_email":   "Please fill in sender email, auth code and recipient",
        "warn_file":    "Data file not found. Run scraper first.",
        "email_ok":     "Email sent successfully!",
        "auto_send":    "Auto-send email after each round",
        "btn_report":   "Generate & Send Daily Report",
        "own_shop":     "My Shop Name",
        "report_ok":    "Daily report sent!",
        "report_fail":  "Report failed: ",
        "no_data":      "No historical data yet. Please collect data for 2+ days first",
        "email_fail":   "Email failed: ",
        "wx_title":     "WeChat Alert",
        "wx_key":       "Token/Key",
        "wx_enable":    "Enable price-drop alert",
        "wx_threshold": "Alert when drop >",
        "wx_pct":       "%",
        "wx_test":      "Send Test",
        "wx_hint_sc":   "Get free key at sct.ftqq.com",
        "wx_hint_pp":   "Get token at pushplus.plus (WeChat scan)",
        "wx_test_ok":   "Test sent! Check WeChat",
        "wx_type_sc":   "ServerChan",
        "wx_type_pp":   "PushPlus",
        "wx_test_fail": "Failed: ",
        "wx_no_key":    "Please enter push Token/Key",
        "rank_drop":    "Alert when rank drops >",
        "rank_pos":     "positions",
        "rank_enable":  "Enable rank-drop alert",
        "sales_title":  "Sales Spike Monitor",
        "sales_enable": "Enable sales spike alert",
        "sales_drop":   "Alert when sales grow >",
        "sales_pct":    "%",
        "multi_title":  "Sync Options (runs with Taobao)",
        "multi_jd":     "Sync JD.com",
        "multi_review": "Scrape review analysis",
        "multi_btn":    "Run Separately",
        "sched_title":  "Auto Schedule",
        "sched_enable": "Enable daily auto-scrape",
        "sched_time":   "Run at",
        "sched_hint":   "(triggers when PC is on)",
        "lang_label":   "Language",
    },
    "ko": {
        "title":        "타오바오 데이터 수집기",
        "subtitle":     "v5.3  자동 수집 + 이메일",
        "ready":        "대기",
        "running":      "실행 중",
        "settings":     "수집 설정",
        "hint":         "기본값이 설정되어 있습니다. 바로 시작하세요",
        "keyword":      "키워드",
        "mode":         "검색 모드",
        "mode_fuzzy":   "유사 검색 (추천)",
        "mode_exact":   "정확히 일치",
        "pages":        "페이지 수",
        "sort":         "정렬 방식",
        "sort1":        "판매량 높은 순",
        "sort2":        "기본 정렬",
        "sort3":        "가격 낮은 순",
        "sort4":        "가격 높은 순",
        "shops":        "쇼핑몰",
        "shops_all":    "전체 (타오바오+티몰)",
        "shops_tmall":  "티몰만",
        "price":        "가격 (위안)",
        "price_any":    "(0 = 제한없음)",
        "repeat":       "반복 주기",
        "rep1":         "1회만 실행",
        "rep2":         "1시간마다",
        "rep3":         "2시간마다",
        "rep4":         "6시간마다",
        "saveas":       "저장 파일",
        "steps":        "사용 방법:\n1. START_GUI.bat 실행\n2. Chrome에서 타오바오 로그인\n3. 아래 시작 버튼 클릭",
        "email_title":  "이메일 발송",
        "qq_email":     "보내는 이메일",
        "auth_code":    "인증 코드",
        "send_to":      "받는 이메일",
        "auth_hint":    "QQ/163/Gmail/Outlook 등 지원\n이메일 설정에서 SMTP 활성화 후 인증 코드 발급",
        "btn_send":     "엑셀 이메일 발송",
        "sending":      "발송 중...",
        "btn_start":    "수집 시작",
        "btn_stop":     "정지",
        "collected":    "수집됨: ",
        "log_title":    "실행 로그",
        "clear":        "지우기",
        "login_title":  "로그인 확인",
        "login_msg":    "Chrome에서 타오바오에 로그인해 주세요.\n완료 후 버튼을 클릭하면 자동으로 수집됩니다.",
        "login_btn":    "로그인 완료 — 시작!",
        "cancel":       "취소",
        "quit_msg":     "수집 중입니다. 종료하시겠습니까?",
        "quit_title":   "종료",
        "warn_kw":      "키워드를 입력해 주세요",
        "warn_kw_max":  "키워드는 최대 10개까지 추가할 수 있습니다 (봇 탐지 방지)",
        "warn_email":   "보내는 이메일, 인증 코드, 수신자를 입력해 주세요",
        "warn_file":    "데이터 파일을 찾을 수 없습니다. 먼저 수집을 실행하세요.",
        "email_ok":     "이메일 발송 성공!",
        "auto_send":    "수집 후 자동 이메일 발송",
        "btn_report":   "일간 보고서 생성 및 발송",
        "own_shop":     "자사 매장명",
        "report_ok":    "일간 보고서 발송 완료!",
        "report_fail":  "보고서 실패: ",
        "no_data":      "비교 데이터 없음. 2일 이상 수집 필요",
        "email_fail":   "이메일 발송 실패: ",
        "wx_title":     "위챗 알림",
        "wx_key":       "Token/Key",
        "wx_enable":    "가격 인하 알림 활성화",
        "wx_threshold": "인하율",
        "wx_pct":       "% 초과 시 알림",
        "wx_test":      "테스트 전송",
        "wx_hint_sc":   "sct.ftqq.com 에서 무료 발급",
        "wx_hint_pp":   "pushplus.plus 위챗 스캔으로 Token 발급",
        "wx_test_ok":   "테스트 전송 완료, 위챗 확인",
        "wx_type_sc":   "Server酱",
        "wx_type_pp":   "PushPlus",
        "wx_test_fail": "실패: ",
        "wx_no_key":    "푸시 Token/Key를 입력하세요",
        "rank_drop":    "순위 하락",
        "rank_pos":     "위 초과 시 알림",
        "rank_enable":  "순위 하락 알림 활성화",
        "sales_title":  "판매량 급증 모니터",
        "sales_enable": "경쟁사 판매량 급증 알림",
        "sales_drop":   "판매량 증가",
        "sales_pct":    "% 초과 시 알림",
        "multi_title":  "동기화 옵션",
        "multi_jd":     "JD 동기화",
        "multi_review": "리뷰 분석 수집",
        "multi_btn":    "단독 실행",
        "sched_title":  "자동 예약 수집",
        "sched_enable": "매일 자동 수집 활성화",
        "sched_time":   "수집 시간",
        "sched_hint":   "(PC 켜져 있을 때 자동 실행)",
        "lang_label":   "언어",
    },
}
log_q = queue.Queue()


class ToggleBtn(tk.Frame):
    """自定义勾选按钮，彻底替代原生Checkbutton，解决Windows下勾选不显示问题"""
    def __init__(self, parent, text, variable, bg=PANEL, fg=TEXT,
                 font=None, command=None, **kw):
        super().__init__(parent, bg=bg, cursor="hand2", **kw)
        self._var  = variable
        self._text = text
        self._bg   = bg
        self._fg   = fg
        self._font = font or ("Segoe UI", 9)
        self._cmd  = command

        self._box = tk.Label(self, width=2, relief="flat",
                             font=("Segoe UI", 10, "bold"), cursor="hand2")
        self._box.pack(side="left", padx=(0, 6))

        self._lbl = tk.Label(self, text=text, bg=bg, fg=fg,
                             font=self._font, cursor="hand2")
        self._lbl.pack(side="left")

        self._box.bind("<Button-1>", self._toggle)
        self._lbl.bind("<Button-1>", self._toggle)
        self.bind("<Button-1>", self._toggle)
        self._render()
        variable.trace_add("write", lambda *_: self._render())

    def _render(self):
        if self._var.get():
            self._box.config(text="✓", bg="#4CAF50", fg="white",
                             relief="flat", padx=4, pady=1)
        else:
            self._box.config(text=" ", bg="#555577", fg="#555577",
                             relief="flat", padx=4, pady=1)

    def _toggle(self, event=None):
        self._var.set(not self._var.get())
        if self._cmd:
            self._cmd()


class QH(logging.Handler):
    def emit(self, r):
        log_q.put(self.format(r))

class App:
    def __init__(self, root):
        self.root = root
        self.lang = "zh"
        self.stop_ev = threading.Event()
        self.login_ev = threading.Event()
        self.is_running = False
        self.total = 0
        self.last_file = ""
        self._setup_log()
        self._build()
        self._load_config()
        self._start_scheduler()
        self._poll()

    def t(self, key):
        return LANG[self.lang].get(key, key)

    def _setup_log(self):
        self.log = logging.getLogger("sc")
        self.log.setLevel(logging.INFO)
        self.log.handlers.clear()
        h = QH()
        h.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))
        self.log.addHandler(h)
        fh = logging.FileHandler("scraper_log.txt", encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))
        self.log.addHandler(fh)

    # ══════════════════════════════════════════════
    #  定时自动采集
    # ══════════════════════════════════════════════

    # ══════════════════════════════════════════════
    #  本地配置持久化
    # ══════════════════════════════════════════════

    def _config_path(self):
        import os
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "scraper_config.json")

    def _load_config(self):
        """启动时从本地JSON读取上次的配置"""
        import json, os
        try:
            path = self._config_path()
            if not os.path.exists(path):
                return
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            if cfg.get("wx_key"):        self.vwxkey.set(cfg["wx_key"])
            if cfg.get("wx_threshold"):  self.vwx_threshold.set(cfg["wx_threshold"])
            if cfg.get("wx_enable") is not None:  self.vauto_wx.set(cfg["wx_enable"])
            if cfg.get("wx_type"):       self.vwx_type.set(cfg["wx_type"])
            if cfg.get("sched_enable") is not None: self.vauto_schedule.set(cfg["sched_enable"])
            if cfg.get("sched_hour"):    self.vsched_hour.set(cfg["sched_hour"])
            if cfg.get("sched_min"):     self.vsched_min.set(cfg["sched_min"])
            if cfg.get("email_from"):    self.vefrom.set(cfg["email_from"])
            if cfg.get("email_pwd"):     self.vepwd.set(cfg["email_pwd"])
            if cfg.get("email_to"):      self.veto.set(cfg["email_to"])
            if cfg.get("own_shop"):      self.vown.set(cfg["own_shop"])
            if cfg.get("keywords"):
                self.kw_listbox.delete(0, "end")
                for kw in cfg["keywords"]:
                    if kw.strip():
                        self.kw_listbox.insert("end", kw.strip())
                if self.kw_listbox.size() > 0:
                    self.vkw.set(self.kw_listbox.get(0))
            elif cfg.get("keyword"):
                self.kw_listbox.delete(0, "end")
                self.kw_listbox.insert("end", cfg["keyword"])
                self.vkw.set(cfg["keyword"])
            if cfg.get("output_file"):   self.vout.set(cfg["output_file"])
            if cfg.get("auto_email") is not None: self.vauto.set(cfg["auto_email"])
            if cfg.get("rank_enable") is not None: self.vauto_rank.set(cfg["rank_enable"])
            if cfg.get("rank_threshold"):  self.vrank_threshold.set(cfg["rank_threshold"])
            if cfg.get("sales_enable") is not None: self.vauto_sales.set(cfg["sales_enable"])
            if cfg.get("sales_threshold"): self.vsales_threshold.set(cfg["sales_threshold"])
            # 多平台选项（默认False，用户改后记住）
            if cfg.get("mp_jd")  is not None: self.vmp_jd.set(cfg["mp_jd"])
            if cfg.get("mp_rev") is not None: self.vmp_rev.set(cfg["mp_rev"])
            if cfg.get("mp_hw")  is not None: self.vmp_hw.set(cfg["mp_hw"])
        except Exception as e:
            self.log.debug(f"配置读取失败: {e}")

    def _save_config(self):
        """保存当前配置到本地JSON"""
        import json
        try:
            cfg = {
                "wx_key":       self.vwxkey.get().strip(),
                "wx_threshold": self.vwx_threshold.get().strip(),
                "wx_enable":    self.vauto_wx.get(),
                "wx_type":      self.vwx_type.get(),
                "sched_enable": self.vauto_schedule.get(),
                "sched_hour":   self.vsched_hour.get().strip(),
                "sched_min":    self.vsched_min.get().strip(),
                "email_from":   self.vefrom.get().strip(),
                "email_pwd":    self.vepwd.get().strip(),
                "email_to":     self.veto.get().strip(),
                "own_shop":     self.vown.get().strip(),
                "keyword":      self.vkw.get().strip(),
                "keywords":     self._get_keywords(),
                "output_file":  self.vout.get().strip(),
                "auto_email":   self.vauto.get(),
                "rank_enable":  self.vauto_rank.get(),
                "rank_threshold":  self.vrank_threshold.get().strip(),
                "sales_enable":    self.vauto_sales.get(),
                "sales_threshold": self.vsales_threshold.get().strip(),
                "mp_jd":           self.vmp_jd.get(),
                "mp_rev":          self.vmp_rev.get(),
                "mp_hw":           self.vmp_hw.get(),
            }
            with open(self._config_path(), "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log.debug(f"配置保存失败: {e}")

    def _start_scheduler(self):
        """启动后台定时检测线程，同时检测今天是否需要补跑"""
        t = threading.Thread(target=self._scheduler_loop, daemon=True)
        t.start()
        # 启动后5秒检测一次：今天还没采集过 且 启用了定时 → 自动补跑
        self.root.after(5000, self._check_on_startup)
        self._chromedriver_path = None
        import threading as _th
        _th.Thread(target=self._preload_chromedriver, daemon=True).start()

    def _preload_chromedriver(self):
        """GUI启动时后台静默下载/缓存ChromeDriver，点采集时直接用"""
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            self._chromedriver_path = ChromeDriverManager().install()
            self.log.info("ChromeDriver 已就绪，可以开始采集")
        except Exception as e:
            self.log.warning(f"ChromeDriver 预加载失败: {e}")

    def _check_on_startup(self):
        """开机/重启工具后，检测今天是否漏采"""
        if not self.vauto_schedule.get():
            return
        import datetime
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        if not self._check_collected_today(today_str):
            self.log.info("[定时] 检测到今天尚未采集，自动补跑...")
            self._auto_start()

    def _scheduler_loop(self):
        """每分钟检查一次，到达预定时间且今天未采集则自动触发"""
        import datetime
        while True:
            time.sleep(60)  # 每分钟检查一次
            if not self.vauto_schedule.get():
                continue
            if self.is_running:
                continue
            now = datetime.datetime.now()
            target_h = int(self.vsched_hour.get()) if self.vsched_hour.get().isdigit() else 9
            target_m = int(self.vsched_min.get()) if self.vsched_min.get().isdigit() else 0
            # 到达目标时间（允许2分钟误差，防止睡眠唤醒后错过整点）
            now_total  = now.hour * 60 + now.minute
            target_total = target_h * 60 + target_m
            if abs(now_total - target_total) <= 2:
                today_str = now.strftime("%Y-%m-%d")
                if self._check_collected_today(today_str):
                    self.log.info(f"[定时] {today_str} 已采集，跳过")
                    continue
                self.log.info(f"[定时] 到达预定时间 {target_h:02d}:{target_m:02d}，自动开始采集...")
                self.root.after(0, self._auto_start)

    def _check_collected_today(self, today_str):
        """检查今天是否已采集过（读取Excel的采集时间列）"""
        try:
            import pandas as pd
            fp = self.vout.get().strip()
            if not fp or not os.path.exists(fp):
                return False
            df = pd.read_excel(fp, sheet_name="采集数据", usecols=["采集时间"])
            if df.empty:
                return False
            latest = str(df["采集时间"].dropna().iloc[-1])
            return latest.startswith(today_str)
        except Exception:
            return False

    def _auto_start(self):
        """定时触发的自动采集（等同于点了开始按钮）"""
        if self.is_running:
            return
        cfg = self._cfg()
        if not cfg["keyword"]:
            self.log.warning("[定时] 关键词为空，跳过自动采集")
            return
        # 强制设为"仅采集一次"，避免定时+循环叠加
        cfg["interval_hr"] = 0
        self.stop_ev.clear()
        self.is_running = True
        self.bstart.configure(state="disabled")
        self.bstop.configure(state="normal")
        self.log.info("[定时] 自动采集启动")
        threading.Thread(target=self._worker, args=(cfg,), daemon=True).start()

    def _poll(self):
        while not log_q.empty():
            msg = log_q.get_nowait()
            self._wlog(msg)
        self.root.after(200, self._poll)

    def _wlog(self, msg):
        self.lb.configure(state="normal")
        if any(x in msg for x in ["success","Success","Saved","Done","sent","Connected","完成","成功"]):
            tag = "green"
        elif any(x in msg for x in ["Error","Failed","error","failed","错误"]):
            tag = "red"
        elif any(x in msg for x in ["Warning","warn","blocked","拦截"]):
            tag = "yellow"
        else:
            tag = "normal"
        self.lb.insert("end", msg+"\n", tag)
        self.lb.see("end")
        self.lb.configure(state="disabled")

    def _build(self):
        self.root.title(self.t("title"))
        self.root.geometry("1280x800")
        self.root.configure(bg=BG)

        # Header
        hdr = tk.Frame(self.root, bg=ACCENT2, height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        self.lbl_title = tk.Label(hdr, text=self.t("title"), font=FT, bg=ACCENT2, fg=TEXT)
        self.lbl_title.pack(side="left", padx=24, pady=10)
        self.lbl_sub = tk.Label(hdr, text=self.t("subtitle"), font=FS, bg=ACCENT2, fg=TEXT_DIM)
        self.lbl_sub.pack(side="left", pady=20)

        # Language switcher
        lang_fr = tk.Frame(hdr, bg=ACCENT2)
        lang_fr.pack(side="right", padx=16)
        for code, label in [("zh","中文"), ("en","EN"), ("ko","한국어")]:
            tk.Button(lang_fr, text=label, font=FS, bg=BORDER, fg=TEXT,
                      relief="flat", bd=0, padx=10, pady=4, cursor="hand2",
                      command=lambda c=code: self._switch_lang(c)).pack(side="left", padx=2)

        self.sdot = tk.Label(hdr, text="\u25cf", font=("Segoe UI",14), bg=ACCENT2, fg=TEXT_DIM)
        self.sdot.pack(side="right", padx=4)
        self.slbl = tk.Label(hdr, text=self.t("ready"), font=FS, bg=ACCENT2, fg=TEXT_DIM)
        self.slbl.pack(side="right")

        # Body
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=16, pady=12)

        # Left scrollable panel
        lo = tk.Frame(body, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        lo.pack(side="left", fill="y", padx=(0,10))
        lo.pack_propagate(False)
        lo.configure(width=480)
        cv = tk.Canvas(lo, bg=PANEL, highlightthickness=0, width=462)
        sb = tk.Scrollbar(lo, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cv.pack(side="left", fill="both", expand=True)
        self.lf = tk.Frame(cv, bg=PANEL)
        cv.create_window((0,0), window=self.lf, anchor="nw")
        self.lf.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        self._left_outer = lo
        self._left_canvas = cv

        # Right log panel
        rp = tk.Frame(body, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        rp.pack(side="left", fill="both", expand=True)
        self._right_panel = rp
        lh = tk.Frame(rp, bg=CARD)
        lh.pack(fill="x", padx=16, pady=(12,4))
        self.lbl_log = tk.Label(lh, text=self.t("log_title"), font=FH, bg=CARD, fg=TEXT)
        self.lbl_log.pack(side="left")
        self.btn_clear = tk.Button(lh, text=self.t("clear"), font=FS, bg=BORDER, fg=TEXT_DIM,
                  relief="flat", bd=0, padx=8, cursor="hand2",
                  command=lambda: [self.lb.configure(state="normal"),
                                   self.lb.delete("1.0","end"),
                                   self.lb.configure(state="disabled")])
        self.btn_clear.pack(side="right")
        self.lb = scrolledtext.ScrolledText(rp, bg="#0A0A14", fg=TEXT, font=FLOG,
                  relief="flat", bd=0, wrap="word", state="disabled")
        self.lb.pack(fill="both", expand=True, padx=12, pady=(0,12))
        for t, c in [("green",GREEN),("red",RED_C),("yellow",YELLOW),("normal",TEXT)]:
            self.lb.tag_configure(t, foreground=c)

        # 鼠标滚轮：在左/右面板任意位置都可滚动
        def _on_mousewheel(event):
            w = event.widget
            while w:
                if w is self._left_outer or w is self._left_canvas:
                    self._left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    return "break"
                if w is self._right_panel:
                    self.lb.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    return "break"
                w = w.master
        self.root.bind_all("<MouseWheel>", _on_mousewheel)

        self._fill_left()
        self._build_btns()

    def _row(self, p, label_text, factory):
        f = tk.Frame(p, bg=PANEL)
        f.pack(fill="x", padx=20, pady=4)
        lbl = tk.Label(f, text=label_text, font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        lbl.pack(side="left")
        w = factory(f)
        w.pack(side="left", fill="x", expand=True)
        return lbl, w

    def _entry(self, p, tv, show=None):
        kw = {"show": show} if show else {}
        return tk.Entry(p, textvariable=tv, bg=ACCENT2, fg=TEXT, font=FL,
                        insertbackground=TEXT, relief="flat", bd=4, **kw)

    def _combo(self, p, vals, tv):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("D.TCombobox", fieldbackground=INPUT_BG, background=INPUT_BG,
                    foreground=INPUT_FG, arrowcolor=ACCENT2, selectbackground=ACCENT2, selectforeground=TEXT)
        return ttk.Combobox(p, values=vals, textvariable=tv, style="D.TCombobox", font=FL, state="readonly")

    def _fill_left(self):
        p = self.lf

        # 保存当前用户输入的值（切换语言时不丢失）
        _saved = {}
        _saved_kw_list = []
        for attr in ("vkw", "vpg", "vpmin", "vpmax", "vout", "vefrom", "vepwd", "veto",
                      "vwxkey", "vwx_threshold", "vwx_type", "vrank_threshold", "vsales_threshold",
                      "vsched_hour", "vsched_min", "vown", "vkw_input"):
            if hasattr(self, attr):
                try: _saved[attr] = getattr(self, attr).get()
                except Exception: pass
        for attr in ("vauto", "vauto_wx", "vauto_rank", "vauto_sales", "vauto_schedule",
                      "vmp_jd", "vmp_rev", "vmp_hw"):
            if hasattr(self, attr):
                try: _saved[attr] = getattr(self, attr).get()
                except Exception: pass
        if hasattr(self, "kw_listbox"):
            try: _saved_kw_list = list(self.kw_listbox.get(0, "end"))
            except Exception: pass

        # Clear existing widgets
        for w in p.winfo_children():
            w.destroy()

        self.lbl_settings = tk.Label(p, text=self.t("settings"), font=FH, bg=PANEL, fg=TEXT)
        self.lbl_settings.pack(anchor="w", padx=20, pady=(16,4))
        self.lbl_hint = tk.Label(p, text=self.t("hint"), font=FS, bg=PANEL, fg=TEXT_DIM)
        self.lbl_hint.pack(anchor="w", padx=20, pady=(0,12))

        # Keywords（多关键词）
        self.vkw = tk.StringVar(value=_saved.get("vkw", KW_DEFAULT))  # 保留用户输入
        fkw = tk.Frame(p, bg=PANEL); fkw.pack(fill="x", padx=20, pady=4)
        self.lkw = tk.Label(fkw, text=self.t("keyword"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="nw")
        self.lkw.pack(side="left", anchor="n")

        # 右侧：列表框 + 按钮
        fkw_right = tk.Frame(fkw, bg=PANEL); fkw_right.pack(side="left", fill="x", expand=True)

        # 列表框（占满整行宽度）
        fkw_list = tk.Frame(fkw_right, bg=PANEL); fkw_list.pack(fill="x")
        self.kw_listbox = tk.Listbox(
            fkw_list, bg="#2A2A45", fg=TEXT, font=FL,
            relief="flat", bd=0, height=4,
            selectbackground=ACCENT, selectforeground="white",
            activestyle="none", highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT)
        self.kw_listbox.pack(side="left", fill="x", expand=True, pady=1)
        sb = tk.Scrollbar(fkw_list, orient="vertical", command=self.kw_listbox.yview,
                          bg=BORDER, troughcolor=PANEL, width=10)
        sb.pack(side="right", fill="y", pady=1)
        self.kw_listbox.configure(yscrollcommand=sb.set)
        # 恢复关键词列表
        if _saved_kw_list:
            for kw in _saved_kw_list:
                self.kw_listbox.insert("end", kw)
        else:
            self.kw_listbox.insert("end", KW_DEFAULT)

        # 输入框一行
        fkw_input = tk.Frame(fkw_right, bg=PANEL); fkw_input.pack(fill="x", pady=(5,0))
        self.vkw_input = tk.StringVar(value=_saved.get("vkw_input", ""))
        tk.Entry(fkw_input, textvariable=self.vkw_input, bg="#2A2A45", fg=TEXT, font=FL,
                 insertbackground=TEXT, relief="flat", bd=0,
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT).pack(side="left", fill="x", expand=True, ipady=4)

        # 添加 / 删除 按钮并排在输入框右侧
        tk.Button(fkw_input, text="添加", bg=ACCENT2, fg="white", font=FS,
                  relief="flat", bd=0, padx=10, cursor="hand2",
                  activebackground="#3A6BC8", activeforeground="white",
                  command=self._kw_add).pack(side="left", padx=(4,0))
        tk.Button(fkw_input, text="删除", bg="#C0392B", fg="white", font=FS,
                  relief="flat", bd=0, padx=10, cursor="hand2",
                  activebackground="#E74C3C", activeforeground="white",
                  command=self._kw_del).pack(side="left", padx=(4,0))

        # Mode
        self.vex = tk.StringVar(value=self.t("mode_fuzzy"))
        f2 = tk.Frame(p, bg=PANEL); f2.pack(fill="x", padx=20, pady=4)
        self.lmode = tk.Label(f2, text=self.t("mode"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lmode.pack(side="left")
        self.cb_mode = self._combo(f2, [self.t("mode_fuzzy"), self.t("mode_exact")], self.vex)
        self.cb_mode.pack(side="left", fill="x", expand=True)

        # Pages
        self.vpg = tk.StringVar(value=_saved.get("vpg", "3"))
        f3 = tk.Frame(p, bg=PANEL); f3.pack(fill="x", padx=20, pady=4)
        self.lpg = tk.Label(f3, text=self.t("pages"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lpg.pack(side="left")
        self._combo(f3, ["1","2","3","5","10"], self.vpg).pack(side="left", fill="x", expand=True)

        # Sort
        self.vso = tk.StringVar(value=self.t("sort1"))
        f4 = tk.Frame(p, bg=PANEL); f4.pack(fill="x", padx=20, pady=4)
        self.lso = tk.Label(f4, text=self.t("sort"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lso.pack(side="left")
        self.cb_sort = self._combo(f4, [self.t("sort1"),self.t("sort2"),self.t("sort3"),self.t("sort4")], self.vso)
        self.cb_sort.pack(side="left", fill="x", expand=True)

        # Shops
        self.vsh = tk.StringVar(value=self.t("shops_all"))
        f5 = tk.Frame(p, bg=PANEL); f5.pack(fill="x", padx=20, pady=4)
        self.lsh = tk.Label(f5, text=self.t("shops"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lsh.pack(side="left")
        self.cb_shop = self._combo(f5, [self.t("shops_all"), self.t("shops_tmall")], self.vsh)
        self.cb_shop.pack(side="left", fill="x", expand=True)

        # Price
        pf = tk.Frame(p, bg=PANEL); pf.pack(fill="x", padx=20, pady=4)
        self.lprice = tk.Label(pf, text=self.t("price"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lprice.pack(side="left")
        self.vpmin = tk.StringVar(value=_saved.get("vpmin", "0"))
        self.vpmax = tk.StringVar(value=_saved.get("vpmax", "0"))
        tk.Entry(pf, textvariable=self.vpmin, width=6, bg=ACCENT2, fg=TEXT, font=FL,
                 relief="flat", bd=4, insertbackground=TEXT).pack(side="left")
        tk.Label(pf, text=" - ", bg=PANEL, fg=TEXT_DIM, font=FL).pack(side="left")
        tk.Entry(pf, textvariable=self.vpmax, width=6, bg=ACCENT2, fg=TEXT, font=FL,
                 relief="flat", bd=4, insertbackground=TEXT).pack(side="left")
        self.lprice_any = tk.Label(pf, text=self.t("price_any"), bg=PANEL, fg=TEXT_DIM, font=FS)
        self.lprice_any.pack(side="left", padx=4)

        # Repeat
        self.viv = tk.StringVar(value=self.t("rep1"))
        f6 = tk.Frame(p, bg=PANEL); f6.pack(fill="x", padx=20, pady=4)
        self.lrep = tk.Label(f6, text=self.t("repeat"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lrep.pack(side="left")
        self.cb_rep = self._combo(f6, [self.t("rep1"),self.t("rep2"),self.t("rep3"),self.t("rep4")], self.viv)
        self.cb_rep.pack(side="left", fill="x", expand=True)

        # Save As
        ff = tk.Frame(p, bg=PANEL); ff.pack(fill="x", padx=20, pady=4)
        self.lsave = tk.Label(ff, text=self.t("saveas"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lsave.pack(side="left")
        self.vout = tk.StringVar(value=_saved.get("vout", OUT_DEFAULT))
        tk.Entry(ff, textvariable=self.vout, bg=INPUT_BG, fg=INPUT_FG, font=FL,
                 insertbackground=INPUT_FG, relief="flat", bd=4).pack(side="left", fill="x", expand=True)
        tk.Button(ff, text="...", bg=BORDER, fg=TEXT, font=FS, relief="flat", bd=0,
                  padx=6, cursor="hand2", command=self._browse).pack(side="left", padx=(4,0))

        tk.Frame(p, bg=BORDER, height=1).pack(fill="x", padx=20, pady=12)
        self.lbl_steps = tk.Label(p, text=self.t("steps"), font=FS, bg=PANEL, fg=TEXT_DIM, justify="left")
        self.lbl_steps.pack(anchor="w", padx=20, pady=(0,8))

        # Email section
        tk.Frame(p, bg=BORDER, height=1).pack(fill="x", padx=20, pady=(4,10))
        self.lbl_email_title = tk.Label(p, text=self.t("email_title"), font=FH, bg=PANEL, fg=TEXT)
        self.lbl_email_title.pack(anchor="w", padx=20, pady=(0,6))

        self.vefrom = tk.StringVar(value=_saved.get("vefrom", ""))
        fe1 = tk.Frame(p, bg=PANEL); fe1.pack(fill="x", padx=20, pady=4)
        self.lqq = tk.Label(fe1, text=self.t("qq_email"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lqq.pack(side="left")
        tk.Entry(fe1, textvariable=self.vefrom, bg=INPUT_BG, fg=INPUT_FG, font=FL,
                 insertbackground=INPUT_FG, relief="flat", bd=4).pack(side="left", fill="x", expand=True)

        self.vepwd = tk.StringVar(value=_saved.get("vepwd", ""))
        fe2 = tk.Frame(p, bg=PANEL); fe2.pack(fill="x", padx=20, pady=4)
        self.lauth = tk.Label(fe2, text=self.t("auth_code"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lauth.pack(side="left")
        tk.Entry(fe2, textvariable=self.vepwd, bg=INPUT_BG, fg=INPUT_FG, font=FL,
                 insertbackground=INPUT_FG, relief="flat", bd=4, show="*").pack(side="left", fill="x", expand=True)

        self.veto = tk.StringVar(value=_saved.get("veto", ""))
        fe3 = tk.Frame(p, bg=PANEL); fe3.pack(fill="x", padx=20, pady=4)
        self.lto = tk.Label(fe3, text=self.t("send_to"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lto.pack(side="left")
        tk.Entry(fe3, textvariable=self.veto, bg=INPUT_BG, fg=INPUT_FG, font=FL,
                 insertbackground=INPUT_FG, relief="flat", bd=4).pack(side="left", fill="x", expand=True)

        self.lbl_auth_hint = tk.Label(p, text=self.t("auth_hint"), font=FS, bg=PANEL, fg=TEXT_DIM, justify="left")
        self.lbl_auth_hint.pack(anchor="w", padx=20, pady=(2,8))

        # 自动发送勾选框
        self.vauto = tk.BooleanVar(value=_saved.get("vauto", False))
        fauto = tk.Frame(p, bg=PANEL); fauto.pack(fill="x", padx=20, pady=(0,6))
        self.cb_auto = ToggleBtn(fauto, text=self.t("auto_send"),
            variable=self.vauto, bg=PANEL, fg=TEXT_DIM, font=FS)
        self.cb_auto.pack(side="left")

        self.bsend = tk.Button(p, text=self.t("btn_send"),
            font=FB, bg=ACCENT2, fg=TEXT, relief="flat", bd=0,
            padx=16, pady=8, cursor="hand2", command=self._send_email)
        self.bsend.pack(padx=20, pady=(0,8), fill="x")

        # ── 微信预警推送 ──
        sep2 = tk.Frame(p, bg="#333333", height=1); sep2.pack(fill="x", padx=20, pady=(12,0))
        fwx0 = tk.Frame(p, bg=PANEL); fwx0.pack(fill="x", padx=20, pady=(8,2))
        tk.Label(fwx0, text="📱 "+self.t("wx_title"), font=FB, bg=PANEL, fg="#07C160").pack(side="left")

        # 推送方式选择
        self.vwx_type = tk.StringVar(value=_saved.get("vwx_type", "server_chan"))
        fwx_type = tk.Frame(p, bg=PANEL); fwx_type.pack(fill="x", padx=20, pady=2)
        tk.Label(fwx_type, text="推送方式", font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w").pack(side="left")
        wx_type_combo = ttk.Combobox(fwx_type, textvariable=self.vwx_type, state="readonly",
            values=["server_chan", "pushplus"], width=15)
        wx_type_combo.pack(side="left")
        # 显示友好名称的映射标签
        self._wx_type_label = tk.Label(fwx_type, text="", font=FS, bg=PANEL, fg=TEXT_DIM)
        self._wx_type_label.pack(side="left", padx=(8,0))

        fwx1 = tk.Frame(p, bg=PANEL); fwx1.pack(fill="x", padx=20, pady=2)
        self.lwxkey = tk.Label(fwx1, text=self.t("wx_key"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lwxkey.pack(side="left")
        self.vwxkey = tk.StringVar(value=_saved.get("vwxkey", ""))
        tk.Entry(fwx1, textvariable=self.vwxkey, show="*",
                 bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
                 font=FL, relief="flat", bd=4).pack(side="left", fill="x", expand=True)

        fwx_hint = tk.Frame(p, bg=PANEL); fwx_hint.pack(fill="x", padx=20, pady=(0,2))
        self._wx_hint_label = tk.Label(fwx_hint, text="", font=("微软雅黑",8),
                 bg=PANEL, fg="#888888")
        self._wx_hint_label.pack(side="left")

        def _on_wx_type_change(*_):
            wt = self.vwx_type.get()
            if wt == "pushplus":
                self._wx_type_label.config(text=self.t("wx_type_pp"))
                self._wx_hint_label.config(text="   ↑ "+self.t("wx_hint_pp"))
            else:
                self._wx_type_label.config(text=self.t("wx_type_sc"))
                self._wx_hint_label.config(text="   ↑ "+self.t("wx_hint_sc"))
        self.vwx_type.trace_add("write", _on_wx_type_change)
        _on_wx_type_change()  # 初始化显示

        fwx2 = tk.Frame(p, bg=PANEL); fwx2.pack(fill="x", padx=20, pady=2)
        self.vauto_wx = tk.BooleanVar(value=_saved.get("vauto_wx", False))
        ToggleBtn(fwx2, text=self.t("wx_enable"), variable=self.vauto_wx,
                  bg=PANEL, fg=TEXT, font=FS).pack(side="left")

        fwx3 = tk.Frame(p, bg=PANEL); fwx3.pack(fill="x", padx=20, pady=2)
        self.lwxpct = tk.Label(fwx3, text=self.t("wx_threshold"), font=FS, bg=PANEL, fg=TEXT_DIM)
        self.lwxpct.pack(side="left")
        self.vwx_threshold = tk.StringVar(value=_saved.get("vwx_threshold", "10"))
        tk.Entry(fwx3, textvariable=self.vwx_threshold, width=5,
                 bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
                 font=FL, relief="flat", bd=4).pack(side="left", padx=4)
        self.lwxpct2 = tk.Label(fwx3, text=self.t("wx_pct"), font=FS, bg=PANEL, fg=TEXT_DIM)
        self.lwxpct2.pack(side="left")

        fwx4 = tk.Frame(p, bg=PANEL); fwx4.pack(fill="x", padx=20, pady=(4,8))
        self.bwxtest = tk.Button(fwx4, text=self.t("wx_test"), command=self._wx_test,
                  bg="#07C160", fg="white", font=FS,
                  relief="flat", bd=0, padx=12, pady=6, cursor="hand2")
        self.bwxtest.pack(side="left")

        sep3 = tk.Frame(p, bg="#333333", height=1); sep3.pack(fill="x", padx=20, pady=(4,0))

        # ── 排名下滑预警 ──
        frank0 = tk.Frame(p, bg=PANEL); frank0.pack(fill="x", padx=20, pady=(8,2))
        tk.Label(frank0, text="📉 "+self.t("rank_enable"), font=FB, bg=PANEL, fg="#E67E22").pack(side="left")

        frank1 = tk.Frame(p, bg=PANEL); frank1.pack(fill="x", padx=20, pady=2)
        self.vauto_rank = tk.BooleanVar(value=_saved.get("vauto_rank", True))
        ToggleBtn(frank1, text=self.t("rank_enable"), variable=self.vauto_rank,
                  bg=PANEL, fg=TEXT, font=FS).pack(side="left")

        frank2 = tk.Frame(p, bg=PANEL); frank2.pack(fill="x", padx=20, pady=2)
        tk.Label(frank2, text=self.t("rank_drop"), font=FS, bg=PANEL, fg=TEXT_DIM).pack(side="left")
        self.vrank_threshold = tk.StringVar(value=_saved.get("vrank_threshold", "10"))
        tk.Entry(frank2, textvariable=self.vrank_threshold, width=5,
                 bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
                 font=FL, relief="flat", bd=4).pack(side="left", padx=4)
        tk.Label(frank2, text=self.t("rank_pos"), font=FS, bg=PANEL, fg=TEXT_DIM).pack(side="left")

        sep4 = tk.Frame(p, bg="#333333", height=1); sep4.pack(fill="x", padx=20, pady=(8,0))

        # ── 销量暴增监控 ──
        fsales0 = tk.Frame(p, bg=PANEL); fsales0.pack(fill="x", padx=20, pady=(8,2))
        tk.Label(fsales0, text="🚀 "+self.t("sales_title"), font=FB,
                 bg=PANEL, fg="#9B59B6").pack(side="left")

        fsales1 = tk.Frame(p, bg=PANEL); fsales1.pack(fill="x", padx=20, pady=2)
        self.vauto_sales = tk.BooleanVar(value=_saved.get("vauto_sales", True))
        ToggleBtn(fsales1, text=self.t("sales_enable"), variable=self.vauto_sales,
                  bg=PANEL, fg=TEXT, font=FS).pack(side="left")

        fsales2 = tk.Frame(p, bg=PANEL); fsales2.pack(fill="x", padx=20, pady=2)
        tk.Label(fsales2, text=self.t("sales_drop"), font=FS, bg=PANEL, fg=TEXT_DIM).pack(side="left")
        self.vsales_threshold = tk.StringVar(value=_saved.get("vsales_threshold", "50"))
        tk.Entry(fsales2, textvariable=self.vsales_threshold, width=5,
                 bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
                 font=FL, relief="flat", bd=4).pack(side="left", padx=4)
        tk.Label(fsales2, text=self.t("sales_pct"), font=FS, bg=PANEL, fg=TEXT_DIM).pack(side="left")

        sep5 = tk.Frame(p, bg="#333333", height=1); sep5.pack(fill="x", padx=20, pady=(8,0))

        # 自家店铺名
        fown = tk.Frame(p, bg=PANEL); fown.pack(fill="x", padx=20, pady=4)
        self.lown = tk.Label(fown, text=self.t("own_shop"), font=FS, bg=PANEL, fg=TEXT_DIM, width=12, anchor="w")
        self.lown.pack(side="left")
        self.vown = tk.StringVar(value=_saved.get("vown", ""))
        tk.Entry(fown, textvariable=self.vown, bg=INPUT_BG, fg=INPUT_FG, font=FL,
                 insertbackground=INPUT_FG, relief="flat", bd=4).pack(side="left", fill="x", expand=True)

        # 日报按钮
        self.breport = tk.Button(p, text=self.t("btn_report"),
            font=FB, bg=GREEN, fg="white", relief="flat", bd=0,
            padx=16, pady=8, cursor="hand2", command=self._gen_report)
        self.breport.pack(padx=20, pady=(0,20), fill="x")

        # ── 定时自动采集 ──
        sep_sched = tk.Frame(p, bg="#333333", height=1); sep_sched.pack(fill="x", padx=20, pady=(4,0))
        fsched0 = tk.Frame(p, bg=PANEL); fsched0.pack(fill="x", padx=20, pady=(8,2))
        tk.Label(fsched0, text="🕘 "+self.t("sched_title"), font=FB,
                 bg=PANEL, fg="#F39C12").pack(side="left")

        fsched1 = tk.Frame(p, bg=PANEL); fsched1.pack(fill="x", padx=20, pady=2)
        self.vauto_schedule = tk.BooleanVar(value=_saved.get("vauto_schedule", False))
        ToggleBtn(fsched1, text=self.t("sched_enable"), variable=self.vauto_schedule,
                  bg=PANEL, fg=TEXT, font=FS).pack(side="left")

        fsched2 = tk.Frame(p, bg=PANEL); fsched2.pack(fill="x", padx=20, pady=2)
        tk.Label(fsched2, text=self.t("sched_time"), font=FS, bg=PANEL, fg=TEXT_DIM).pack(side="left")
        self.vsched_hour = tk.StringVar(value=_saved.get("vsched_hour", "9"))
        tk.Entry(fsched2, textvariable=self.vsched_hour, width=4,
                 bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
                 font=FL, relief="flat", bd=4).pack(side="left", padx=4)
        tk.Label(fsched2, text=":", font=FS, bg=PANEL, fg=TEXT_DIM).pack(side="left")
        self.vsched_min = tk.StringVar(value=_saved.get("vsched_min", "00"))
        tk.Entry(fsched2, textvariable=self.vsched_min, width=4,
                 bg=INPUT_BG, fg=INPUT_FG, insertbackground=INPUT_FG,
                 font=FL, relief="flat", bd=4).pack(side="left", padx=2)
        tk.Label(fsched2, text=self.t("sched_hint"), font=("Segoe UI",8),
                 bg=PANEL, fg="#888888").pack(side="left", padx=(6,0))

        # ── 多平台采集 ──
        sep6 = tk.Frame(p, bg="#333333", height=1); sep6.pack(fill="x", padx=20, pady=(4,0))
        fmp0 = tk.Frame(p, bg=PANEL); fmp0.pack(fill="x", padx=20, pady=(8,2))
        tk.Label(fmp0, text="🌐 "+self.t("multi_title"), font=FB,
                 bg=PANEL, fg="#16A085").pack(side="left")

        fmp1 = tk.Frame(p, bg=PANEL); fmp1.pack(fill="x", padx=20, pady=2)
        self.vmp_jd  = tk.BooleanVar(value=_saved.get("vmp_jd", True))
        self.vmp_rev = tk.BooleanVar(value=_saved.get("vmp_rev", True))
        ToggleBtn(fmp1, text=self.t("multi_jd"),  variable=self.vmp_jd,
                  bg=PANEL, fg=TEXT, font=FS).pack(side="left", padx=(0,8))

        fmp2 = tk.Frame(p, bg=PANEL); fmp2.pack(fill="x", padx=20, pady=2)
        self.vmp_hw  = tk.BooleanVar(value=_saved.get("vmp_hw", True))
        ToggleBtn(fmp2, text=self.t("multi_review"), variable=self.vmp_rev,
                  bg=PANEL, fg=TEXT, font=FS).pack(side="left", padx=(0,8))
        ToggleBtn(fmp2, text=self.t("multi_hw"), variable=self.vmp_hw,
                  bg=PANEL, fg=TEXT, font=FS).pack(side="left")

        fmp3 = tk.Frame(p, bg=PANEL); fmp3.pack(fill="x", padx=20, pady=(4,12))
        self.bmp = tk.Button(fmp3, text=self.t("multi_btn"), command=self._run_multiplatform,
                  bg="#4A4A4A", fg="#AAAAAA", font=("Arial", 8),
                  relief="flat", bd=0, padx=8, pady=4, cursor="hand2")
        self.bmp.pack(side="left")
        tk.Label(fmp3, text="← 已勾选时淘宝采集后自动运行，无需手动点击",
                 font=("Arial", 7), bg=PANEL, fg="#666666").pack(side="left", padx=6)

    def _build_btns(self):
        bb = tk.Frame(self.root, bg=BG)
        bb.pack(fill="x", padx=16, pady=(0,14))
        self.bstart = tk.Button(bb, text=self.t("btn_start"), font=FB, bg=ACCENT, fg="white",
            relief="flat", bd=0, padx=28, pady=10, cursor="hand2",
            activebackground="#C0392B", activeforeground="white", command=self._start)
        self.bstart.pack(side="left", padx=(0,10))
        self.bstop = tk.Button(bb, text=self.t("btn_stop"), font=FB, bg=BORDER, fg=TEXT_DIM,
            relief="flat", bd=0, padx=20, pady=10, cursor="hand2",
            activebackground="#444", activeforeground=TEXT, state="disabled", command=self._stop)
        self.bstop.pack(side="left")
        self.lstats = tk.Label(bb, text=self.t("collected")+"0", font=FS, bg=BG, fg=TEXT_DIM)
        self.lstats.pack(side="right", padx=8)

    def _switch_lang(self, code):
        self.lang = code
        self.root.title(self.t("title"))
        self.lbl_title.configure(text=self.t("title"))
        self.lbl_sub.configure(text=self.t("subtitle"))
        self.slbl.configure(text=self.t("ready") if not self.is_running else self.t("running"))
        self.lbl_log.configure(text=self.t("log_title"))
        self.btn_clear.configure(text=self.t("clear"))
        self.bstart.configure(text=self.t("btn_start"))
        self.bstop.configure(text=self.t("btn_stop"))
        self.lstats.configure(text=self.t("collected")+str(self.total))
        self._fill_left()

    def _browse(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")], initialfile=self.vout.get())
        if path:
            self.vout.set(path)

    def _set_status(self, key, color):
        self.sdot.configure(fg=color)
        self.slbl.configure(fg=color, text=self.t(key))

    MAX_KEYWORDS = 10  # 关键词上限，避免长时间采集触发风控

    def _kw_add(self):
        kw = self.vkw_input.get().strip()
        if not kw:
            return
        existing = list(self.kw_listbox.get(0, "end"))
        if len(existing) >= self.MAX_KEYWORDS:
            messagebox.showwarning("", self.t("warn_kw_max"))
            return
        if kw not in existing:
            self.kw_listbox.insert("end", kw)
        self.vkw_input.set("")
        # 更新vkw为第一个关键词（兼容旧逻辑）
        if self.kw_listbox.size() > 0:
            self.vkw.set(self.kw_listbox.get(0))

    def _kw_del(self):
        sel = self.kw_listbox.curselection()
        if sel:
            self.kw_listbox.delete(sel[0])
        elif self.kw_listbox.size() > 0:
            self.kw_listbox.delete("end")
        if self.kw_listbox.size() > 0:
            self.vkw.set(self.kw_listbox.get(0))

    def _get_keywords(self):
        """获取关键词列表，至少返回一个"""
        kws = [self.kw_listbox.get(i).strip()
               for i in range(self.kw_listbox.size())]
        kws = [k for k in kws if k]
        return kws if kws else [self.vkw.get().strip()]

    def _start(self):
        if self.is_running:
            return
        keywords = self._get_keywords()
        if not keywords or not keywords[0]:
            messagebox.showwarning("", self.t("warn_kw"))
            return
        kw = keywords[0]
        safe = re.sub(r'[\\/:*?"<>|]', "_", kw)
        if self.vout.get() == OUT_DEFAULT:
            self.vout.set(safe + "_data.xlsx")
        cfg = self._cfg()
        cfg["keywords"] = keywords  # 注入关键词列表
        self.is_running = True
        self.stop_ev.clear()
        self.login_ev.clear()
        self.total = 0
        self.bstart.configure(state="disabled", bg="#555")
        self.bstop.configure(state="normal", bg=ACCENT)
        self._set_status("running", GREEN)
        threading.Thread(target=self._worker, args=(cfg,), daemon=True).start()

    def _stop(self):
        self.stop_ev.set()
        self.login_ev.set()
        self.log.info("Stopped by user")
        self._reset()

    def _reset(self):
        self.is_running = False
        self.bstart.configure(state="normal", bg=ACCENT)
        self.bstop.configure(state="disabled", bg=BORDER)
        self._set_status("ready", TEXT_DIM)

    def _cfg(self):
        zh = LANG["zh"]; en = LANG["en"]; ko = LANG["ko"]
        sort_map = {}
        for d in [zh, en, ko]:
            sort_map[d["sort1"]] = "sale"
            sort_map[d["sort2"]] = "default"
            sort_map[d["sort3"]] = "price_asc"
            sort_map[d["sort4"]] = "price_desc"
        rep_map = {}
        for d in [zh, en, ko]:
            rep_map[d["rep1"]] = 0
            rep_map[d["rep2"]] = 1
            rep_map[d["rep3"]] = 2
            rep_map[d["rep4"]] = 6
        exact_vals = [d["mode_exact"] for d in [zh, en, ko]]
        tmall_vals = [d["shops_tmall"] for d in [zh, en, ko]]
        try: pmin = float(self.vpmin.get() or 0)
        except Exception: pmin = 0
        try: pmax = float(self.vpmax.get() or 0)
        except Exception: pmax = 0
        return {
            "keyword": self.vkw.get().strip(),
            "exact": self.vex.get() in exact_vals,
            "pages": int(self.vpg.get()),
            "sort_by": sort_map.get(self.vso.get(), "sale"),
            "tmall_only": self.vsh.get() in tmall_vals,
            "price_min": pmin,
            "price_max": pmax,
            "interval_hr": rep_map.get(self.viv.get(), 1),
            "output_file": self.vout.get(),
            "login_timeout": 300,
        }

    def _worker(self, cfg):
        self._save_config()  # 每次采集前保存当前配置
        import time
        driver = None
        try:
            for pkg in ["selenium","pandas","openpyxl","webdriver_manager"]:
                __import__(pkg)
            from selenium import webdriver
            from selenium.webdriver.chrome.service import Service
            from selenium.webdriver.chrome.options import Options
            from webdriver_manager.chrome import ChromeDriverManager
            import socket, subprocess, time as _time, os

            # ChromeDriver只获取一次路径，避免每次联网检查
            if not getattr(self, '_chromedriver_path', None):
                self.log.info("正在初始化 ChromeDriver（首次需要下载）...")
                try:
                    self._chromedriver_path = ChromeDriverManager().install()
                except Exception as _e:
                    self.log.error(f"ChromeDriver 获取失败: {_e}")
                    return
            _driver_path = self._chromedriver_path

            def _try_connect(port):
                try:
                    s = socket.socket(); s.settimeout(0.5)
                    result = s.connect_ex(("127.0.0.1", port)); s.close()
                    if result == 0:
                        _opts = Options()
                        _opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
                        _d = webdriver.Chrome(service=Service(_driver_path), options=_opts)
                        return _d, port
                except Exception:
                    pass
                return None, None

            def _find_chrome():
                """查找Chrome可执行文件路径"""
                import winreg
                # 常见安装路径
                candidates = [
                    os.path.join(os.environ.get("PROGRAMFILES", ""), "Google", "Chrome", "Application", "chrome.exe"),
                    os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
                    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
                ]
                for p in candidates:
                    if p and os.path.isfile(p):
                        return p
                # 注册表查找
                try:
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe")
                    val, _ = winreg.QueryValueEx(key, "")
                    winreg.CloseKey(key)
                    if os.path.isfile(val):
                        return val
                except Exception:
                    pass
                return None

            def _launch_chrome_with_debug(port=9222):
                """强制关闭Chrome后以调试端口重新启动，保留用户数据"""
                import urllib.request as _urlreq
                import tempfile

                chrome_path = _find_chrome()
                if not chrome_path:
                    self.log.error("❌ 未找到Chrome，请确认已安装Google Chrome")
                    return False

                self.log.info(f"  Chrome路径: {chrome_path}")

                # ── 第1步：彻底关闭所有Chrome及相关进程 ──
                self.log.info("正在关闭现有Chrome进程...")
                for proc_name in ["chromedriver.exe", "chrome.exe",
                                  "GoogleCrashHandler.exe", "GoogleCrashHandler64.exe"]:
                    subprocess.call(["taskkill", "/f", "/t", "/im", proc_name],
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                _time.sleep(5)

                # 验证Chrome已完全退出
                for _retry in range(10):
                    result = subprocess.run(["tasklist", "/fi", "imagename eq chrome.exe"],
                                            capture_output=True, text=True)
                    if "chrome.exe" not in result.stdout.lower():
                        self.log.info("  Chrome进程已全部退出")
                        break
                    self.log.info(f"  Chrome仍在运行，重试关闭（{_retry+1}/10）...")
                    subprocess.call(["taskkill", "/f", "/t", "/im", "chrome.exe"],
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    _time.sleep(2)
                else:
                    self.log.error("  ⚠️ 无法完全关闭Chrome，继续尝试启动...")

                # ── 第2步：清除锁文件 ──
                # 使用独立目录，避免 Chrome 136+ 默认目录静默忽略 --remote-debugging-port
                user_data = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Debug Profile")
                self.log.info(f"  User Data目录: {user_data}")
                for lock_name in ["SingletonLock", "SingletonSocket", "SingletonCookie"]:
                    lock_path = os.path.join(user_data, lock_name)
                    try:
                        if os.path.exists(lock_path):
                            os.remove(lock_path)
                            self.log.info(f"  已清除锁文件: {lock_name}")
                    except Exception as _e:
                        self.log.warning(f"  清除锁文件失败 {lock_name}: {_e}")

                # 额外等待确保文件锁释放
                _time.sleep(2)

                # ── 第3步：通过临时BAT文件启动Chrome（彻底避免引号问题）──
                bat_content = f'''@echo off
start "" "{chrome_path}" ^
  --remote-debugging-port={port} ^
  --user-data-dir="{user_data}" ^
  --disable-blink-features=AutomationControlled ^
  --no-first-run ^
  --no-default-browser-check ^
  --no-restore-last-session
'''
                bat_path = os.path.join(tempfile.gettempdir(), "_chrome_launch.bat")
                with open(bat_path, "w", encoding="utf-8") as f:
                    f.write(bat_content)

                # 清除脏环境变量，避免继承已死进程的crashpad管道导致初始化异常
                clean_env = {k: v for k, v in os.environ.items()
                             if k != "CHROME_CRASHPAD_PIPE_NAME"}

                self.log.info(f"正在启动Chrome（调试端口 {port}）...")
                self.log.info(f"  通过临时脚本启动: {bat_path}")
                subprocess.Popen(["cmd", "/c", bat_path],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                 env=clean_env)

                # ── 第4步：等待端口就绪（最多45秒）──
                self.log.info(f"  等待调试端口就绪...")
                for i in range(90):
                    _time.sleep(0.5)
                    if i > 0 and i % 20 == 0:
                        self.log.info(f"  仍在等待... ({i//2}秒)")
                    # TCP检查
                    try:
                        s = socket.socket(); s.settimeout(0.5)
                        rc = s.connect_ex(("127.0.0.1", port))
                        s.close()
                        if rc != 0:
                            continue
                    except Exception:
                        continue
                    # HTTP健康检查
                    try:
                        req = _urlreq.Request(f"http://127.0.0.1:{port}/json/version")
                        resp = _urlreq.urlopen(req, timeout=2)
                        if resp.status == 200:
                            self.log.info(f"✅ Chrome已启动，调试端口 {port} 就绪")
                            # 清理临时文件
                            try: os.remove(bat_path)
                            except Exception: pass
                            return True
                    except Exception:
                        pass

                # ── 第5步：如果失败，尝试用独立的user-data-dir启动（放弃保留登录态）──
                self.log.warning("  使用默认User Data启动失败，尝试用独立目录启动...")
                # 先杀掉刚才启动的Chrome
                subprocess.call(["taskkill", "/f", "/t", "/im", "chrome.exe"],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                _time.sleep(3)

                alt_data = os.path.join(tempfile.gettempdir(), "chrome_scraper_profile")
                bat_content2 = f'''@echo off
start "" "{chrome_path}" ^
  --remote-debugging-port={port} ^
  --user-data-dir="{alt_data}" ^
  --disable-blink-features=AutomationControlled ^
  --no-first-run ^
  --no-default-browser-check
'''
                with open(bat_path, "w", encoding="utf-8") as f:
                    f.write(bat_content2)

                self.log.info(f"  使用独立目录: {alt_data}")
                subprocess.Popen(["cmd", "/c", bat_path],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                 env=clean_env)

                for i in range(40):
                    _time.sleep(0.5)
                    try:
                        req = _urlreq.Request(f"http://127.0.0.1:{port}/json/version")
                        resp = _urlreq.urlopen(req, timeout=2)
                        if resp.status == 200:
                            self.log.info(f"✅ Chrome已用独立目录启动，端口 {port} 就绪")
                            self.log.warning("  ⚠️ 使用了独立目录，需要重新登录淘宝")
                            try: os.remove(bat_path)
                            except Exception: pass
                            return True
                    except Exception:
                        pass

                # 最终诊断
                self.log.error("❌ 两种方式都无法启动带调试端口的Chrome")
                result = subprocess.run(["tasklist", "/fi", "imagename eq chrome.exe"],
                                        capture_output=True, text=True)
                if "chrome.exe" in result.stdout.lower():
                    self.log.error("  Chrome进程存在但调试端口未开启")
                    self.log.error("  请尝试：重启电脑后再运行")
                else:
                    self.log.error("  Chrome未能启动")
                    self.log.error("  请确认Chrome已正确安装")
                try: os.remove(bat_path)
                except Exception: pass
                return False

            # ── 连接Chrome ──
            driver = None
            connected_port = None

            # 第1步：尝试连接已有的调试端口
            for _port in [9222, 9223, 9224, 9225]:
                driver, connected_port = _try_connect(_port)
                if driver:
                    self.log.info(f"✅ 已连接到Chrome（端口 {connected_port}）")
                    break

            # 第2步：连接失败，自动启动带调试端口的Chrome
            if not driver:
                self.log.info("未检测到调试端口，正在自动启动Chrome...")
                if _launch_chrome_with_debug(9222):
                    _time.sleep(2)
                    driver, connected_port = _try_connect(9222)
                    if driver:
                        self.log.info("✅ Chrome已启动并连接成功")

            if not driver:
                self.log.error("❌ 无法连接Chrome")
                self.root.after(0, lambda: messagebox.showwarning(
                    "连接失败",
                    "无法启动或连接Chrome。\n\n请确认：\n1. 已安装Google Chrome\n2. 手动关闭所有Chrome窗口后重试"
                ))
                return

            # 连接后把GUI置回前台
            try:
                import ctypes
                hwnd = ctypes.windll.user32.FindWindowW(None, self.root.title())
                if hwnd:
                    ctypes.windll.user32.SetForegroundWindow(hwnd)
            except Exception:
                pass

            if not self._login_dialog():
                return


            rn = 0
            while not self.stop_ev.is_set():
                rn += 1
                keywords = cfg.get("keywords") or [cfg["keyword"]]

                # ══ 多关键词循环 ══
                for kw_idx, kw in enumerate(keywords):
                    if self.stop_ev.is_set():
                        break

                    # 每个关键词独立的output文件
                    # 单关键词时直接用用户设置的文件名；多关键词时以关键词命名，
                    # 放在用户设置的output_file同目录下
                    safe_kw = re.sub(r'[\\/:*?"<>|]', "_", kw)
                    if len(keywords) == 1:
                        kw_out = cfg.get("output_file") or (safe_kw + "_data.xlsx")
                    else:
                        base_dir = os.path.dirname(os.path.abspath(
                            cfg.get("output_file") or (safe_kw + "_data.xlsx")))
                        kw_out = os.path.join(base_dir, safe_kw + "_data.xlsx")
                    kw_cfg = dict(cfg)
                    kw_cfg["keyword"]     = kw
                    kw_cfg["output_file"] = kw_out

                    if len(keywords) > 1:
                        self.log.info(f"\n{'='*50}")
                        self.log.info(f"  [{kw_idx+1}/{len(keywords)}] 关键词: {kw}")
                        self.log.info(f"{'='*50}")

                    from scraper_core import run_round
                    cnt = run_round(driver, kw_cfg, rn, self.log, self.stop_ev)
                    self.total += cnt
                    self.last_file = kw_out
                    total_now = self.total
                    self.root.after(0, lambda c=total_now:
                        self.lstats.configure(text=self.t("collected")+str(c)))

                    # ── 淘宝完成后，如勾选了京东/评价/热词，自动触发多平台采集 ──
                    if not self.stop_ev.is_set() and (
                            self.vmp_jd.get() or self.vmp_rev.get() or self.vmp_hw.get()):
                        self.log.info("[自动多平台] 淘宝采集完成，开始自动多平台采集...")
                        try:
                            from scraper_core import run_multiplatform
                            platforms = []
                            if self.vmp_jd.get(): platforms.append("jd")

                            jd_handle = None
                            for h in driver.window_handles:
                                try:
                                    driver.switch_to.window(h)
                                    t = driver.title; u = driver.current_url.lower()
                                    if jd_handle is None and ("jd.com" in u or "360buy" in u or "京东" in t):
                                        jd_handle = h
                                except Exception: pass

                            mp_cfg = {
                                "keyword":         kw,
                                "pages":           kw_cfg["pages"],
                                "output_file":     kw_out,
                                "platforms":       platforms,
                                "scrape_reviews":  self.vmp_rev.get(),
                                "scrape_hotwords": self.vmp_hw.get(),
                                "jd_handle":       jd_handle,
                            }
                            jd, _pdd, rev = run_multiplatform(
                                driver, mp_cfg, self.log, self.stop_ev)
                            jd_cnt = len(jd)
                        except Exception as _e:
                            self.log.error(f"[自动多平台] 出错: {_e}")
                            jd_cnt = 0

                        parts = [f"淘宝 {cnt} 条"]
                        if self.vmp_jd.get(): parts.append(f"京东 {jd_cnt} 条")
                        self.log.info(f"\n  ✅ [{kw}] 本轮完成，共采集 {' / '.join(parts)}")
                    else:
                        self.log.info(f"\n  ✅ [{kw}] 本轮完成，共采集 淘宝 {cnt} 条")

                    # 多关键词之间稍作间隔，避免触发风控
                    if len(keywords) > 1 and kw_idx < len(keywords)-1:
                        if not self.stop_ev.is_set():
                            self.log.info(f"  ⏳ 切换到下一关键词，等待5秒...")
                            time.sleep(5)

                if len(keywords) > 1 and not self.stop_ev.is_set():
                    self.log.info(f"\n  🎉 本轮全部关键词采集完成（共{len(keywords)}个）")

                # 自动发送邮件+日报
                fp  = self.last_file
                own = self.vown.get().strip()
                kw  = self.vkw.get().strip()

                # 微信简报推送：只要配了Key，不管有没有邮件都推
                if fp and os.path.exists(fp) and self.vauto_wx.get() and self.vwxkey.get().strip():
                    threading.Thread(target=self._push_brief_wx,
                        args=(fp, own, kw), daemon=True).start()

                # 邮件日报：需要邮件配置
                if self.vauto.get():
                    sender = self.vefrom.get().strip()
                    pwd    = self.vepwd.get().strip()
                    to     = self.veto.get().strip()
                    if sender and pwd and to and fp and os.path.exists(fp):
                        self.log.info("Auto-sending email...")
                        threading.Thread(target=self._report_thread,
                            args=(sender, pwd, to, fp, own, kw), daemon=True).start()
                    else:
                        self.log.warning("Auto-send skipped: email info incomplete")
                # 降价预警检测
                if self.last_file and os.path.exists(self.last_file):
                    kw = self.vkw.get().strip()
                    fp = self.last_file
                    if self.vauto_wx.get():
                        threading.Thread(target=self._check_price_alerts,
                            args=(fp, kw), daemon=True).start()
                    # 排名下滑检测
                    threading.Thread(target=self._check_rank_alerts,
                        args=(fp, kw), daemon=True).start()
                    # 销量暴增检测
                    threading.Thread(target=self._check_sales_spike,
                        args=(fp, kw), daemon=True).start()
                if cfg["interval_hr"] <= 0 or self.stop_ev.is_set():
                    break
                for _ in range(cfg["interval_hr"] * 3600):
                    if self.stop_ev.is_set():
                        break
                    time.sleep(1)
        except Exception as e:
            tb = traceback.format_exc()
            self.log.error("Error: " + str(e))
            self.log.error(tb)
            traceback.print_exc()
            err_msg = str(e) + "\n\n" + tb[-500:]
            self.root.after(0, lambda m=err_msg: messagebox.showerror("Error", m))
        finally:
            # 采集结束，恢复Chrome窗口到可见位置
            try:
                from scraper_core import restore_chrome_window
                restore_chrome_window(driver)
            except Exception:
                pass
            try:
                driver.quit()
            except Exception:
                pass
            self.root.after(0, self._reset)
            self.log.info("Done.")

    def _login_dialog(self):
        # 全自动：连接Chrome成功后直接开始采集，无需手动确认
        self.log.info("Chrome connected, starting scrape automatically...")
        return True

    def _send_email(self):
        sender = self.vefrom.get().strip()
        pwd    = self.vepwd.get().strip()
        to     = self.veto.get().strip()
        fp     = self.last_file or self.vout.get().strip()
        if not sender or not pwd or not to:
            messagebox.showwarning("", self.t("warn_email"))
            return
        if not fp or not os.path.exists(fp):
            messagebox.showwarning("", self.t("warn_file"))
            return
        self.bsend.configure(state="disabled", text=self.t("sending"))
        threading.Thread(target=self._send_thread, args=(sender, pwd, to, fp), daemon=True).start()

    def _send_thread(self, sender, pwd, to, fp):
        try:
            self.log.info("Sending email to " + to + "...")
            msg = MIMEMultipart()
            msg["From"] = sender
            msg["To"] = to
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            msg["Subject"] = "Taobao Data - " + os.path.basename(fp) + " - " + now_str
            body = "Data file attached.\nFile: " + os.path.basename(fp) + "\nSent: " + str(datetime.now())
            msg.attach(MIMEText(body, "plain", "utf-8"))
            with open(fp, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                fname = os.path.basename(fp)
                from email.header import Header
                part.add_header("Content-Disposition", "attachment",
                                filename=("utf-8", "", fname))
                msg.attach(part)
            host, port, use_ssl = get_smtp_config(sender)
            self.log.info(f"  SMTP: {host}:{port} (SSL={use_ssl})")
            with smtp_connect(host, port, use_ssl) as srv:
                srv.login(sender, pwd)
                srv.sendmail(sender, [to], msg.as_string())
            self.log.info("Email sent to " + to)
            self.root.after(0, lambda: messagebox.showinfo("", self.t("email_ok")))
        except Exception as e:
            err = str(e)
            self.log.error("Email error: " + err)
            self.root.after(0, lambda m=err: messagebox.showerror("", self.t("email_fail") + m))
        finally:
            self.root.after(0, lambda: self.bsend.configure(state="normal", text=self.t("btn_send")))


    def _run_multiplatform(self):
        """触发多平台采集"""
        if self.is_running:
            messagebox.showwarning("", "请先停止当前采集")
            return
        if not self.vmp_jd.get() and not self.vmp_rev.get() and not self.vmp_hw.get():
            messagebox.showwarning("", "请至少勾选一个平台或评价分析")
            return
        fp = self.vout.get().strip()
        if not fp or not os.path.exists(fp):
            messagebox.showwarning("", "请先完成淘宝采集，再进行多平台对比")
            return

        self.bmp.configure(state="disabled")
        self.stop_ev.clear()   # 重置停止标志，防止上次停止后残留
        self.log.info("[多平台] 开始多平台采集...")

        def _worker():
            try:
                from selenium import webdriver
                from selenium.webdriver.chrome.service import Service
                from selenium.webdriver.chrome.options import Options
                from webdriver_manager.chrome import ChromeDriverManager
                from scraper_core import run_multiplatform

                import socket as _sock
                _dp = self._chromedriver_path or ChromeDriverManager().install()
                driver = None
                for _p in [9222, 9223, 9224, 9225]:
                    try:
                        _s = _sock.socket(); _s.settimeout(0.5)
                        if _s.connect_ex(("127.0.0.1", _p)) == 0:
                            _s.close()
                            opts = Options()
                            opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{_p}")
                            driver = webdriver.Chrome(service=Service(_dp), options=opts)
                            self.log.info(f"[多平台] 已连接Chrome（端口 {_p}）")
                            break
                        _s.close()
                    except Exception:
                        pass
                if not driver:
                    self.log.error("[多平台] 无法连接Chrome，请先通过开始采集启动Chrome")
                    self.root.after(0, lambda: self.bmp.configure(state="normal"))
                    return

                import threading as _threading
                import time as _time

                platforms = []
                if self.vmp_jd.get():  platforms.append("jd")

                # ── 扫描所有标签页，按标题/URL匹配京东 ──
                jd_handle  = None
                self.log.info("[多平台] 扫描浏览器标签页...")
                for h in driver.window_handles:
                    try:
                        driver.switch_to.window(h)
                        t = driver.title
                        u = driver.current_url.lower()
                        self.log.info(f"  标签: {t[:35]} | {u[:50]}")
                        if jd_handle is None and ("jd.com" in u or "360buy" in u or "京东" in t):
                            jd_handle = h
                            self.log.info(f"  → 识别为京东")
                    except Exception: pass

                if "jd" in platforms and not jd_handle:
                    self.log.warning("[多平台] 未找到京东标签页，请在浏览器打开京东后重试")

                cfg = {
                    "keyword":        self.vkw.get().strip(),
                    "pages":          int(self.vpg.get()),
                    "output_file":    fp,
                    "platforms":      platforms,
                    "scrape_reviews":  self.vmp_rev.get(),
                    "scrape_hotwords": self.vmp_hw.get(),
                    "jd_handle":       jd_handle,
                }
                # 登录等待callback：弹窗让用户在浏览器登录后点确定
                def _login_wait(platform_name):
                    import threading as _t
                    done = _t.Event()
                    def _ask():
                        from tkinter import messagebox as _mb
                        _mb.showinfo(
                            f"{platform_name} 需要登录",
                            f"检测到 {platform_name} 需要登录。\n\n请在浏览器中完成登录，\n登录完成后点【确定】继续采集。"
                        )
                        done.set()
                    self.root.after(0, _ask)
                    done.wait(timeout=300)

                jd, _pdd, rev = run_multiplatform(driver, cfg, self.log, self.stop_ev,
                                                   wait_login_callback=_login_wait)
                self.log.info(f"[多平台] 完成！京东 {len(jd)} 条")
                if rev:
                    self.log.info(f"[多平台] 评价分析完成，共 {rev.get('total',0)} 条评价")
                hw = cfg.get("_hotwords", {})
                hw_tb = len(hw.get("淘宝", [])); hw_jd = len(hw.get("京东", []))
                hw_msg = f"\n• 热门关键词（淘宝{hw_tb}个 / 京东{hw_jd}个）" if cfg.get("scrape_hotwords") else ""
                msg = f"多平台采集完成！\n\n京东：{len(jd)} 条\n\n数据已写入同一Excel文件的新sheet：\n• 京东数据\n• 跨平台对比{hw_msg}"
                self.root.after(0, lambda m=msg: messagebox.showinfo("", m))
            except Exception as e:
                tb = traceback.format_exc()
                self.log.error(f"[多平台] 出错: {e}")
                self.root.after(0, lambda m=str(e): messagebox.showerror("", f"多平台采集失败：{m}"))
            finally:
                self.root.after(0, lambda: self.bmp.configure(state="normal"))

        threading.Thread(target=_worker, daemon=True).start()

    def _gen_report(self):
        sender = self.vefrom.get().strip()
        pwd    = self.vepwd.get().strip()
        to     = self.veto.get().strip()
        fp     = self.last_file or self.vout.get().strip()
        own    = self.vown.get().strip()
        kw     = self.vkw.get().strip()
        if not sender or not pwd or not to:
            messagebox.showwarning("", self.t("warn_email")); return
        if not fp or not os.path.exists(fp):
            messagebox.showwarning("", self.t("warn_file")); return
        self.breport.configure(state="disabled")
        threading.Thread(target=self._report_thread,
            args=(sender, pwd, to, fp, own, kw), daemon=True).start()

    # ══════════════════════════════════════════════
    #  微信预警推送（Server酱 / PushPlus）
    # ══════════════════════════════════════════════

    def _wx_send(self, title, content):
        """推送微信消息（自动根据选择使用Server酱或PushPlus），返回是否成功"""
        import urllib.request, urllib.parse, json
        key = self.vwxkey.get().strip()
        if not key:
            return False
        wx_type = self.vwx_type.get() if hasattr(self, 'vwx_type') else "server_chan"
        try:
            if wx_type == "pushplus":
                # PushPlus 推送
                url = "https://www.pushplus.plus/send"
                payload = json.dumps({
                    "token": key,
                    "title": title,
                    "content": content,
                    "template": "markdown",
                }).encode("utf-8")
                req = urllib.request.Request(url, data=payload, method="POST")
                req.add_header("Content-Type", "application/json")
                with urllib.request.urlopen(req, timeout=10) as r:
                    resp = json.loads(r.read().decode("utf-8"))
                    return resp.get("code") == 200
            else:
                # Server酱 推送
                url = f"https://sctapi.ftqq.com/{key}.send"
                data = urllib.parse.urlencode({
                    "title": title,
                    "desp":  content,
                }).encode("utf-8")
                req = urllib.request.Request(url, data=data, method="POST")
                req.add_header("Content-Type", "application/x-www-form-urlencoded")
                with urllib.request.urlopen(req, timeout=10) as r:
                    resp = json.loads(r.read().decode("utf-8"))
                    return resp.get("data", {}).get("errno", -1) == 0 or resp.get("code") == 0
        except Exception as e:
            self.log.warning(f"微信推送失败: {e}")
            return False

    def _wx_test(self):
        """测试Server酱推送"""
        key = self.vwxkey.get().strip()
        if not key:
            messagebox.showwarning("", self.t("wx_no_key"))
            return
        def do_test():
            ok = self._wx_send(
                "✅ 淘宝采集工具 - 测试消息",
                "**推送配置成功！**\n\n当竞品降价超过设定阈值时，你会收到类似这样的预警消息。"
            )
            if ok:
                self.root.after(0, lambda: messagebox.showinfo("", self.t("wx_test_ok")))
            else:
                self.root.after(0, lambda: messagebox.showerror("", self.t("wx_test_fail") + "请检查Key是否正确"))
        threading.Thread(target=do_test, daemon=True).start()

    def _check_price_alerts(self, data_file, keyword):
        """采集完成后检查是否有竞品大幅降价，有则推送微信"""
        if not self.vauto_wx.get():
            return
        key = self.vwxkey.get().strip()
        if not key:
            return
        try:
            import pandas as pd, datetime
            threshold = float(self.vwx_threshold.get().strip() or "10")
            df = pd.read_excel(data_file, sheet_name="采集数据")
            if df.empty or "采集时间" not in df.columns:
                return

            df["_date"] = pd.to_datetime(df["采集时间"], errors="coerce").dt.date
            dates = sorted(df["_date"].dropna().unique())
            if len(dates) < 2:
                self.log.info("[预警] 数据不足2天，跳过降价检测")
                return

            today     = dates[-1]
            yesterday = dates[-2]
            price_col = "当前售价(元)" if "当前售价(元)" in df.columns else "当前售价"
            title_col = "商品标题"     if "商品标题"     in df.columns else "标题"
            shop_col  = "店铺名称"     if "店铺名称"     in df.columns else "店铺"

            df_t = df[df["_date"] == today].copy()
            df_y = df[df["_date"] == yesterday].copy()
            df_t[price_col] = pd.to_numeric(df_t[price_col], errors="coerce")
            df_y[price_col] = pd.to_numeric(df_y[price_col], errors="coerce")

            merged = df_t.merge(
                df_y[[title_col, price_col]].rename(columns={price_col: "_old_price"}),
                on=title_col, how="inner"
            )
            merged["_drop_pct"] = (merged["_old_price"] - merged[price_col]) / merged["_old_price"].replace(0, 1) * 100
            alerts = merged[merged["_drop_pct"] >= threshold].nlargest(5, "_drop_pct")

            if alerts.empty:
                self.log.info(f"[预警] 无竞品降价超过 {threshold}%")
                return

            lines = [f"## 🚨 竞品降价预警 - {keyword}\n"]
            lines.append(f"发现 **{len(alerts)}** 个竞品降价超过 {threshold}%：\n")
            for _, r in alerts.iterrows():
                shop  = str(r.get(shop_col, ""))[:15]
                title = str(r[title_col])[:25]
                old_p = r["_old_price"]
                new_p = r[price_col]
                drop  = r["_drop_pct"]
                lines.append(f"- **{title}**")
                lines.append(f"  {shop} | ¥{old_p:.0f} → ¥{new_p:.0f} | ↓{drop:.1f}%\n")

            lines.append(f"\n> 采集时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")
            msg_body = "\n".join(lines)

            ok = self._wx_send(f"🚨 竞品降价预警【{keyword}】{len(alerts)}条", msg_body)
            if ok:
                self.log.info(f"[预警] 微信推送成功，{len(alerts)} 条降价提醒")
            else:
                self.log.warning("[预警] 微信推送失败，请检查Server酱Key")

        except Exception as e:
            self.log.warning(f"[预警] 降价检测出错: {e}")

    # ══════════════════════════════════════════════
    #  竞品销量暴增监控
    # ══════════════════════════════════════════════

    def _check_sales_spike(self, data_file, keyword):
        """检测竞品销量暴增，推送微信预警"""
        if not self.vauto_sales.get():
            return
        key = self.vwxkey.get().strip()
        if not key:
            return
        try:
            import pandas as pd, datetime
            threshold = float(self.vsales_threshold.get().strip() or "50")
            own_shop  = self.vown.get().strip()

            df = pd.read_excel(data_file, sheet_name="采集数据")
            if df.empty or "采集时间" not in df.columns:
                return

            df["_date"] = pd.to_datetime(df["采集时间"], errors="coerce").dt.date
            dates = sorted(df["_date"].dropna().unique())
            if len(dates) < 2:
                self.log.info("[销量] 数据不足2天，跳过销量检测")
                return

            today     = dates[-1]
            yesterday = dates[-2]
            price_col = "当前售价(元)" if "当前售价(元)" in df.columns else "当前售价"
            sales_col = "月销量(件)"   if "月销量(件)"   in df.columns else "月销量"
            title_col = "商品标题"     if "商品标题"     in df.columns else "标题"
            shop_col  = "店铺名称"     if "店铺名称"     in df.columns else "店铺"

            df_t = df[df["_date"] == today].reset_index(drop=True).copy()
            df_y = df[df["_date"] == yesterday].copy()
            df_t["_rank"]  = df_t.index + 1
            df_t[sales_col] = pd.to_numeric(df_t[sales_col], errors="coerce").fillna(0)
            df_y[sales_col] = pd.to_numeric(df_y[sales_col], errors="coerce").fillna(0)
            df_t[price_col] = pd.to_numeric(df_t[price_col], errors="coerce").fillna(0)

            # 排除自家产品，只看竞品
            df_t_comp = df_t[~df_t[shop_col].astype(str).str.contains(own_shop, na=False)]

            merged = df_t_comp.merge(
                df_y[[title_col, sales_col]].rename(columns={sales_col: "_old_sales"}),
                on=title_col, how="inner"
            )
            merged = merged[merged["_old_sales"] > 0]  # 过滤昨天0销量的
            merged["_growth_pct"] = (merged[sales_col] - merged["_old_sales"]) / merged["_old_sales"] * 100
            spikes = merged[merged["_growth_pct"] >= threshold].nlargest(5, "_growth_pct")

            if spikes.empty:
                self.log.info(f"[销量] 无竞品销量增长超过 {threshold}%")
                return

            self.log.warning(f"[销量] 发现 {len(spikes)} 个竞品销量暴增！")

            # 构建微信推送内容
            lines = [f"## 🚀 竞品销量暴增预警 - {keyword}\n"]
            lines.append(f"发现 **{len(spikes)}** 个竞品销量增长超过 {threshold:.0f}%：\n")

            for _, r in spikes.iterrows():
                title      = str(r[title_col])[:25]
                shop       = str(r.get(shop_col, ""))[:15]
                old_sales  = int(r["_old_sales"])
                new_sales  = int(r[sales_col])
                growth     = r["_growth_pct"]
                rank       = int(r["_rank"])
                price      = float(r[price_col])
                analysis   = _analyse_sales_spike(title, new_sales, old_sales, rank, price, growth)

                lines.append(f"**{title}**")
                lines.append(f"  {shop} | 排名#{rank} | ¥{price:.0f}")
                lines.append(f"  销量：{old_sales:,} → {new_sales:,}（↑{growth:.0f}%）")
                lines.append(f"  📊 {analysis}\n")

            lines.append(f"> {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")

            ok = self._wx_send(
                f"🚀 销量暴增预警【{keyword}】{len(spikes)}条",
                "\n".join(lines)
            )
            if ok:
                self.log.info(f"[销量] 微信推送成功，{len(spikes)} 条暴增提醒")
            else:
                self.log.warning("[销量] 微信推送失败")

        except Exception as e:
            self.log.warning(f"[销量] 检测出错: {e}")
            import traceback; traceback.print_exc()

    # ══════════════════════════════════════════════
    #  排名下滑预警 + 调价建议
    # ══════════════════════════════════════════════

    def _check_rank_alerts(self, data_file, keyword):
        """检测自家产品排名下滑，生成调价建议并推送"""
        if not self.vauto_rank.get():
            return
        try:
            import pandas as pd, datetime

            threshold = int(self.vrank_threshold.get().strip() or "10")
            own_shop  = self.vown.get().strip()
            if not own_shop:
                return

            df = pd.read_excel(data_file, sheet_name="采集数据")
            if df.empty or "采集时间" not in df.columns:
                return

            df["_date"] = pd.to_datetime(df["采集时间"], errors="coerce").dt.date
            dates = sorted(df["_date"].dropna().unique())
            if len(dates) < 2:
                self.log.info("[排名] 数据不足2天，跳过排名检测")
                return

            today     = dates[-1]
            yesterday = dates[-2]
            price_col = "当前售价(元)" if "当前售价(元)" in df.columns else "当前售价"
            sales_col = "月销量(件)"   if "月销量(件)"   in df.columns else "月销量"
            title_col = "商品标题"     if "商品标题"     in df.columns else "标题"
            shop_col  = "店铺名称"     if "店铺名称"     in df.columns else "店铺"

            df_t = df[df["_date"] == today].reset_index(drop=True)
            df_y = df[df["_date"] == yesterday].reset_index(drop=True)
            df_t["_rank"] = df_t.index + 1
            df_y["_rank"] = df_y.index + 1
            df_t[price_col] = pd.to_numeric(df_t[price_col], errors="coerce").fillna(0)
            df_y[price_col] = pd.to_numeric(df_y[price_col], errors="coerce").fillna(0)
            df_t[sales_col] = pd.to_numeric(df_t[sales_col], errors="coerce").fillna(0)

            # 找自家产品
            own_today = df_t[df_t[shop_col].astype(str).str.contains(own_shop, na=False)]
            own_yest  = df_y[df_y[shop_col].astype(str).str.contains(own_shop, na=False)]
            if own_today.empty:
                return

            drop_alerts = []
            for _, row in own_today.iterrows():
                match = own_yest[own_yest[title_col] == row[title_col]]
                if match.empty:
                    continue
                rank_today = int(row["_rank"])
                rank_yest  = int(match["_rank"].values[0])
                drop = rank_today - rank_yest
                if drop >= threshold:
                    # 分析周边竞品，找出可能原因
                    nearby = df_t[
                        (df_t["_rank"] >= max(1, rank_today-5)) &
                        (df_t["_rank"] <= rank_today+5) &
                        (~df_t[shop_col].astype(str).str.contains(own_shop, na=False))
                    ]
                    avg_comp_price = nearby[price_col].mean() if not nearby.empty else 0
                    own_price      = float(row[price_col])
                    own_sales      = int(row[sales_col])

                    # 生成调价建议
                    suggestion = _make_price_suggestion(
                        own_price, avg_comp_price, rank_today, rank_yest, own_sales, nearby
                    )
                    drop_alerts.append({
                        "title":      str(row[title_col])[:30],
                        "rank_now":   rank_today,
                        "rank_yest":  rank_yest,
                        "drop":       drop,
                        "price":      own_price,
                        "avg_comp":   avg_comp_price,
                        "suggestion": suggestion,
                    })

            if not drop_alerts:
                self.log.info(f"[排名] 自家产品排名稳定，无下滑超{threshold}位")
                return

            self.log.warning(f"[排名] 发现 {len(drop_alerts)} 个产品排名下滑超{threshold}位！")

            # ── 微信推送（简要版）──
            wx_lines = [f"## 📉 排名下滑预警 - {own_shop}\n"]
            for a in drop_alerts:
                wx_lines.append(f"**{a['title']}**")
                wx_lines.append(f"排名：#{a['rank_yest']} → #{a['rank_now']} （↓{a['drop']}位）")
                wx_lines.append(f"建议：{a['suggestion']['short']}\n")
            wx_lines.append(f"> {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")
            self._wx_send(
                f"📉 排名预警【{own_shop}】下滑{drop_alerts[0]['drop']}位",
                "\n".join(wx_lines)
            )

            # 仅微信推送，不单独发邮件（排名信息已包含在日报邮件里）

        except Exception as e:
            self.log.warning(f"[排名] 检测出错: {e}")
            import traceback; traceback.print_exc()

    def _push_brief_wx(self, fp, own_shop, keyword):
        """仅推送市场简报到微信，不依赖邮件配置"""
        try:
            _, _, market_brief = _generate_daily_report(fp, own_shop, keyword, self.log)
            if not market_brief:
                self.log.warning("[简报] 数据不足，跳过微信推送")
                return
            ok = self._wx_send(f"📊 {keyword} 市场日报", market_brief)
            if ok:
                self.log.info("[简报] 市场简报已推送微信 ✅")
            else:
                self.log.warning("[简报] 微信推送失败，请检查Server酱Key")
        except Exception as e:
            self.log.error(f"[简报] 推送出错: {e}")

    def _report_thread(self, sender, pwd, to, fp, own_shop, keyword):
        try:
            report_html, report_excel, market_brief = _generate_daily_report(fp, own_shop, keyword, self.log)
            if not report_html:
                self.log.warning(self.t("no_data")); return

            # 微信推送（手动点生成日报时才推，自动采集时由_push_brief_wx处理）
            if market_brief and self.vauto_wx.get() and not getattr(self, '_auto_wx_sent', False):
                ok = self._wx_send(f"📊 {keyword} 市场日报", market_brief)
                if ok:
                    self.log.info("[简报] 市场简报已推送微信 ✅")
                else:
                    self.log.info("[简报] 微信未配置，简报仅发邮件")

            # 发邮件：正文=HTML日报，附件=Excel+日报Excel
            msg = MIMEMultipart("mixed")
            msg["From"] = sender
            msg["To"]   = to
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            msg["Subject"] = f"📊 淘宝日报 {keyword} - {now_str}"

            # HTML正文
            msg.attach(MIMEText(report_html, "html", "utf-8"))

            # 附件1: 原始数据Excel
            for attach_fp in [fp, report_excel]:
                if attach_fp and os.path.exists(attach_fp):
                    with open(attach_fp, "rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", "attachment",
                                        filename=("utf-8", "", os.path.basename(attach_fp)))
                        msg.attach(part)

            host, port, use_ssl = get_smtp_config(sender)
            self.log.info(f"  SMTP: {host}:{port} (SSL={use_ssl})")
            with smtp_connect(host, port, use_ssl) as srv:
                srv.login(sender, pwd)
                srv.sendmail(sender, [to], msg.as_string())
            self.log.info("日报已发送至 " + to)
            self.root.after(0, lambda: messagebox.showinfo("", self.t("report_ok")))
        except Exception as e:
            err = str(e)
            self.log.error("日报错误: " + err)
            self.root.after(0, lambda m=err: messagebox.showerror("", self.t("report_fail") + m))
        finally:
            self.root.after(0, lambda: self.breport.configure(state="normal"))


def _analyse_sales_spike(title, sales_today, sales_yest, rank, price, growth_pct):
    """分析销量暴增的可能原因，返回分析文字"""
    reasons = []
    if growth_pct >= 200:
        reasons.append("增幅极大，疑似大促/秒杀活动")
    elif growth_pct >= 100:
        reasons.append("增幅显著，可能在做限时活动或直播带货")
    else:
        reasons.append("稳定增长，市场需求上升或关键词权重提升")

    if rank <= 10:
        reasons.append(f"排名靠前（#{rank}），自然流量强")
    if price > 0:
        reasons.append(f"当前售价 ¥{price:.0f}")

    suggestion = "建议：关注其促销策略，参考定价和活动形式。"
    if growth_pct >= 100:
        suggestion = "建议：重点分析其活动方式，考虑跟进类似促销。"

    return "，".join(reasons) + "。" + suggestion


def _make_price_suggestion(own_price, avg_comp_price, rank_now, rank_yest, own_sales, nearby_df):
    """根据价格和排名数据生成具体调价建议"""
    suggestions = []
    short = ""

    if avg_comp_price > 0 and own_price > 0:
        diff_pct = (own_price - avg_comp_price) / avg_comp_price * 100

        if diff_pct > 15:
            # 自家价格明显偏高
            target = round(avg_comp_price * 1.05, 0)  # 建议降到竞品均价+5%
            cut = own_price - target
            short  = f"建议降价 ¥{cut:.0f}，调至 ¥{target:.0f}（当前高于均价{diff_pct:.0f}%）"
            detail = (f"当前售价 ¥{own_price:.0f} 高于周边竞品均价 ¥{avg_comp_price:.0f} 约 {diff_pct:.0f}%，"
                      f"建议调整至 ¥{target:.0f} 左右，仍保持小幅溢价但不影响竞争力。")
        elif diff_pct > 5:
            target = round(avg_comp_price * 1.0, 0)
            cut = own_price - target
            short  = f"可考虑小幅降价 ¥{cut:.0f}，与竞品持平"
            detail = (f"当前售价 ¥{own_price:.0f} 略高于均价 ¥{avg_comp_price:.0f}，"
                      f"建议降至 ¥{target:.0f} 与竞品持平，提升价格竞争力。")
        elif diff_pct < -15:
            # 自家价格偏低，排名还下滑，说明价格不是问题
            short  = "价格低于均价但排名仍下滑，建议检查主图/标题优化"
            detail = (f"当前售价 ¥{own_price:.0f} 已低于均价 ¥{avg_comp_price:.0f}，"
                      f"价格非主因，建议优化商品主图、标题关键词或增加促销标签。")
        else:
            short  = "价格居中，建议增加促销标签或限时活动拉升转化"
            detail = (f"售价 ¥{own_price:.0f} 与均价 ¥{avg_comp_price:.0f} 接近，"
                      f"价格差异不大，建议通过促销标签、满减活动或优化评价提升排名。")
    else:
        short  = "暂无足够竞品数据，建议手动核查"
        detail = "周边竞品价格数据不足，请手动查看当前搜索结果进行对比。"

    return {"short": short, "detail": detail}


def _generate_daily_report(data_file, own_shop, keyword, log):
    """对比今天和昨天的数据，生成HTML日报和Excel日报"""
    try:
        import pandas as pd
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment
        import re

        wb_data = pd.read_excel(data_file, sheet_name="采集数据")
        if wb_data.empty or "采集时间" not in wb_data.columns:
            return None, None, None

        # 解析日期
        wb_data["_date"] = pd.to_datetime(wb_data["采集时间"], errors="coerce").dt.date
        dates = sorted(wb_data["_date"].dropna().unique())
        if len(dates) < 1:
            return None, None, None

        today = dates[-1]
        df_today = wb_data[wb_data["_date"] == today].copy()
        df_yesterday = wb_data[wb_data["_date"] == dates[-2]].copy() if len(dates) >= 2 else pd.DataFrame()

        # ── 计算指标 ──
        price_col = "当前售价(元)" if "当前售价(元)" in df_today.columns else "当前售价"
        sales_col = "月销量(件)" if "月销量(件)" in df_today.columns else "月销量"
        shop_col  = "店铺名称" if "店铺名称" in df_today.columns else "店铺"
        title_col = "商品标题" if "商品标题" in df_today.columns else "标题"

        # 类型清理：防止NaN/float导致后续报错
        df_today[sales_col]  = pd.to_numeric(df_today[sales_col],  errors="coerce").fillna(0)
        df_today[price_col]  = pd.to_numeric(df_today[price_col],  errors="coerce").fillna(0)
        if not df_yesterday.empty:
            df_yesterday[sales_col] = pd.to_numeric(df_yesterday[sales_col], errors="coerce").fillna(0)
            df_yesterday[price_col] = pd.to_numeric(df_yesterday[price_col], errors="coerce").fillna(0)

        # 今日Top5销量
        top5 = df_today.nlargest(5, sales_col)[[title_col, price_col, sales_col, shop_col]]

        # 自家产品
        own_df = df_today[df_today[shop_col].astype(str).str.contains(own_shop, na=False)]

        # 价格异动（对比昨天，降价超5%）
        price_changes = []
        if not df_yesterday.empty:
            merged = df_today.merge(
                df_yesterday[[title_col, price_col]].rename(columns={price_col: "_old_price"}),
                on=title_col, how="inner")
            merged["_chg"] = (merged[price_col] - merged["_old_price"]) / merged["_old_price"].replace(0, 1) * 100
            drops = merged[merged["_chg"] < -5].nsmallest(5, "_chg")
            for _, r in drops.iterrows():
                price_changes.append({
                    "title": str(r[title_col])[:30],
                    "old": r["_old_price"],
                    "new": r[price_col],
                    "chg": r["_chg"]
                })

        # 排名变化（自家）
        own_rank_info = []
        if not own_df.empty and not df_yesterday.empty:
            df_today_r = df_today.reset_index(drop=True); df_today_r["_rank"] = df_today_r.index + 1
            df_yest_r  = df_yesterday.reset_index(drop=True); df_yest_r["_rank_y"] = df_yest_r.index + 1
            own_today = df_today_r[df_today_r[shop_col].astype(str).str.contains(own_shop, na=False)]
            own_yest  = df_yest_r[df_yest_r[shop_col].astype(str).str.contains(own_shop, na=False)]
            for _, r in own_today.iterrows():
                match = own_yest[own_yest[title_col] == r[title_col]]
                rank_y = int(match["_rank_y"].values[0]) if not match.empty else None
                own_rank_info.append({
                    "title": str(r[title_col])[:30],
                    "rank": int(r["_rank"]),
                    "rank_y": rank_y,
                    "sales": r[sales_col],
                    "price": r[price_col]
                })

        # ── 规则引擎：生成市场简报 ──
        today_prices = df_today[price_col][df_today[price_col] > 0]
        yest_prices  = df_yesterday[price_col][df_yesterday[price_col] > 0] if not df_yesterday.empty else pd.Series(dtype=float)
        avg_today  = round(float(today_prices.mean()), 1) if len(today_prices) else 0
        avg_yest   = round(float(yest_prices.mean()),  1) if len(yest_prices)  else 0
        min_today  = round(float(today_prices.min()),  1) if len(today_prices) else 0
        total_sales_today = int(df_today[sales_col].sum())
        total_sales_yest  = int(df_yesterday[sales_col].sum()) if not df_yesterday.empty else 0
        total_items_today = len(df_today)
        total_items_yest  = len(df_yesterday) if not df_yesterday.empty else 0

        bl = []  # brief_lines
        bl.append(f"📊 {keyword} 市场日报 {today}")
        bl.append("")

        # 价格
        price_line = f"【价格】均价 ¥{avg_today}"
        if avg_yest > 0:
            price_diff = avg_today - avg_yest
            price_pct  = price_diff / avg_yest * 100
            if price_pct <= -5:
                price_line += f"，较昨日下跌 {abs(price_pct):.1f}%↓ ⚠️"
            elif price_pct >= 5:
                price_line += f"，较昨日上涨 {price_pct:.1f}%↑"
            else:
                price_line += f"，与昨日持平（{price_diff:+.1f}元）"
        if min_today > 0:
            price_line += f"，最低价 ¥{min_today}"
        bl.append(price_line)

        # 竞品数量
        new_items_cnt = total_items_today - total_items_yest if total_items_yest > 0 else 0
        comp_line = f"【竞品】共 {total_items_today} 个在售商品"
        if new_items_cnt > 0:
            comp_line += f"，今日新增 {new_items_cnt} 个"
        elif new_items_cnt < 0:
            comp_line += f"，减少 {abs(new_items_cnt)} 个"
        bl.append(comp_line)

        # 促销情况
        promo_col2 = "促销标签" if "促销标签" in df_today.columns else None
        if promo_col2:
            promo_cnt = int((df_today[promo_col2].astype(str).str.len() > 1).sum())
            promo_pct = promo_cnt / total_items_today * 100 if total_items_today else 0
            if promo_pct >= 30:
                bl.append(f"【促销】{promo_cnt} 个商品有促销活动（占比 {promo_pct:.0f}%），市场促销活跃 🔥")
            else:
                bl.append(f"【促销】{promo_cnt} 个商品有促销活动（占比 {promo_pct:.0f}%）")

        # 销量
        sales_line = f"【销量】月销量合计 {total_sales_today:,} 件"
        if total_sales_yest > 0:
            sp = (total_sales_today - total_sales_yest) / total_sales_yest * 100
            if sp >= 10:
                sales_line += f"，较昨日增长 {sp:.1f}%↑"
            elif sp <= -10:
                sales_line += f"，较昨日下降 {abs(sp):.1f}%↓"
            else:
                sales_line += "，与昨日持平"
        bl.append(sales_line)

        # 自家排名
        if own_rank_info:
            avg_rank_v = sum(r["rank"] for r in own_rank_info) / len(own_rank_info)
            rank_line = f"【自家】{own_shop} 平均排名 #{avg_rank_v:.0f}"
            rchgs = [r["rank_y"] - r["rank"] for r in own_rank_info if r["rank_y"]]
            if rchgs:
                avg_chg = sum(rchgs) / len(rchgs)
                if avg_chg > 0:
                    rank_line += f"，排名上升 {avg_chg:.0f} 位↑"
                elif avg_chg < 0:
                    rank_line += f"，排名下滑 {abs(avg_chg):.0f} 位↓ ⚠️"
                else:
                    rank_line += "，排名持平"
            bl.append(rank_line)
        else:
            bl.append(f"【自家】未在搜索结果中找到 {own_shop}")

        # 降价竞品
        if price_changes:
            max_drop = min(p["chg"] for p in price_changes)
            bl.append(f"【降价】{len(price_changes)} 个竞品今日降价超5%，最大降幅 {max_drop:.1f}%")

        # 综合建议
        bl.append("")
        suggestions = []
        if avg_yest > 0 and (avg_today - avg_yest) / avg_yest * 100 <= -5:
            suggestions.append("市场均价下滑，建议评估是否跟进促销或调整定价")
        if price_changes and len(price_changes) >= 3:
            suggestions.append("多个竞品集中降价，可能有大促活动，建议重点关注")
        if own_rank_info:
            rchgs2 = [r["rank_y"] - r["rank"] for r in own_rank_info if r["rank_y"]]
            if rchgs2 and sum(rchgs2) / len(rchgs2) < -3:
                suggestions.append("自家排名明显下滑，建议优化主图/标题或参与活动")
        if not suggestions:
            suggestions.append("市场平稳，继续保持监控")
        bl.append("💡 建议：" + "；".join(suggestions))

        market_brief = "\n".join(bl)
        log.info("[简报] 市场简报生成完成")

        # ── 生成HTML ──
        def trend(v): return "↑" if v and v > 0 else ("↓" if v and v < 0 else "→")
        def color(v): return "#E74C3C" if v and v > 0 else ("#2ECC71" if v and v < 0 else "#888")

        rows_top5 = "".join([
            f"<tr><td>{i+1}</td><td>{str(r[title_col])[:28]}</td>"
            f"<td>¥{r[price_col]}</td><td>{int(r[sales_col]):,}</td>"
            f"<td>{r[shop_col]}</td></tr>"
            for i, (_, r) in enumerate(top5.iterrows())
        ])

        rows_own = "".join([
            f"<tr><td>#{r['rank']}</td><td>{r['title']}</td>"
            f"<td>¥{r['price']}</td><td>{int(r['sales']):,}</td>"
            f"<td style='color:{color(None if r['rank_y'] is None else r['rank_y']-r["rank"])}'>"
            f"{trend(None if r['rank_y'] is None else r['rank_y']-r['rank'])} "
            f"{'(新)' if r['rank_y'] is None else f'昨#{r[chr(114)+chr(97)+chr(110)+chr(107)+chr(95)+chr(121)]}'}</td></tr>"
            for r in own_rank_info
        ]) or "<tr><td colspan='5' style='color:#888'>未找到自家产品</td></tr>"

        rows_price = "".join([
            f"<tr><td>{p['title']}</td><td>¥{p['old']:.1f}</td>"
            f"<td>¥{p['new']:.1f}</td>"
            f"<td style='color:#E74C3C'>{p['chg']:.1f}%</td></tr>"
            for p in price_changes
        ]) or "<tr><td colspan='4' style='color:#888'>今日无明显降价</td></tr>"

        # 明日关注
        tomorrow_tips = []
        if price_changes:
            tomorrow_tips.append(f"⚠️ {len(price_changes)}个竞品降价，关注是否持续")
        if own_rank_info:
            avg_rank = sum(r["rank"] for r in own_rank_info) / len(own_rank_info)
            tomorrow_tips.append(f"📌 自家产品平均排名 #{avg_rank:.0f}，{'需优化' if avg_rank > 10 else '保持优势'}")
        top1 = top5.iloc[0] if not top5.empty else None
        if top1 is not None:
            tomorrow_tips.append(f"🏆 今日销冠：{str(top1[title_col])[:20]}（月销{int(top1[sales_col]):,}件，¥{top1[price_col]}）")
        if not tomorrow_tips:
            tomorrow_tips.append("✅ 市场平稳，继续监控")

        tips_html = "".join(f"<li>{t}</li>" for t in tomorrow_tips)

        td_style = "padding:8px 12px;border-bottom:1px solid #eee;"
        th_style = "padding:8px 12px;background:#2A5298;color:white;text-align:left;"
        table_style = "width:100%;border-collapse:collapse;margin:12px 0;font-size:13px;"

        html = f"""
<html><body style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto;padding:20px;color:#333;">
<h2 style="color:#E94560;border-bottom:2px solid #E94560;padding-bottom:8px;">
  📊 淘宝日报 — {keyword} — {today}</h2>

<h3 style="color:#2A5298;">🏆 今日销量Top5</h3>
<table style="{table_style}">
<tr><th style="{th_style}">排名</th><th style="{th_style}">商品</th>
<th style="{th_style}">价格</th><th style="{th_style}">月销量</th><th style="{th_style}">店铺</th></tr>
{rows_top5}
</table>

<h3 style="color:#2A5298;">🏪 自家产品（{own_shop}）</h3>
<table style="{table_style}">
<tr><th style="{th_style}">排名</th><th style="{th_style}">商品</th>
<th style="{th_style}">价格</th><th style="{th_style}">月销量</th><th style="{th_style}">排名变化</th></tr>
{rows_own}
</table>

<h3 style="color:#2A5298;">💰 价格异动（降价超5%）</h3>
<table style="{table_style}">
<tr><th style="{th_style}">商品</th><th style="{th_style}">昨日价</th>
<th style="{th_style}">今日价</th><th style="{th_style}">变化</th></tr>
{rows_price}
</table>

<h3 style="color:#2A5298;">🔭 明日关注</h3>
<ul style="background:#f8f9fa;padding:16px 28px;border-radius:6px;line-height:2;">
{tips_html}
</ul>

<h3 style="color:#E94560;">📋 市场简报（可直接转发）</h3>
<div style="background:#1A1A2E;color:#F0F0FF;padding:20px;border-radius:8px;
            font-size:14px;line-height:2;white-space:pre-line;font-family:monospace;">
{market_brief}
</div>

<p style="color:#aaa;font-size:11px;margin-top:20px;">
  本报告由淘宝数据采集工具自动生成 · {datetime.now().strftime("%Y-%m-%d %H:%M")}</p>
</body></html>"""

        # ── 生成日报Excel ──
        report_path = os.path.join(os.path.dirname(data_file),
                                   f"日报_{keyword}_{today}.xlsx")
        wb = Workbook()
        ws = wb.active; ws.title = "日报"
        title_fill = PatternFill("solid", fgColor="2A5298")
        title_font = Font(color="FFFFFF", bold=True)

        def write_section(ws, row, title, headers, data_rows):
            ws.cell(row, 1, title).font = Font(bold=True, size=12, color="E94560")
            row += 1
            for ci, h in enumerate(headers, 1):
                c = ws.cell(row, ci, h)
                c.fill = title_fill; c.font = title_font
                c.alignment = Alignment(horizontal="center")
            row += 1
            for dr in data_rows:
                for ci, val in enumerate(dr, 1):
                    ws.cell(row, ci, val)
                row += 1
            return row + 1

        r = 1
        top5_rows = [(i+1, str(row[title_col])[:30], float(row[price_col]) if row[price_col] else 0,
                      int(float(row[sales_col])) if row[sales_col] else 0, str(row[shop_col]))
                     for i, (_, row) in enumerate(top5.iterrows())]
        r = write_section(ws, r, f"今日销量Top5 ({today})",
                          ["排名","商品标题","价格","月销量","店铺"], top5_rows)

        own_rows = [(f"#{ri['rank']}", ri["title"], ri["price"],
                     int(ri["sales"]),
                     f"↑{ri['rank_y']-ri['rank']}" if ri["rank_y"] and ri["rank_y"] > ri["rank"]
                     else (f"↓{ri['rank']-ri['rank_y']}" if ri["rank_y"] and ri["rank_y"] < ri["rank"]
                     else "新上榜" if not ri["rank_y"] else "→"))
                    for ri in own_rank_info] or [("–","未找到自家产品","–","–","–")]
        r = write_section(ws, r, f"自家产品排名（{own_shop}）",
                          ["排名","商品标题","价格","月销量","排名变化"], own_rows)

        price_rows = [(p["title"], f"¥{p['old']:.1f}", f"¥{p['new']:.1f}",
                       f"{p['chg']:.1f}%") for p in price_changes] or [("今日无明显降价","–","–","–")]
        r = write_section(ws, r, "价格异动（降价超5%）",
                          ["商品","昨日价","今日价","变化幅度"], price_rows)

        r = write_section(ws, r, "明日关注",
                          ["建议"], [(t,) for t in tomorrow_tips])

        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

        wb.save(report_path)
        log.info(f"日报Excel已生成: {report_path}")
        return html, report_path, market_brief

    except Exception as e:
        log.error(f"日报生成错误: {e}")
        import traceback; traceback.print_exc()
        return None, None, None


def main():
    root = tk.Tk()
    app = App(root)

    def on_quit():
        app._save_config()  # 退出时保存配置
        if app.is_running:
            if messagebox.askyesno(app.t("quit_title"), app.t("quit_msg")):
                app.stop_ev.set()
                root.destroy()
        else:
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_quit)
    root.mainloop()


if __name__ == "__main__":
    main()
