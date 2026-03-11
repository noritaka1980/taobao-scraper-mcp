# -*- coding: utf-8 -*-
"""
scraper_core.py — 核心采集逻辑 v2.2
图片：JS获取currentSrc，兼容懒加载
反检测：增强随机延时
修复：safe_url(driver)为None时崩溃 / 评价数选择器更新
"""

import time, random, re, os, io, logging, traceback, urllib.request
from datetime import datetime
from urllib.parse import quote

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage


def safe_url(driver):
    """安全获取当前URL，多种方式兜底"""
    # Method 1: direct property (fastest, most reliable)
    try:
        url = driver.current_url
        if url and url != "data:," and url != "about:blank":
            return url
        if url:
            return url
    except Exception:
        pass
    # Method 2: execute_script
    try:
        url = driver.execute_script("return window.location.href;")
        if url:
            return url
    except Exception:
        pass
    return ""


def safe_title(driver):
    """安全获取当前页面标题，多种方式兜底"""
    try:
        t = driver.title
        if t:
            return t
    except Exception:
        pass
    try:
        t = driver.execute_script("return document.title;")
        if t:
            return t
    except Exception:
        pass
    return ""


_chrome_offscreen = False  # 是否已将Chrome移到屏幕外

def keep_minimized(driver):
    """将浏览器移到屏幕外并最小化，防止采集中弹出干扰用户。
    每次调用都通过CDP将窗口移到(-32000,-32000)并最小化，
    因为Selenium操作(driver.get/switch_to/send_keys等)会恢复窗口位置。"""
    global _chrome_offscreen
    # 方案1: CDP 移到屏幕外 + 最小化（最可靠）
    try:
        winfo = driver.execute_cdp_cmd('Browser.getWindowForTarget', {})
        wid = winfo['windowId']
        # 每次都重新定位到屏幕外（Selenium操作可能恢复窗口位置）
        bounds = driver.execute_cdp_cmd('Browser.getWindowBounds', {'windowId': wid})
        if bounds.get('windowState') in ('minimized', 'maximized', 'fullscreen'):
            driver.execute_cdp_cmd('Browser.setWindowBounds', {
                'windowId': wid, 'bounds': {'windowState': 'normal'}
            })
        driver.execute_cdp_cmd('Browser.setWindowBounds', {
            'windowId': wid,
            'bounds': {'left': -32000, 'top': -32000}
        })
        _chrome_offscreen = True
        # 最小化
        driver.execute_cdp_cmd('Browser.setWindowBounds', {
            'windowId': wid, 'bounds': {'windowState': 'minimized'}
        })
        return
    except Exception:
        pass
    # 方案2: Win32 API 强制最小化前台Chrome窗口
    try:
        import ctypes
        user32 = ctypes.windll.user32
        hwnd = user32.GetForegroundWindow()
        cls = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls, 256)
        if cls.value == "Chrome_WidgetWin_1":
            user32.ShowWindow(hwnd, 11)  # SW_FORCEMINIMIZE
    except Exception:
        pass
    # 方案3: Selenium fallback
    try:
        driver.minimize_window()
    except Exception:
        pass

_saved_foreground = None
def save_foreground():
    pass
def restore_foreground():
    pass
def patch_selenium_driver(driver):
    pass

def restore_chrome_window(driver):
    """采集结束后将Chrome窗口恢复到屏幕可见位置"""
    global _chrome_offscreen
    try:
        winfo = driver.execute_cdp_cmd('Browser.getWindowForTarget', {})
        wid = winfo['windowId']
        # 先恢复为普通窗口状态
        driver.execute_cdp_cmd('Browser.setWindowBounds', {
            'windowId': wid, 'bounds': {'windowState': 'normal'}
        })
        # 移回屏幕可见区域
        driver.execute_cdp_cmd('Browser.setWindowBounds', {
            'windowId': wid,
            'bounds': {'left': 100, 'top': 100, 'width': 1200, 'height': 800}
        })
        _chrome_offscreen = False
    except Exception:
        # fallback: Selenium方式
        try:
            driver.set_window_position(100, 100)
            driver.set_window_size(1200, 800)
            _chrome_offscreen = False
        except Exception:
            pass


def rsleep(a=2, b=5, stop_event=None):
    end = time.time() + random.uniform(a, b)
    while time.time() < end:
        if stop_event and stop_event.is_set():
            return
        time.sleep(0.2)


def parse_num(text):
    if not text:
        return 0
    t = str(text).replace(",", "").strip()
    m = re.search(r"([\d.]+)\s*万", t)
    if m:
        return int(float(m.group(1)) * 10000)
    m = re.search(r"\d+", t)
    return int(m.group()) if m else 0


def clean_price(text):
    if not text:
        return 0.0
    t = re.sub(r"[^\d.]", "", str(text).split("\n")[0])
    try:
        return float(t)
    except Exception:
        return 0.0



def extract_spec(title: str, subtitle: str = "") -> str:
    """从商品标题+副标题中提取产品规格信息"""
    import re as _re
    text = (title + " " + subtitle).strip()
    specs = []
    seen = set()

    def add(s):
        s = s.strip()
        if s and s not in seen:
            seen.add(s); specs.append(s)

    # 容量：150ml、30g、150mL 等（含带+的组合 150ml+110ml）
    for m in _re.finditer(r'\d+\.?\d*\s*(?:ml|mL|ML|g|G|kg|KG|oz|OZ|升|克|毫升)', text):
        add(m.group())

    # 套装件数：2件套、3件装 等
    for m in _re.finditer(r'\d+\s*(?:件套|件装|套装|支装|瓶装|盒装|件组合|合一|件)', text):
        add(m.group())

    # 数量×规格组合：2*150ml、3x50g 等
    for m in _re.finditer(r'\d+\s*[xX×*]\s*\d+\.?\d*\s*(?:ml|mL|g|G|片|粒|支|瓶)?', text):
        add(m.group())

    # 关键词规格
    for kw in ['大号','中号','小号','加大','加量','升级版','正装','旅行装',
               '小样','体验装','礼盒装','礼盒','套盒','pro版','Pro版','mini','Mini']:
        if kw.lower() in text.lower():
            add(kw)

    # 两件套/三件套等中文数字
    for m in _re.finditer(r'[两三四五六七八九十]+\s*(?:件套|件装|套装)', text):
        add(m.group())

    return ' / '.join(specs) if specs else ""

def extract_bundle_type(title: str, subtitle: str = "") -> str:
    """从商品标题识别套装类型，返回标准化标签"""
    import re as _re
    text = (title + " " + subtitle).lower()

    # 优先匹配件数明确的套装
    patterns = [
        (r'十\s*件套|10\s*件套',        '十件套'),
        (r'九\s*件套|9\s*件套',          '九件套'),
        (r'八\s*件套|8\s*件套',          '八件套'),
        (r'七\s*件套|7\s*件套',          '七件套'),
        (r'六\s*件套|6\s*件套',          '六件套'),
        (r'五\s*件套|5\s*件套',          '五件套'),
        (r'四\s*件套|4\s*件套',          '四件套'),
        (r'三\s*件套|3\s*件套',          '三件套'),
        (r'两\s*件套|二\s*件套|2\s*件套', '两件套'),
        (r'单品|单件',                    '单品'),
        (r'礼盒装|礼盒|gift\s*set',       '礼盒装'),
        (r'套盒|套装',                    '套装'),
        (r'旅行装|旅行套',                '旅行装'),
        (r'小样|体验装|试用装',           '小样/体验装'),
    ]
    for pattern, label in patterns:
        if _re.search(pattern, text):
            return label
    return ''


def get_text(el, selectors):
    for sel in selectors:
        try:
            found = el.find_elements(By.CSS_SELECTOR, sel)
            for f in found:
                t = f.text.strip()
                if t:
                    return t
        except StaleElementReferenceException:
            pass
    return ""


def get_link(card):
    try:
        # 1. 卡片自身可能就是<a>标签（新版淘宝）
        if card.tag_name == "a":
            href = card.get_attribute("href") or ""
            if "item.taobao.com" in href or "detail.tmall.com" in href:
                return re.sub(r"\?.*", "", href)
        # 2. 内部<a>标签
        for a in card.find_elements(By.TAG_NAME, "a"):
            href = a.get_attribute("href") or ""
            if "item.taobao.com" in href or "detail.tmall.com" in href:
                return re.sub(r"\?.*", "", href)
        # 3. 从 id/data-spm-act-id 提取商品ID拼接链接
        item_id = ""
        el_id = card.get_attribute("id") or ""
        if el_id.startswith("item_id_"):
            item_id = el_id.replace("item_id_", "")
        if not item_id:
            item_id = card.get_attribute("data-spm-act-id") or ""
        if item_id and item_id.isdigit():
            return f"https://item.taobao.com/item.htm?id={item_id}"
    except Exception:
        pass
    return ""


def get_image_url(card, driver=None):
    """获取商品主图URL
    淘宝图片格式: https://img.alicdn.com/xxx.jpg_360x360q90.jpg_.webp
    处理后变为:  https://img.alicdn.com/xxx.jpg
    """
    def clean_url(src):
        if not src or not src.strip(): return ""
        src = src.strip()
        if src.startswith("data:"): return ""
        if src.startswith("//"): src = "https:" + src
        if src.startswith("http://"): src = "https://" + src[7:]
        if not src.startswith("https://"): return ""
        # 去掉淘宝尺寸后缀: .jpg_360x360q90.jpg_.webp -> .jpg
        src = re.sub(r'\.(jpg|jpeg|png|webp)_[\w.]+$', r'.\1', src)
        if not any(src.endswith(e) for e in ['.jpg','.jpeg','.png','.webp']):
            return ""
        return src

    try:
        # 优先找 class含mainImg 的img（淘宝主图标准class名）
        for sel in ["[class*='mainImg']", "[class*='main-img']", "[class*='mainPic'] img"]:
            try:
                for img in card.find_elements(By.CSS_SELECTOR, sel):
                    for attr in ["src", "data-src", "data-lazy-src"]:
                        raw = img.get_attribute(attr) or ""
                        url = clean_url(raw)
                        if url: return url
            except Exception: continue

        # 兜底：找所有img里alicdn域名的
        for img in card.find_elements(By.TAG_NAME, "img"):
            for attr in ["src", "data-src", "data-lazy-src"]:
                raw = img.get_attribute(attr) or ""
                if "alicdn.com" in raw or "taobaocdn" in raw:
                    url = clean_url(raw)
                    if url: return url
    except Exception: pass
    return ""


def download_image(url, size=80):
    """下载图片，返回 PIL Image 对象的 BytesIO"""
    try:
        if not url:
            return None
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Referer": "https://s.taobao.com/",
        }
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = resp.read()
        # 用PIL缩放到统一尺寸
        try:
            from PIL import Image as PILImage
            img = PILImage.open(io.BytesIO(data))
            img = img.convert("RGB")
            img = img.resize((size, size), PILImage.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=85)
            buf.seek(0)
            return buf
        except ImportError:
            # 没有PIL就直接返回原始数据
            return io.BytesIO(data)
    except Exception as e:
        return None


CARD_SELS = [
    "[class*='doubleCardWrapper']",
    "[class*='cardWrapper']",
    "[class*='itemWrapper']",
    "[data-item-id]",
    ".item.J_MouserOnverReq",
]


def scroll_page(driver, stop_event=None):
    h = driver.execute_script("return document.body.scrollHeight")
    for pos in range(0, h + 300, 400):
        if stop_event and stop_event.is_set():
            return
        driver.execute_script(f"window.scrollTo(0, {pos});")
        time.sleep(random.uniform(0.05, 0.15))
    rsleep(1, 2, stop_event)


def scrape_page(driver, cfg, log, stop_event=None, now=None):  # driver传给get_image_url
    log.info(f"  采集当前页: {safe_url(driver)[:70]}...")
    if "deny_h5" in safe_url(driver) or "punish" in safe_url(driver):
        log.warning("  页面被拦截，跳过")
        return []
    if "login" in safe_url(driver) or "passport" in safe_url(driver):
        log.warning("  需要登录，跳过")
        return []

    scroll_page(driver, stop_event)
    if stop_event and stop_event.is_set():
        return []
    # 等待懒加载图片触发
    time.sleep(3)

    cards = []
    for sel in CARD_SELS:
        found = driver.find_elements(By.CSS_SELECTOR, sel)
        if len(found) >= 5:
            cards = found
            log.info(f"  Found {len(found)} product cards")
            break

    if not cards:
        log.warning("  No cards found! Saving screenshot: debug_screenshot.png")
        driver.save_screenshot("debug_screenshot.png")
        log.warning(f"  URL was: {safe_url(driver)}")
        return []

    items = []
    # now 由调用方（run_round）传入，保证整轮采集时间戳一致
    if now is None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M")

    for card in cards:
        if stop_event and stop_event.is_set():
            break
        try:
            title = get_text(card, [
                "[class*='title--']", "[class*='itemTitle']",
                "h3[class*='title']", ".title", "h3"])
            subtitle = get_text(card, [
                "[class*='desc--']", "[class*='subTitle']", "[class*='sub-title']",
                "[class*='subtitle']", "[class*='note--']", "[class*='label--']"])
            price_raw = get_text(card, [
                "[class*='priceInt']", "[class*='price--']",
                "[class*='price']", ".price"])
            price = clean_price(price_raw)
            if not title or not price:
                continue

            orig_raw = get_text(card, [
                "[class*='originalPrice']", "[class*='origin-price']", ".g-price-del"])
            orig = clean_price(orig_raw) if orig_raw else ""

            sales_raw = get_text(card, [
                "[class*='realSales']", "[class*='monthSold']",
                "[class*='sale--']", "[class*='sold']"])
            sales = parse_num(sales_raw)

            shop = get_text(card, [
                "[class*='shopName']", "[class*='store--']",
                "[class*='shopInfo']", ".shop-name"])
            reviews = parse_num(get_text(card, [
                "[class*='comment']", "[class*='review']", "[class*='rateNum']",
                "[class*='realSales'] + *", "[class*='sold--']",
                "[class*='count--']", "[class*='evaluateNum']",
                "[class*='ContentRow--contentInner'] span:last-child"]))
            location = get_text(card, [
                "[class*='location']", "[class*='place']",
                "[class*='area']", "[class*='region']"])

            link = get_link(card)
            img_url = get_image_url(card, driver)

            is_tmall = "tmall" in link.lower() or bool(
                card.find_elements(By.CSS_SELECTOR, "[class*='tmall'],[class*='Tmall']"))
            shop_type = "Tmall" if is_tmall else "Taobao"

            promo_els = []
            for sel in ["[class*='promoText']", "[class*='tagList']", "[class*='label']"]:
                try:
                    promo_els += card.find_elements(By.CSS_SELECTOR, sel)
                except Exception:
                    pass
            promo = " | ".join(
                e.text.strip() for e in promo_els[:3]
                if e.text.strip() and len(e.text.strip()) < 15)

            pmin = cfg.get("price_min", 0)
            pmax = cfg.get("price_max", 0)
            if pmin and price < pmin:
                continue
            if pmax and price > pmax:
                continue
            if cfg.get("tmall_only") and shop_type != "Tmall":
                continue

            discount = ""
            if orig and orig > price:
                discount = f"{round(price / orig * 10, 1)}折"

            items.append({
                "采集时间":     now,
                "商品主图":     img_url,   # 图片URL，保存后嵌入
                "商品标题":     title,
                "产品规格":     extract_spec(title, subtitle),
                "套装类型":     extract_bundle_type(title, subtitle),
                "当前售价(元)": price,
                "原价(元)":     orig,
                "折扣":         discount,
                "月销量(件)":   sales,
                "评价数量":     reviews,
                "店铺名称":     shop,
                "店铺类型":     shop_type,
                "店铺地区":     location,
                "促销标签":     promo,
                "商品链接":     link,
            })
        except Exception as e:
            log.debug(f"  Card error: {e}")
            continue

    return items


def _insert_images_win32(xlsx_path, img_urls, img_size, log):
    """用win32com往Excel/WPS里插入图片，兼容WPS"""
    try:
        import win32com.client as win32
        import os, tempfile, urllib.request

        abs_path = os.path.abspath(xlsx_path)
        app = win32.Dispatch("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(abs_path)

        # 找"采集数据"sheet
        ws = None
        for sh in wb.Sheets:
            if "采集" in sh.Name or sh.Index == 1:
                ws = sh
                break
        if ws is None:
            wb.Close(False)
            app.Quit()
            return 0

        # 先删除A列现有图片
        for sh in [ws]:
            for pic in list(sh.Pictures()):
                try: pic.Delete()
                except Exception: pass

        ok = 0
        tmp_dir = tempfile.mkdtemp()
        for i, url in enumerate(img_urls):
            if not url or str(url) == "nan" or not isinstance(url, str):
                continue
            url = str(url).strip()
            if not url:
                continue
            row = i + 2  # 数据从第2行
            try:
                tmp_file = os.path.join(tmp_dir, f"img_{i}.jpg")
                req = urllib.request.Request(url, headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Referer": "https://s.taobao.com/"
                })
                with urllib.request.urlopen(req, timeout=10) as r:
                    with open(tmp_file, 'wb') as f:
                        f.write(r.read())

                cell = ws.Cells(row, 1)  # A列
                left = cell.Left + 2
                top = cell.Top + 2
                pic = ws.Pictures().Insert(tmp_file)
                pic.Left = left
                pic.Top = top
                pic.Width = img_size - 4
                pic.Height = img_size - 4
                pic.Placement = 1  # xlMoveAndSize
                ok += 1
            except Exception as e:
                log.debug(f"  win32 row{row}: {e}")
                continue

        wb.Save()
        wb.Close(True)
        app.Quit()

        # 清理临时文件
        import shutil
        try: shutil.rmtree(tmp_dir)
        except Exception: pass

        return ok
    except ImportError:
        return 0  # 没有win32com，跳过
    except Exception as e:
        log.debug(f"  win32com error: {e}")
        return 0



# ── 各sheet列备注说明 ──
COL_COMMENTS = {
    "采集数据": {
        "商品主图":     "商品缩略图（自动从淘宝/京东抓取嵌入）",
        "采集时间":     "数据采集的日期和时间",
        "商品标题":     "商品在平台上显示的完整标题",
        "产品规格":     "从标题中自动提取的规格信息，如容量、克重等",
        "套装类型":     "自动识别的套装件数，如两件套/三件套/礼盒装等",
        "当前售价(元)": "采集时商品的实际到手价格（元）",
        "原价(元)":     "划线原价，有促销时显示（元）",
        "折扣":         "当前售价相对原价的折扣，如8.5折",
        "月销量(件)":   "该商品近30天的销售件数",
        "评价数量":     "商品累计收到的买家评价总数",
        "店铺名称":     "销售该商品的店铺名称",
        "店铺类型":     "Tmall=天猫旗舰店，Taobao=淘宝店",
        "店铺地区":     "店铺所在省市",
        "促销标签":     "商品当前的促销活动标签，如满减、秒杀等",
        "商品链接":     "商品详情页URL，点击可直接跳转",
        "累计总销量(估算)": "月销量 × 估算上架月数（评价数÷月均销量），反映商品长期销售总量",
    },
    "汇总统计": {
        "采集时间":   "本次采集的日期和时间\n注意：同一天多次采集只保留最新一次，旧数据自动覆盖",
        "商品总数":   "本次采集到的商品总条数\n计算：搜索结果页中所有解析成功的商品卡片数量",
        "最低价(元)": "本次采集中所有商品的最低售价\n计算：MIN(当前售价列)",
        "最高价(元)": "本次采集中所有商品的最高售价\n计算：MAX(当前售价列)",
        "平均价(元)": "所有商品售价的算术平均值\n计算：SUM(售价)÷商品总数（含促销低价，可能偏低）",
        "中位价(元)": "按价格排序后中间位置的价格\n计算：MEDIAN(售价列)，比平均价更能反映真实市场定价",
        "月销量合计": "本次采集所有商品月销量之和\n计算：SUM(月销量列)，数值越大说明该品类市场越活跃",
        "月销量最高": "本次采集中单品最高月销量\n计算：MAX(月销量列)，反映头部爆款的销量天花板",
        "天猫店数":   "本次采集中天猫旗舰店的数量\n计算：店铺类型=Tmall 的商品条数",
        "淘宝店数":   "本次采集中淘宝普通店的数量\n计算：店铺类型=Taobao 的商品条数",
    },
    "价格分布": {
        "价格区间": "按价格段划分的区间，如200-400元",
        "商品数量": "落在该价格区间内的商品数量",
        "占比(%)":  "该价格区间商品数量占全部商品的百分比",
    },
    "TOP50店铺": {
        "排名":     "按月销量合计从高到低的排名",
        "店铺名称": "店铺名称",
        "店铺类型": "Tmall=天猫旗舰店，Taobao=淘宝店",
        "店铺地区": "店铺所在省市",
        "在售商品数": "该店铺在本次搜索结果中出现的商品数量",
        "月销量合计": "该店铺所有在售商品月销量之和",
        "最高月销量": "该店铺单品最高月销量",
        "平均售价":   "该店铺所有商品的平均售价（元）",
    },
    "京东数据": {
        "采集时间":     "数据采集的日期和时间",
        "平台":         "数据来源平台（京东）",
        "商品主图":     "商品缩略图",
        "商品标题":     "商品在京东上显示的完整标题",
        "产品规格":     "从标题中自动提取的规格信息",
        "套装类型":     "自动识别的套装类型",
        "当前售价(元)": "采集时京东商品的实际价格（元）",
        "月销量(件)":   "商品评价数量（京东以评价数代替销量）",
        "店铺名称":     "销售该商品的京东店铺名称",
        "商品链接":     "京东商品详情页URL",
    },
    "跨平台对比": {
        "平台":         "数据来源（淘宝/京东）",
        "商品总数":     "该平台采集到的商品总数",
        "平均售价(元)": "该平台所有商品的平均售价",
        "最低价(元)":   "该平台最低商品售价",
        "最高价(元)":   "该平台最高商品售价",
        "月销量合计":   "该平台所有商品月销量之和",
    },
    "评价分析": {
        "商品标题": "被分析评价的商品标题",
        "评价内容": "买家评价的原文内容",
        "评价时间": "买家发表评价的时间",
        "评价星级": "买家给出的星级评分（1-5星）",
    },
    "热门关键词": {
        "排名":         "搜索热度排名（越靠前越热门）",
        "淘宝热门关键词": "淘宝搜索框联想出的相关搜索词，反映买家真实搜索习惯",
        "京东热门关键词": "京东搜索框联想出的相关搜索词，反映买家真实搜索习惯",
    },
    "数据图表": {
        "说明": "本sheet包含4张自动生成的数据分析图表：价格分布柱状图、销量Top10条形图、店铺类型饼图、价格-销量散点图",
    },
}


def apply_col_comments(ws, sheet_name):
    """给指定sheet的表头行（第1行或第2行）加备注"""
    from openpyxl.comments import Comment
    comments = COL_COMMENTS.get(sheet_name, {})
    if not comments:
        return
    # 找表头行（第1行有内容的行）
    for header_row in [1, 2]:
        if ws.cell(row=header_row, column=1).value:
            for col in ws.iter_cols(min_row=header_row, max_row=header_row):
                cell = col[0]
                col_name = str(cell.value or "").strip()
                if col_name in comments:
                    c = Comment(comments[col_name], "采集工具")
                    c.width = 280
                    c.height = 60
                    cell.comment = c
            break


def save_excel(new_data, filepath, log):
    if not new_data:
        log.warning("No data to save.")
        return

    df_new = pd.DataFrame(new_data)

    if os.path.exists(filepath):
        try:
            df_old = pd.read_excel(filepath, sheet_name="采集数据")
            # 旧数据没有"商品主图"列时补空列，避免合并后丢失
            if "商品主图" not in df_old.columns:
                df_old.insert(0, "商品主图", "")
            if "产品规格" not in df_old.columns:
                df_old.insert(3, "产品规格", "")
            if "套装类型" not in df_old.columns:
                df_old.insert(4, "套装类型", "")
            if "累计总销量(估算)" not in df_old.columns:
                df_old["累计总销量(估算)"] = 0
            df_combined = pd.concat([df_old, df_new], ignore_index=True)
            # 同一天的数据只保留最新一次采集（按日期去重，保留最后）
            df_combined["_date"] = pd.to_datetime(
                df_combined["采集时间"], errors="coerce").dt.date
            # 找最新采集日期
            latest_date = pd.to_datetime(
                df_new["采集时间"], errors="coerce").dt.date.max()
            # 删除旧数据中与本次采集日期相同的记录，再拼接新数据
            df_old_clean = df_combined[df_combined["_date"] != latest_date].drop(
                columns=["_date"])
            df_new_clean  = df_new.copy()
            df = pd.concat([df_old_clean, df_new_clean], ignore_index=True)
        except Exception as e:
            log.error(f"  ⚠ 合并历史数据失败: {e}，保留新数据，旧数据不丢弃")
            # 尝试备份旧文件，避免数据丢失
            try:
                import shutil
                backup = filepath.replace(".xlsx", "_backup.xlsx")
                shutil.copy2(filepath, backup)
                log.info(f"  旧文件已备份到: {backup}")
            except Exception:
                pass
            df = df_new
    else:
        df = df_new
    # 确保关键列存在
    if "商品主图" not in df.columns:
        df.insert(0, "商品主图", "")
    if "产品规格" not in df.columns:
        df.insert(3, "产品规格", "")
    if "套装类型" not in df.columns:
        df.insert(4, "套装类型", "")

    # 汇总统计
    summary = []
    for t, g in df.groupby("采集时间"):
        prices = pd.to_numeric(g["当前售价(元)"], errors="coerce").dropna()
        sales  = pd.to_numeric(g["月销量(件)"],   errors="coerce").dropna()
        summary.append({
            "采集时间":   t,
            "商品总数":   len(g),
            "最低价(元)": round(prices.min(), 2) if len(prices) else "",
            "最高价(元)": round(prices.max(), 2) if len(prices) else "",
            "平均价(元)": round(prices.mean(), 2) if len(prices) else "",
            "中位价(元)": round(prices.median(), 2) if len(prices) else "",
            "月销量合计": int(sales.sum()) if len(sales) else 0,
            "月销量最高": int(sales.max()) if len(sales) else 0,
            "天猫店数":   int((g["店铺类型"] == "Tmall").sum()),
            "淘宝店数":   int((g["店铺类型"] == "Taobao").sum()),
        })
    df_sum = pd.DataFrame(summary)

    prices_all = pd.to_numeric(df["当前售价(元)"], errors="coerce").dropna()
    bins   = [0, 200, 400, 600, 800, 1200, 99999]
    labels = ["200以下", "200-400", "400-600", "600-800", "800-1200", "1200以上"]
    dist = pd.cut(prices_all, bins=bins, labels=labels).value_counts().sort_index()
    total = len(prices_all)
    df_dist = pd.DataFrame({
        "价格区间": dist.index,
        "商品数量": dist.values,
        "占比(%)": (dist.values / total * 100).round(1) if total else dist.values * 0,
    })

    sg = df.groupby("店铺名称").agg(
        店铺类型   =("店铺类型", "first"),
        店铺地区   =("店铺地区", "first"),
        在售商品数 =("商品标题", "count"),
        月销量合计 =("月销量(件)", "sum"),
        最高月销量 =("月销量(件)", "max"),
        平均售价   =("当前售价(元)", "mean"),
    ).reset_index().sort_values("月销量合计", ascending=False).head(50)
    sg["平均售价"] = sg["平均售价"].round(2)
    sg.insert(0, "排名", range(1, len(sg) + 1))

    # ── 计算每个商品的"累计总销量"（月销量 × 上架天数估算）──
    # 上架天数 = 评价数量 / (月销量/30)，若无评价则用1个月
    import numpy as np
    sales_num  = pd.to_numeric(df["月销量(件)"],  errors="coerce").fillna(0)
    review_num = pd.to_numeric(df.get("评价数量", pd.Series([0]*len(df))),
                               errors="coerce").fillna(0)
    # 上架月数估算：评价数 / 月均销量，最少1个月，最多60个月（5年）
    monthly    = sales_num.replace(0, np.nan)
    est_months = (review_num / monthly).clip(1, 60).fillna(1).round(1)
    df["累计总销量(估算)"] = (sales_num * est_months).astype(int)

    # 写入Excel（不含图片列，图片单独嵌入）
    df_for_excel = df.drop(columns=["商品主图"], errors="ignore")
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        df_for_excel.to_excel(writer, sheet_name="采集数据",  index=False)
        df_sum.to_excel(writer,       sheet_name="汇总统计",  index=False)
        df_dist.to_excel(writer,      sheet_name="价格分布",  index=False)
        sg.to_excel(writer,           sheet_name="TOP50店铺", index=False)

    # 美化样式
    RED  = PatternFill("solid", fgColor="C0392B")
    GRAY = PatternFill("solid", fgColor="F5F5F5")
    WHT  = PatternFill("solid", fgColor="FFFFFF")
    WF   = Font(bold=True, color="FFFFFF", size=10, name="Arial")
    NF   = Font(size=9, name="Arial")
    CA   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LA   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    BD   = Border(
        left=Side(style="thin",   color="DDDDDD"),
        right=Side(style="thin",  color="DDDDDD"),
        top=Side(style="thin",    color="DDDDDD"),
        bottom=Side(style="thin", color="DDDDDD"),
    )

    wb = load_workbook(filepath)

    # ── 先处理非"采集数据"的sheet样式 ──
    for ws in wb.worksheets:
        if ws.title == "采集数据":
            continue
        for cell in ws[1]:
            cell.fill, cell.font, cell.alignment, cell.border = RED, WF, CA, BD
        for i, row in enumerate(ws.iter_rows(min_row=2)):
            fill = GRAY if i % 2 == 0 else WHT
            for cell in row:
                cell.fill, cell.font, cell.alignment, cell.border = fill, NF, LA, BD
        for col in ws.columns:
            w = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(w + 4, 50)
        ws.freeze_panes = "A2"
        apply_col_comments(ws, ws.title)   # 给表头加备注

    # ── 处理"采集数据"sheet，插入图片列 ──
    ws_data = wb["采集数据"]

    # 在A列前插入一列作为图片列
    ws_data.insert_cols(1)
    ws_data.cell(row=1, column=1).value = "商品主图"
    ws_data.cell(row=1, column=1).fill = RED
    ws_data.cell(row=1, column=1).font = WF
    ws_data.cell(row=1, column=1).alignment = CA
    ws_data.cell(row=1, column=1).border = BD
    apply_col_comments(ws_data, "采集数据")   # 给采集数据表头加备注

    # 设置图片列宽和行高
    IMG_SIZE = 80
    ws_data.column_dimensions["A"].width = 12

    # 遍历全部数据行嵌入图片；旧数据的URL为空会自动跳过，新数据正常嵌入
    img_urls = [str(u) if (u and str(u) != "nan") else "" for u in (df["商品主图"].tolist() if "商品主图" in df.columns else [])]
    filled = [u for u in img_urls if u]
    log.info(f"  图片URL: {len(filled)}/{len(img_urls)} 条有效（本次新增 {len(df_new)} 行）")
    if filled:
        log.info(f"  示例URL: {filled[0][:100]}")
    else:
        log.warning("  所有商品的图片URL均为空！请检查采集是否被拦截")

    ok_count = 0
    for i, url in enumerate(img_urls):
        row_num = i + 2
        if not url:
            continue
        ws_data.row_dimensions[row_num].height = 65

        # 样式
        fill = GRAY if i % 2 == 0 else WHT
        for col in range(2, ws_data.max_column + 1):
            cell = ws_data.cell(row=row_num, column=col)
            cell.fill = fill
            cell.font = NF
            cell.alignment = LA
            cell.border = BD
        ws_data.cell(row=row_num, column=1).fill = fill
        ws_data.cell(row=row_num, column=1).border = BD

        if not url:
            continue
        try:
            img_buf = download_image(url, IMG_SIZE)
            if img_buf:
                xl_img = XLImage(img_buf)
                xl_img.width  = IMG_SIZE
                xl_img.height = IMG_SIZE
                # anchor: 列A（index 0），行 row_num-1（0-indexed）
                cell_addr = "A" + str(row_num)
                ws_data.add_image(xl_img, cell_addr)
                ok_count += 1
        except Exception as e:
            log.debug(f"  图片嵌入失败 row{row_num}: {e}")
            continue

    # 第一行样式
    for cell in ws_data[1]:
        cell.fill = RED
        cell.font = WF
        cell.alignment = CA
        cell.border = BD

    # 其余列宽
    for col in ws_data.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == "A":
            continue
        w = max((len(str(c.value or "")) for c in col), default=8)
        ws_data.column_dimensions[col_letter].width = min(w + 4, 50)

    ws_data.freeze_panes = "B2"
    try:
        wb.save(filepath)
    except PermissionError:
        import time as _t
        alt = filepath.replace(".xlsx", f"_{int(_t.time())}.xlsx")
        wb.save(alt)
        log.warning(f"  原文件被占用，已另存为: {alt}")
        filepath = alt

    log.info(f"Saved: {filepath}  ({len(df)} total rows, {len(df_new)} new, {ok_count} images)")

    # 生成可视化图表
    try:
        _add_charts(filepath, df, log)
    except Exception as e:
        log.warning(f"  图表生成失败（不影响数据）: {e}")

    return filepath


def save_history(new_data, filepath, keyword, log, platform="Taobao"):
    """
    历史数据归档：把每次采集的数据追加到历史文件
    - 文件名：天气丹两件套_历史数据.xlsx
    - 每天一组数据，同一天同平台重复采集则覆盖当天该平台数据
    - 最多保留180天，超出自动清理最早的
    """
    if not new_data:
        return

    import os
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    # 历史文件路径：和主Excel同目录，文件名加"_历史数据"
    # 历史文件名用关键词命名，去掉主文件名里可能有的_data后缀
    base_dir = os.path.dirname(os.path.abspath(filepath))
    safe_kw = keyword.replace("/", "_").replace("\\", "_").strip() or "数据"
    hist_path = os.path.join(base_dir, f"{safe_kw}_历史数据.xlsx")

    # 把新数据转成DataFrame方便处理
    df_new = pd.DataFrame(new_data)
    today_str = pd.to_datetime(df_new["采集时间"].iloc[0], errors="coerce").strftime("%Y-%m-%d")

    # ── 读取或新建历史文件 ──
    if os.path.exists(hist_path):
        try:
            df_hist = pd.read_excel(hist_path, sheet_name="历史记录")
        except Exception:
            df_hist = pd.DataFrame()
    else:
        df_hist = pd.DataFrame()

    # ── 去掉今天同平台的旧数据（覆盖逻辑），追加新数据 ──
    if not df_hist.empty and "采集时间" in df_hist.columns:
        df_hist["_date"] = pd.to_datetime(df_hist["采集时间"], errors="coerce").dt.strftime("%Y-%m-%d")
        df_hist["_plat"] = df_hist.get("店铺类型", "").fillna("").astype(str)
        is_today_same_platform = (df_hist["_date"] == today_str) & (df_hist["_plat"] == platform)
        df_hist = df_hist[~is_today_same_platform].drop(columns=["_date", "_plat"])

    # 只保留核心字段写入历史（不含图片URL，节省体积）
    keep_cols = ["采集时间", "商品标题", "产品规格", "套装类型",
                 "当前售价(元)", "原价(元)", "折扣", "月销量(件)",
                 "累计总销量(估算)", "评价数量", "店铺名称", "店铺类型", "店铺地区", "促销标签"]
    cols_exist = [c for c in keep_cols if c in df_new.columns]
    df_append = df_new[cols_exist].copy()

    df_combined = pd.concat([df_hist, df_append], ignore_index=True)

    # ── 最多保留180天 ──
    if "采集时间" in df_combined.columns:
        df_combined["_date"] = pd.to_datetime(df_combined["采集时间"], errors="coerce").dt.date
        all_dates = sorted(df_combined["_date"].dropna().unique())
        if len(all_dates) > 180:
            cutoff = all_dates[-180]
            df_combined = df_combined[df_combined["_date"] >= cutoff]
        df_combined = df_combined.drop(columns=["_date"])

    # ── 写入Excel ──
    try:
        from openpyxl import Workbook
        wb = Workbook()

        # ── Sheet1：历史记录（完整原始数据）──
        ws_raw = wb.active
        ws_raw.title = "历史记录"

        HDR = PatternFill("solid", fgColor="1a1a2e")
        HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
        BD = Border(
            left=Side(style="thin", color="DDDDDD"),
            right=Side(style="thin", color="DDDDDD"),
            top=Side(style="thin", color="DDDDDD"),
            bottom=Side(style="thin", color="DDDDDD"))
        ALT = PatternFill("solid", fgColor="F8F9FA")

        headers = list(df_combined.columns)
        for ci, h in enumerate(headers, 1):
            cell = ws_raw.cell(1, ci, h)
            cell.fill = HDR
            cell.font = HDR_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BD
        ws_raw.row_dimensions[1].height = 22

        for ri_idx, (_, row) in enumerate(df_combined.iterrows()):
            fill = ALT if ri_idx % 2 == 0 else None
            for ci, h in enumerate(headers, 1):
                cell = ws_raw.cell(ri_idx + 2, ci, row.get(h, ""))
                cell.border = BD
                cell.alignment = Alignment(vertical="center", wrap_text=False)
                if fill:
                    cell.fill = fill

        # 列宽
        col_widths = {"采集时间": 18, "商品标题": 45, "产品规格": 15, "套装类型": 12,
                      "当前售价(元)": 13, "原价(元)": 12, "折扣": 8,
                      "月销量(件)": 12, "累计总销量(估算)": 16, "评价数量": 10,
                      "店铺名称": 20, "店铺类型": 10, "店铺地区": 10, "促销标签": 20}
        for ci, h in enumerate(headers, 1):
            ws_raw.column_dimensions[get_column_letter(ci)].width = col_widths.get(h, 14)

        # ── Sheet2：每日统计摘要 ──
        ws_sum = wb.create_sheet("每日统计")

        sum_headers = ["日期", "采集商品数", "最低价(元)", "最高价(元)", "平均价(元)",
                       "中位价(元)", "月销量合计", "月销量最高", "天猫店数", "淘宝店数", "京东店数", "新增商品数"]
        for ci, h in enumerate(sum_headers, 1):
            cell = ws_sum.cell(1, ci, h)
            cell.fill = HDR
            cell.font = HDR_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BD
        ws_sum.row_dimensions[1].height = 22

        # 按日期分组统计
        df_combined["_date2"] = pd.to_datetime(df_combined["采集时间"], errors="coerce").dt.date
        dates_sorted = sorted(df_combined["_date2"].dropna().unique())

        # 用于新品检测：每天的商品标题集合
        date_titles = {}
        for d in dates_sorted:
            grp = df_combined[df_combined["_date2"] == d]
            date_titles[d] = set(grp["商品标题"].dropna().tolist())

        for ri, d in enumerate(dates_sorted, 2):
            grp = df_combined[df_combined["_date2"] == d]
            prices = pd.to_numeric(grp["当前售价(元)"], errors="coerce").dropna()
            sales = pd.to_numeric(grp["月销量(件)"], errors="coerce").dropna()

            # 新增商品数：和前一天对比
            prev_dates = [x for x in dates_sorted if x < d]
            if prev_dates:
                prev_titles = date_titles[prev_dates[-1]]
                new_count = len(date_titles[d] - prev_titles)
            else:
                new_count = len(date_titles[d])

            row_data = [
                str(d),
                len(grp),
                round(float(prices.min()), 2) if len(prices) else "",
                round(float(prices.max()), 2) if len(prices) else "",
                round(float(prices.mean()), 2) if len(prices) else "",
                round(float(prices.median()), 2) if len(prices) else "",
                int(sales.sum()) if len(sales) else 0,
                int(sales.max()) if len(sales) else 0,
                int((grp["店铺类型"] == "Tmall").sum()),
                int((grp["店铺类型"] == "Taobao").sum()),
                int((grp["店铺类型"] == "JD").sum()),
                new_count,
            ]
            fill = ALT if ri % 2 == 0 else None
            for ci, val in enumerate(row_data, 1):
                cell = ws_sum.cell(ri, ci, val)
                cell.border = BD
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if fill:
                    cell.fill = fill

        sum_widths = [12, 12, 12, 12, 12, 12, 14, 14, 10, 10, 10, 12]
        for ci, w in enumerate(sum_widths, 1):
            ws_sum.column_dimensions[get_column_letter(ci)].width = w

        df_combined = df_combined.drop(columns=["_date2"])

        # ── Sheet3：新品检测 ──
        ws_new = wb.create_sheet("新增竞品")
        new_hdr = ["发现日期", "商品标题", "产品规格", "套装类型",
                   "当前售价(元)", "月销量(件)", "店铺名称", "店铺类型"]
        for ci, h in enumerate(new_hdr, 1):
            cell = ws_new.cell(1, ci, h)
            cell.fill = PatternFill("solid", fgColor="27AE60")
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BD
        ws_new.row_dimensions[1].height = 22

        # 找所有新品：每天和前一天对比，用商品链接做唯一key（标题可能变化）
        all_new_items = []
        df_hist_check = df_combined.copy()
        df_hist_check["_date3"] = pd.to_datetime(
            df_hist_check["采集时间"], errors="coerce").dt.date
        dates_chk = sorted(df_hist_check["_date3"].dropna().unique())

        # 用商品标题前20字作为key（链接历史里不存）
        for idx, d in enumerate(dates_chk):
            grp_today = df_hist_check[df_hist_check["_date3"] == d]
            today_keys = set(grp_today["商品标题"].dropna().str[:20].tolist())
            if idx == 0:
                # 第一天所有商品都算新品
                for _, r in grp_today.iterrows():
                    all_new_items.append({
                        "发现日期": str(d),
                        "商品标题": r.get("商品标题", ""),
                        "产品规格": r.get("产品规格", ""),
                        "套装类型": r.get("套装类型", ""),
                        "当前售价(元)": r.get("当前售价(元)", ""),
                        "月销量(件)": r.get("月销量(件)", ""),
                        "店铺名称": r.get("店铺名称", ""),
                        "店铺类型": r.get("店铺类型", ""),
                    })
            else:
                grp_prev = df_hist_check[df_hist_check["_date3"] == dates_chk[idx-1]]
                prev_keys = set(grp_prev["商品标题"].dropna().str[:20].tolist())
                truly_new = today_keys - prev_keys
                for _, r in grp_today.iterrows():
                    if str(r.get("商品标题", ""))[:20] in truly_new:
                        all_new_items.append({
                            "发现日期": str(d),
                            "商品标题": r.get("商品标题", ""),
                            "产品规格": r.get("产品规格", ""),
                            "套装类型": r.get("套装类型", ""),
                            "当前售价(元)": r.get("当前售价(元)", ""),
                            "月销量(件)": r.get("月销量(件)", ""),
                            "店铺名称": r.get("店铺名称", ""),
                            "店铺类型": r.get("店铺类型", ""),
                        })

        ALT2 = PatternFill("solid", fgColor="F0FFF4")
        for ri, item in enumerate(all_new_items, 2):
            fill = ALT2 if ri % 2 == 0 else None
            for ci, h in enumerate(new_hdr, 1):
                cell = ws_new.cell(ri, ci, item.get(h, ""))
                cell.border = BD
                cell.alignment = Alignment(vertical="center", wrap_text=False)
                if fill:
                    cell.fill = fill

        new_col_widths = [12, 45, 15, 12, 13, 12, 22, 10]
        for ci, w in enumerate(new_col_widths, 1):
            ws_new.column_dimensions[get_column_letter(ci)].width = w

        # 今天新增的商品单独记录，用于日志输出
        today_new = [x for x in all_new_items if x["发现日期"] == today_str]
        if len(dates_sorted) > 1 and today_new:
            log.info(f"  [新增竞品] 今日新增竞品 {len(today_new)} 条: "
                     + " | ".join(x['商品标题'][:15] for x in today_new[:5])
                     + ("..." if len(today_new) > 5 else ""))
        elif len(dates_sorted) > 1:
            log.info(f"  [新增竞品] 今日无新增竞品")

        # ── Sheet4：价格趋势预测 ──
        try:
            ws_trend = wb.create_sheet("价格趋势")

            trend_hdr = ["平台", "店铺名称", "商品标题(截断)", "套装类型",
                         "最新价(元)", "7日均价", "30日均价",
                         "7日趋势", "30日趋势", "综合判断", "数据天数"]
            TEAL  = PatternFill("solid", fgColor="16617A")
            GREEN = PatternFill("solid", fgColor="D5F5E3")
            RED   = PatternFill("solid", fgColor="FADBD8")
            YELL  = PatternFill("solid", fgColor="FEF9E7")

            for ci, h in enumerate(trend_hdr, 1):
                cell = ws_trend.cell(1, ci, h)
                cell.fill = TEAL
                cell.font = Font(bold=True, color="FFFFFF", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = BD
            ws_trend.row_dimensions[1].height = 22

            # 用"店铺名称+标题前15字"做商品唯一key
            df_trend = df_combined.copy()
            df_trend["_date_t"] = pd.to_datetime(
                df_trend["采集时间"], errors="coerce").dt.date
            df_trend["_key"] = (
                df_trend["店铺名称"].fillna("").str.strip() + "||" +
                df_trend["商品标题"].fillna("").str[:15]
            )
            df_trend["_price"] = pd.to_numeric(
                df_trend["当前售价(元)"], errors="coerce")

            today_date = pd.to_datetime(today_str).date()
            cutoff_7  = today_date - __import__("datetime").timedelta(days=7)
            cutoff_30 = today_date - __import__("datetime").timedelta(days=30)

            # 只取今天有数据的商品做趋势分析
            today_keys = set(
                df_trend[df_trend["_date_t"] == today_date]["_key"].tolist())

            trend_rows = []
            for key in today_keys:
                grp = df_trend[df_trend["_key"] == key].sort_values("_date_t")
                grp = grp.dropna(subset=["_price"])
                if grp.empty:
                    continue

                latest_price = float(grp.iloc[-1]["_price"])
                platform = grp.iloc[-1].get("店铺类型", "Taobao")
                shop  = grp.iloc[-1].get("店铺名称", "")
                title = grp.iloc[-1].get("商品标题", "")[:20]
                bundle = grp.iloc[-1].get("套装类型", "")

                # 7日/30日均价
                grp7  = grp[grp["_date_t"] >= cutoff_7]
                grp30 = grp[grp["_date_t"] >= cutoff_30]
                avg7  = round(float(grp7["_price"].mean()), 2) if len(grp7) >= 2 else ""
                avg30 = round(float(grp30["_price"].mean()), 2) if len(grp30) >= 3 else ""
                days  = int(grp["_date_t"].nunique())

                # 趋势判断函数
                def calc_trend(grp_sub, label_days):
                    if len(grp_sub) < 2:
                        return "数据不足"
                    # 按日期取每天最新价
                    daily = grp_sub.groupby("_date_t")["_price"].last()
                    if len(daily) < 2:
                        return "数据不足"
                    first_p = float(daily.iloc[0])
                    last_p  = float(daily.iloc[-1])
                    if first_p == 0:
                        return "数据异常"
                    change_pct = (last_p - first_p) / first_p * 100
                    # 线性回归斜率方向
                    xs = list(range(len(daily)))
                    ys = daily.values.tolist()
                    n  = len(xs)
                    sx = sum(xs); sy = sum(ys)
                    sxy = sum(x*y for x,y in zip(xs,ys))
                    sxx = sum(x*x for x in xs)
                    slope = (n*sxy - sx*sy) / (n*sxx - sx*sx + 1e-9)
                    if change_pct < -5 or slope < -1:
                        return "📉 持续降价"
                    elif change_pct > 5 or slope > 1:
                        return "📈 持续涨价"
                    elif abs(change_pct) <= 2:
                        return "➡️ 价格稳定"
                    else:
                        return "🔄 价格波动"

                trend7  = calc_trend(grp7, 7)  if len(grp7)  >= 2 else "数据不足"
                trend30 = calc_trend(grp30, 30) if len(grp30) >= 3 else "数据不足"

                # 综合判断
                if trend7 == "📉 持续降价" and trend30 == "📉 持续降价":
                    verdict = "⚠️ 长期降价，注意跟进"
                elif trend7 == "📈 持续涨价":
                    verdict = "✅ 近期涨价，竞争压力减小"
                elif trend7 == "📉 持续降价" and trend30 in ("➡️ 价格稳定", "数据不足"):
                    verdict = "🔔 近期突然降价，需关注"
                elif trend7 == "🔄 价格波动":
                    verdict = "🔄 促销活动中"
                else:
                    verdict = "正常"

                trend_rows.append({
                    "平台": platform,
                    "店铺名称": shop,
                    "商品标题(截断)": title,
                    "套装类型": bundle,
                    "最新价(元)": latest_price,
                    "7日均价": avg7,
                    "30日均价": avg30,
                    "7日趋势": trend7,
                    "30日趋势": trend30,
                    "综合判断": verdict,
                    "数据天数": days,
                    "_verdict": verdict,
                })

            # 按综合判断排序：警告类在前
            priority = {"⚠️ 长期降价，注意跟进": 0, "🔔 近期突然降价，需关注": 1,
                        "🔄 促销活动中": 2, "✅ 近期涨价，竞争压力减小": 3, "正常": 4}
            trend_rows.sort(key=lambda r: priority.get(r["_verdict"], 9))

            for ri, row in enumerate(trend_rows, 2):
                verdict = row["_verdict"]
                if "长期降价" in verdict:
                    row_fill = RED
                elif "突然降价" in verdict:
                    row_fill = YELL
                elif "涨价" in verdict:
                    row_fill = GREEN
                else:
                    row_fill = None

                for ci, h in enumerate(trend_hdr, 1):
                    cell = ws_trend.cell(ri, ci, row.get(h, ""))
                    cell.border = BD
                    cell.alignment = Alignment(vertical="center", wrap_text=False)
                    if row_fill:
                        cell.fill = row_fill

            trend_widths = [10, 22, 28, 12, 12, 10, 10, 14, 14, 22, 10]
            for ci, w in enumerate(trend_widths, 1):
                ws_trend.column_dimensions[get_column_letter(ci)].width = w

            warn_cnt = sum(1 for r in trend_rows if r["_verdict"] in
                           ("⚠️ 长期降价，注意跟进", "🔔 近期突然降价，需关注"))
            log.info(f"  [趋势] 价格趋势分析完成: {len(trend_rows)} 个商品"
                     + (f"，其中 {warn_cnt} 个需关注" if warn_cnt else ""))

        except Exception as _te:
            log.warning(f"  [趋势] 趋势分析失败: {_te}")

        # 保存
        try:
            wb.save(hist_path)
            log.info(f"  [历史] 已归档到: {os.path.basename(hist_path)}  "
                     f"(共{len(df_combined)}条 / {len(dates_sorted)}天)")
        except PermissionError:
            alt = hist_path.replace(".xlsx", f"_{int(__import__('time').time())}.xlsx")
            wb.save(alt)
            log.warning(f"  [历史] 文件被占用，另存为: {os.path.basename(alt)}")

    except Exception as e:
        log.warning(f"  [历史] 归档失败: {e}")



def _add_charts(filepath, df, log):
    """在Excel里生成数据可视化图表sheet"""
    from openpyxl import load_workbook
    from openpyxl.chart import BarChart, PieChart, ScatterChart, Reference, Series
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter

    wb = load_workbook(filepath)

    # ── 准备数据 ──
    price_col = "当前售价(元)" if "当前售价(元)" in df.columns else "当前售价"
    sales_col = "月销量(件)"   if "月销量(件)"   in df.columns else "月销量"
    title_col = "商品标题"     if "商品标题"     in df.columns else "标题"
    shop_col  = "店铺类型"     if "店铺类型"     in df.columns else "店铺名称"

    df2 = df.copy()
    df2[price_col] = pd.to_numeric(df2[price_col], errors="coerce").fillna(0)
    df2[sales_col] = pd.to_numeric(df2[sales_col], errors="coerce").fillna(0)

    # ── 创建/重置图表数据sheet ──
    chart_sheet = "数据图表"
    if chart_sheet in wb.sheetnames:
        del wb[chart_sheet]
    ws = wb.create_sheet(chart_sheet)

    HDR = PatternFill("solid", fgColor="2A5298")
    WF  = Font(bold=True, color="FFFFFF", name="微软雅黑", size=10)
    NF  = Font(name="微软雅黑", size=9)
    CA  = Alignment(horizontal="center", vertical="center")

    def write_header(ws, row, col, text):
        c = ws.cell(row=row, column=col, value=text)
        c.fill = HDR; c.font = WF; c.alignment = CA

    # ══ 图1：价格分布柱状图 ══
    # 计算价格区间
    p_max = df2[price_col].max()
    if p_max <= 0: p_max = 500
    step  = max(50, round(p_max / 8 / 50) * 50)
    bins  = list(range(0, int(p_max) + step, step))
    labels = [f"¥{bins[i]}-{bins[i+1]}" for i in range(len(bins)-1)]
    counts = []
    for i in range(len(bins)-1):
        cnt = int(((df2[price_col] >= bins[i]) & (df2[price_col] < bins[i+1])).sum())
        counts.append(cnt)

    ws.cell(row=1, column=1, value="📊 价格分布").font = Font(bold=True, size=12, name="微软雅黑", color="2A5298")
    write_header(ws, 2, 1, "价格区间")
    write_header(ws, 2, 2, "商品数量")
    for i, (label, cnt) in enumerate(zip(labels, counts)):
        ws.cell(row=3+i, column=1, value=label).font = NF
        ws.cell(row=3+i, column=2, value=cnt).font = NF

    chart1 = BarChart()
    chart1.type = "col"
    chart1.title = "价格分布"
    chart1.y_axis.title = "商品数量"
    chart1.x_axis.title = "价格区间"
    chart1.style = 10
    chart1.width = 18; chart1.height = 12
    data1   = Reference(ws, min_col=2, min_row=2, max_row=2+len(labels))
    cats1   = Reference(ws, min_col=1, min_row=3, max_row=2+len(labels))
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.series[0].graphicalProperties.solidFill = "2A5298"
    ws.add_chart(chart1, "D2")
    log.info("  [图表] 价格分布柱状图 ✅")

    # ══ 图2：销量 Top10 条形图 ══
    top10 = df2.nlargest(10, sales_col)[[title_col, sales_col]].reset_index(drop=True)
    r2 = 22
    ws.cell(row=r2, column=1, value="📈 销量 Top10").font = Font(bold=True, size=12, name="微软雅黑", color="27AE60")
    write_header(ws, r2+1, 1, "商品标题（截断）")
    write_header(ws, r2+1, 2, "月销量")
    for i, row in top10.iterrows():
        short_title = str(row[title_col])[:15] + "..."
        ws.cell(row=r2+2+i, column=1, value=short_title).font = NF
        ws.cell(row=r2+2+i, column=2, value=int(row[sales_col])).font = NF

    chart2 = BarChart()
    chart2.type = "bar"   # 横向
    chart2.title = "销量 Top10"
    chart2.x_axis.title = "月销量（件）"
    chart2.style = 10
    chart2.width = 20; chart2.height = 13
    data2 = Reference(ws, min_col=2, min_row=r2+1, max_row=r2+1+len(top10))
    cats2 = Reference(ws, min_col=1, min_row=r2+2, max_row=r2+1+len(top10))
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.series[0].graphicalProperties.solidFill = "27AE60"
    ws.add_chart(chart2, "D22")
    log.info("  [图表] 销量Top10条形图 ✅")

    # ══ 图3：店铺类型饼图 ══
    r3 = 46
    if "店铺类型" in df2.columns:
        type_counts = df2["店铺类型"].value_counts().head(5)
    else:
        type_counts = df2[shop_col].value_counts().head(5) if shop_col in df2.columns else pd.Series({"未知": len(df2)})

    ws.cell(row=r3, column=1, value="🥧 店铺类型分布").font = Font(bold=True, size=12, name="微软雅黑", color="E67E22")
    write_header(ws, r3+1, 1, "店铺类型")
    write_header(ws, r3+1, 2, "数量")
    for i, (k, v) in enumerate(type_counts.items()):
        ws.cell(row=r3+2+i, column=1, value=str(k)).font = NF
        ws.cell(row=r3+2+i, column=2, value=int(v)).font = NF

    chart3 = PieChart()
    chart3.title = "店铺类型分布"
    chart3.style = 10
    chart3.width = 16; chart3.height = 12
    data3 = Reference(ws, min_col=2, min_row=r3+1, max_row=r3+1+len(type_counts))
    cats3 = Reference(ws, min_col=1, min_row=r3+2, max_row=r3+1+len(type_counts))
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    ws.add_chart(chart3, "D46")
    log.info("  [图表] 店铺类型饼图 ✅")

    # ══ 图4：价格-销量散点图 ══
    r4 = 66
    sample = df2[df2[price_col] > 0][df2[sales_col] > 0][[price_col, sales_col]].head(60)
    ws.cell(row=r4, column=1, value="📉 价格-销量关系").font = Font(bold=True, size=12, name="微软雅黑", color="8E44AD")
    write_header(ws, r4+1, 1, "售价(元)")
    write_header(ws, r4+1, 2, "月销量")
    for i, (_, row) in enumerate(sample.iterrows()):
        ws.cell(row=r4+2+i, column=1, value=round(float(row[price_col]), 1)).font = NF
        ws.cell(row=r4+2+i, column=2, value=int(row[sales_col])).font = NF

    chart4 = ScatterChart()
    chart4.title = "价格与销量关系"
    chart4.x_axis.title = "售价（元）"
    chart4.y_axis.title = "月销量（件）"
    chart4.style = 10
    chart4.width = 18; chart4.height = 12
    xvals = Reference(ws, min_col=1, min_row=r4+2, max_row=r4+1+len(sample))
    yvals = Reference(ws, min_col=2, min_row=r4+2, max_row=r4+1+len(sample))
    series4 = Series(yvals, xvals, title="商品")
    series4.marker.symbol = "circle"
    series4.marker.size   = 5
    series4.marker.graphicalProperties.solidFill  = "8E44AD"
    series4.marker.graphicalProperties.line.solidFill = "8E44AD"
    series4.graphicalProperties.line.noFill = True
    chart4.series.append(series4)
    ws.add_chart(chart4, "D66")
    log.info("  [图表] 价格-销量散点图 ✅")

    # 调整列宽
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 12

    wb.save(filepath)
    log.info("  [图表] 数据图表sheet已写入Excel ✅")


def human_type(element, text, stop_event=None):
    for ch in text:
        if stop_event and stop_event.is_set():
            return
        element.send_keys(ch)
        time.sleep(random.uniform(0.05, 0.18))


def human_search(driver, keyword, exact=False, log=None, stop_event=None):
    if log is None:
        log = logging.getLogger()
    log.info("  切换到淘宝页面并搜索...")

    # ── Step 1: Scan ALL tabs, find the best Taobao tab ──
    search_handle = None   # s.taobao.com/search page (best)
    taobao_handle = None   # any taobao.com page (fallback)

    all_handles = driver.window_handles
    log.info(f"  共 {len(all_handles)} 个标签页，逐个检查...")

    for handle in all_handles:
        try:
            driver.switch_to.window(handle)
            time.sleep(0.3)  # Give tab time to activate
            url = safe_url(driver)
            title = safe_title(driver)
            log.info(f"    标签: {title[:30]} | {url[:60]}")

            url_lower = url.lower()
            if "s.taobao.com" in url_lower:
                search_handle = handle
                log.info(f"    → 这是淘宝搜索结果页 ✓")
            elif "taobao.com" in url_lower or "tmall.com" in url_lower:
                if taobao_handle is None:
                    taobao_handle = handle
                    log.info(f"    → 这是淘宝页面 ✓")
        except Exception as e:
            log.info(f"    标签切换失败: {e}")
            continue

    # ── Step 2: 找到搜索结果页，检查关键词是否匹配 ──
    if search_handle:
        driver.switch_to.window(search_handle)
        keep_minimized(driver)
        time.sleep(0.5)
        cur_url = safe_url(driver)
        log.info(f"  切换到已有的淘宝搜索结果页: {cur_url[:70]}")

        # 从 URL 的 q= 参数提取当前搜索关键词
        from urllib.parse import urlparse, parse_qs, unquote
        try:
            parsed_qs = parse_qs(urlparse(cur_url).query)
            url_keyword = parsed_qs.get("q", [""])[0]
        except Exception:
            url_keyword = ""

        keyword_match = keyword in url_keyword or url_keyword in keyword
        if not keyword_match:
            log.info(f"  搜索结果页关键词不匹配（页面: '{url_keyword}', 目标: '{keyword}'），将重新搜索...")
            # 关键词不匹配，把这个标签当作淘宝标签来重新搜索
            taobao_handle = search_handle
            search_handle = None
        else:
            # 关键词匹配，检查是否有商品卡片
            try:
                cards = driver.find_elements(By.CSS_SELECTOR,
                    "[class*='doubleCardWrapper'], [class*='card--doubleCardWrapper'], "
                    "[class*='itemWrapper'], div[data-item-id], [class*='Item--doubleCardWrapper'],"
                    "[class*='CardPC'], [class*='cardPC']")
                if cards:
                    log.info(f"  ✅ 已在搜索结果页({len(cards)}个商品)，关键词匹配，直接采集")
                    keep_minimized(driver)
                    return True
                else:
                    log.info("  搜索页无商品卡片，将重新搜索...")
            except Exception:
                log.info("  ✅ 已在搜索结果页，关键词匹配，直接采集")
                keep_minimized(driver)
                return True

    # ── Step 3: 有淘宝标签但不是搜索页 → 在该标签内搜索 ──
    if taobao_handle:
        driver.switch_to.window(taobao_handle)
        keep_minimized(driver)
        log.info(f"  切换到淘宝标签进行搜索...")
    else:
        # 没有任何淘宝标签，在当前标签打开淘宝
        log.info("  没找到淘宝标签，在当前标签打开淘宝...")
        driver.get("https://www.taobao.com")
        keep_minimized(driver)
        rsleep(3, 5, stop_event)

    rsleep(2, 3, stop_event)
    if stop_event and stop_event.is_set():
        return False

    try:
        body = driver.find_element(By.TAG_NAME, "body")
        ac = ActionChains(driver)
        for _ in range(random.randint(2, 3)):
            ac.move_to_element_with_offset(body, random.randint(200, 700), random.randint(100, 400))
            ac.pause(random.uniform(0.3, 0.6))
        ac.perform()
        rsleep(0.5, 1, stop_event)
    except Exception:
        pass

    search_box = None
    for sel in ["#q", "input[name='q']", ".search-combobox-input",
                "[class*='searchInput']", "input[type='search']", "[class*='search-input']"]:
        try:
            els = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in els:
                if el.is_displayed() and el.is_enabled():
                    search_box = el
                    log.info(f"  找到搜索框: {sel}")
                    break
            if search_box:
                break
        except Exception:
            continue

    if not search_box:
        log.warning("  找不到搜索框")
        return False

    driver.execute_script("arguments[0].scrollIntoView(true);", search_box)
    rsleep(0.5, 1, stop_event)

    try:
        ActionChains(driver).move_to_element(search_box).pause(random.uniform(0.4, 0.8)).click().perform()
    except Exception:
        driver.execute_script("arguments[0].click();", search_box)
    rsleep(0.5, 1, stop_event)

    search_box.send_keys(Keys.CONTROL + "a")
    rsleep(0.2, 0.4, stop_event)
    search_box.send_keys(Keys.DELETE)
    rsleep(0.3, 0.5, stop_event)

    q = f'"{keyword}"' if exact else keyword
    human_type(search_box, q, stop_event)
    rsleep(0.8, 1.5, stop_event)

    actual = search_box.get_attribute("value") or ""
    if keyword not in actual:
        log.warning("  输入验证失败，强制写入...")
        driver.execute_script("arguments[0].value = arguments[1];", search_box, q)
        rsleep(0.5, 1, stop_event)

    clicked = False
    for btn_sel in [".btn-search", "[class*='searchBtn']", "button[type='submit']",
                    "[class*='search-button']", ".search-btn", "button.btn"]:
        try:
            btns = driver.find_elements(By.CSS_SELECTOR, btn_sel)
            for btn in btns:
                if btn.is_displayed():
                    rsleep(0.3, 0.7, stop_event)
                    ActionChains(driver).move_to_element(btn).pause(random.uniform(0.2, 0.5)).click().perform()
                    clicked = True
                    log.info("  点击搜索按钮")
                    break
            if clicked:
                break
        except Exception:
            continue

    if not clicked:
        search_box.send_keys(Keys.RETURN)
        log.info("  按回车搜索")

    log.info(f"  已搜索: {q}")

    handles_before = set(driver.window_handles)

    for i in range(25):
        if stop_event and stop_event.is_set():
            return False
        time.sleep(1)
        # 检查当前标签页
        cur = safe_url(driver)
        if "s.taobao.com" in cur and "deny" not in cur and "punish" not in cur:
            log.info(f"  搜索成功: {cur[:80]}")
            keep_minimized(driver)
            return True
        if "deny_h5" in cur or "punish" in cur:
            log.warning("  搜索后被拦截，等待30秒...")
            for _ in range(30):
                if stop_event and stop_event.is_set():
                    return False
                time.sleep(1)
            return False
        # 检查是否在新标签页打开了搜索结果
        new_handles = set(driver.window_handles) - handles_before
        for nh in new_handles:
            try:
                driver.switch_to.window(nh)
                new_url = safe_url(driver)
                if "s.taobao.com" in new_url and "deny" not in new_url:
                    log.info(f"  搜索成功（新标签页）: {new_url[:80]}")
                    keep_minimized(driver)
                    return True
            except Exception:
                pass

    log.warning(f"  等待超时: {safe_url(driver)[:80]}")
    return False


def apply_sort(driver, sort_by, tmall_only, log, stop_event=None):
    try:
        for _ in range(20):
            if stop_event and stop_event.is_set():
                return
            if "s.taobao.com" in safe_url(driver):
                break
            time.sleep(1)
        else:
            log.warning(f"  未到达搜索结果页，跳过排序")
            return
        current_url = safe_url(driver)
        sort_map = {"sale": "sale-desc", "price_asc": "price-asc", "price_desc": "price-desc"}
        # 使用 urllib 正确解析和替换URL参数
        from urllib.parse import urlparse, parse_qs, urlencode, urlunparse
        parsed = urlparse(current_url)
        qs = parse_qs(parsed.query, keep_blank_values=True)
        changed = False
        if sort_by in sort_map:
            new_sort = sort_map[sort_by]
            if qs.get("sort", [None])[0] != new_sort:
                qs["sort"] = [new_sort]
                changed = True
        if tmall_only and qs.get("tab", [None])[0] != "mall":
            qs["tab"] = ["mall"]
            changed = True
        if changed:
            new_query = urlencode({k: v[0] for k, v in qs.items()})
            new_url = urlunparse(parsed._replace(query=new_query))
            driver.get(new_url)
            keep_minimized(driver)
            rsleep(2, 4, stop_event)
    except Exception as e:
        log.warning(f"  排序设置失败: {e}")


def goto_page(driver, page_i, log, stop_event=None):
    if page_i == 0:
        return True
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        rsleep(1.5, 2.5, stop_event)
        next_sels = [
            "a.next", "[class*='next']",
            "button[aria-label*='下一页']", "a[aria-label*='下一页']",
            "[class*='pagination'] [class*='next']",
            "[class*='page-next']", ".page-next", "[title='下一页']",
        ]
        for sel in next_sels:
            try:
                btns = driver.find_elements(By.CSS_SELECTOR, sel)
                for btn in btns:
                    if btn.is_displayed() and btn.is_enabled():
                        driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                        rsleep(0.5, 1, stop_event)
                        ActionChains(driver).move_to_element(btn).pause(
                            random.uniform(0.3, 0.6)).click().perform()
                        log.info(f"  点击翻页按钮: {sel}")
                        rsleep(3, 5, stop_event)
                        return True
            except Exception:
                continue

        log.info("  未找到翻页按钮，尝试输入页码...")
        page_inputs = driver.find_elements(By.CSS_SELECTOR,
            "input[class*='page'], input[aria-label*='页'], [class*='pagination'] input")
        for inp in page_inputs:
            if inp.is_displayed():
                inp.clear()
                rsleep(0.3, 0.5, stop_event)
                human_type(inp, str(page_i + 1), stop_event)
                rsleep(0.3, 0.5, stop_event)
                inp.send_keys(Keys.RETURN)
                log.info(f"  输入页码: {page_i + 1}")
                rsleep(3, 5, stop_event)
                return True

        log.warning("  翻页失败")
        return False
    except Exception as e:
        log.warning(f"  翻页异常: {e}")
        return False


def run_round(driver, cfg, round_num, log, stop_event=None):
    keep_minimized(driver)
    mode_label = "精确" if cfg.get("exact") else "模糊"
    log.info(f"\n{'=' * 55}")
    log.info(f"  Round {round_num}  |  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log.info(f"  关键词: {cfg['keyword']} ({mode_label})  |  Pages: {cfg['pages']}")
    log.info(f"{'=' * 55}")

    all_items = []
    # 整轮采集用同一个时间戳，避免多页采集时间不同导致汇总重复行
    round_now = datetime.now().strftime("%Y-%m-%d %H:%M")

    ok = human_search(driver, cfg["keyword"], cfg.get("exact", False), log, stop_event)
    keep_minimized(driver)
    if not ok:
        log.warning("  首页搜索失败，尝试备用方式...")
        # 备用方式：在当前（新）标签页中直接跳转
        driver.get("https://www.taobao.com")
        keep_minimized(driver)
        rsleep(2, 3, stop_event)
        sort_map = {"sale": "sale-desc", "default": "default",
                    "price_asc": "price-asc", "price_desc": "price-desc"}
        q = f'"{cfg["keyword"]}"' if cfg.get("exact") else cfg["keyword"]
        fallback_url = (f"https://s.taobao.com/search?q={quote(q)}"
                        f"&sort={sort_map.get(cfg['sort_by'], 'default')}"
                        f"{'&tab=mall' if cfg['tmall_only'] else ''}")
        driver.get(fallback_url)
        keep_minimized(driver)
        rsleep(4, 7, stop_event)

    for _ in range(15):
        if stop_event and stop_event.is_set():
            return 0
        if "s.taobao.com" in safe_url(driver):
            break
        time.sleep(1)

    if "deny_h5" in safe_url(driver) or "punish" in safe_url(driver):
        log.warning("  被拦截！等待60秒后重试...")
        for _ in range(60):
            if stop_event and stop_event.is_set():
                return 0
            time.sleep(1)
        driver.get("https://www.taobao.com")
        keep_minimized(driver)
        rsleep(5, 8, stop_event)
        human_search(driver, cfg["keyword"], cfg.get("exact", False), log, stop_event)
        keep_minimized(driver)

    log.info(f"  搜索结果页: {safe_url(driver)[:70]}")
    apply_sort(driver, cfg["sort_by"], cfg["tmall_only"], log, stop_event)
    keep_minimized(driver)

    for page_i in range(cfg["pages"]):
        if stop_event and stop_event.is_set():
            break

        if page_i > 0:
            log.info(f"\n  翻到第 {page_i + 1} 页...")
            goto_page(driver, page_i, log, stop_event)

            cur_url = safe_url(driver)
            if "deny_h5" in cur_url or "punish" in cur_url:
                log.warning("  翻页时被拦截，等待30秒...")
                for _ in range(30):
                    if stop_event and stop_event.is_set():
                        break
                    time.sleep(1)
                # 重新搜索而非跳回首页（首页没有翻页按钮）
                human_search(driver, cfg["keyword"], cfg.get("exact", False), log, stop_event)
                rsleep(2, 4, stop_event)
                goto_page(driver, page_i, log, stop_event)
                keep_minimized(driver)

        log.info(f"\n  第 {page_i + 1}/{cfg['pages']} 页")
        items = scrape_page(driver, cfg, log, stop_event, now=round_now)
        log.info(f"  -> {len(items)} 条商品")
        all_items.extend(items)

        if page_i < cfg["pages"] - 1 and not (stop_event and stop_event.is_set()):
            rsleep(5, 12, stop_event)

    actual_path = save_excel(all_items, cfg["output_file"], log)
    if actual_path:
        cfg["output_file"] = actual_path
    # 历史数据归档
    try:
        save_history(all_items, cfg["output_file"], cfg.get("keyword", ""), log)
    except Exception as _he:
        log.warning(f"  [历史] 归档异常: {_he}")

    return len(all_items)


# ══════════════════════════════════════════════════════════
#  第二阶段：多平台采集 + 评价分析
# ══════════════════════════════════════════════════════════


def scrape_hot_keywords(driver, keyword, log, stop_event=None):
    """从淘宝和京东抓取热门搜索联想词"""
    keep_minimized(driver)
    import json, re as _re
    results = {"淘宝": [], "京东": []}

    # ── 淘宝热词：先尝试搜索框输入触发下拉，再fallback到API ──
    try:
        log.info("  [热词] 采集淘宝搜索联想词...")
        tb_words = []

        # 方式1：切换到淘宝标签页，输入搜索框触发下拉联想
        try:
            from selenium.webdriver.common.keys import Keys as _Keys
            # 找淘宝标签页
            tb_handle = None
            for h in driver.window_handles:
                try:
                    driver.switch_to.window(h)
                    if "taobao.com" in safe_url(driver).lower():
                        tb_handle = h; break
                except Exception: pass

            if tb_handle:
                driver.switch_to.window(tb_handle)
                keep_minimized(driver)
                time.sleep(1)
                log.info(f"  [热词] 淘宝标签页URL: {safe_url(driver)[:60]}")
                # 找搜索框 - 先滚到顶部让搜索框可见
                try:
                    driver.execute_script("window.scrollTo(0,0);")
                    time.sleep(0.5)
                except Exception: pass
                box = None
                for sel in ["#q", "input[name='q']", ".search-combobox-input",
                            "input.search-combobox-input", "[class*='searchbar'] input",
                            "[class*='search-bar'] input", "input[type='text']"]:
                    try:
                        el = driver.find_element(By.CSS_SELECTOR, sel)
                        if el.is_displayed():
                            box = el; break
                    except Exception: pass

                if box:
                    log.info(f"  [热词] 淘宝搜索框找到，开始输入...")
                    try:
                        # Scroll to very top to make search box visible
                        driver.execute_script("window.scrollTo(0,0);")
                        time.sleep(1)
                        # Use JS click to avoid element interception by nav bar
                        driver.execute_script("arguments[0].click();", box)
                        time.sleep(0.3)
                        driver.execute_script("arguments[0].focus();", box)
                        time.sleep(0.2)
                        # 全选清空
                        box.send_keys(_Keys.CONTROL + 'a')
                        time.sleep(0.1)
                        box.send_keys(_Keys.DELETE)
                        time.sleep(0.3)
                        # 逐字输入
                        for ch in keyword:
                            box.send_keys(ch)
                            time.sleep(0.18)
                        time.sleep(3)

                        # 扫描可视区域所有文字元素，取含关键词的联想词
                        all_visible = driver.execute_script("""
                            var found = [];
                            var all = document.querySelectorAll('div,span,a,li,p');
                            for(var i=0;i<all.length;i++){
                                var el=all[i];
                                var rect=el.getBoundingClientRect();
                                if(rect.top>50 && rect.top<500 && rect.height>5 && rect.height<60){
                                    var t=(el.innerText||el.textContent||'').trim();
                                    if(t.length>1 && t.length<35 && found.indexOf(t)<0) found.push(t);
                                }
                            }
                            return found.slice(0,50);
                        """)
                        kw_core = keyword[:2]
                        for w in (all_visible or []):
                            if kw_core in w and w not in tb_words:
                                tb_words.append(w)
                        if tb_words:
                            log.info(f"  [热词] 淘宝抓取成功: {len(tb_words)} 个")
                    except Exception as _je:
                        log.warning(f"  [热词] 淘宝输入失败: {_je}")
                    try:
                        box.send_keys(_Keys.ESCAPE)
                    except Exception: pass
        except Exception as _e:
            log.warning(f"  [热词] 淘宝搜索框方式失败: {_e}")

        # 方式2：JSONP注入方式获取淘宝suggest
        if not tb_words:
            try:
                kw_enc = quote(keyword)
                src = driver.execute_script(f"""
                    return new Promise(function(resolve) {{
                        window.__tb_suggest_cb__ = function(d) {{
                            window.__tb_suggest_result__ = d;
                        }};
                        var s = document.createElement('script');
                        s.src = 'https://suggest.taobao.com/sug?code=utf-8&q={kw_enc}&callback=__tb_suggest_cb__';
                        s.onload = function() {{ resolve(JSON.stringify(window.__tb_suggest_result__ || {{}})); }};
                        s.onerror = function() {{ resolve(''); }};
                        document.head.appendChild(s);
                        setTimeout(function(){{ resolve(JSON.stringify(window.__tb_suggest_result__ || {{}})); }}, 3000);
                    }});
                """)
                log.info(f"  [热词] 淘宝JSONP返回: {(src or '')[:80]}")
                if src and src.strip().startswith("{"):
                    import json as _json
                    data = _json.loads(src)
                    raw = data.get("result", [])
                    if raw and isinstance(raw[0], list):
                        tb_words = [r[0] for r in raw if r and len(r[0]) > 1]
                    elif raw and isinstance(raw[0], str):
                        tb_words = [w for w in raw if len(w) > 1]
                    if tb_words:
                        log.info(f"  [热词] 淘宝JSONP成功: {len(tb_words)} 个")
            except Exception as _pe:
                log.warning(f"  [热词] 淘宝JSONP失败: {_pe}")

        results["淘宝"] = tb_words[:20]
        if not results["淘宝"]:
            log.warning("  [热词] 淘宝未采集到联想词")
        else:
            log.info(f"  [热词] 淘宝: {len(results['淘宝'])} 个词")
    except Exception as e:
        log.warning(f"  [热词] 淘宝异常: {e}")

    if stop_event and stop_event.is_set():
        return results

    # ── 京东热词：切换到已有的京东标签页操作 ──
    try:
        log.info("  [热词] 采集京东搜索联想词...")
        jd_words = []

        # 找已有的京东标签（search.jd.com），切过去
        try:
            jd_handle = None
            for h in driver.window_handles:
                try:
                    driver.switch_to.window(h)
                    if "jd.com" in safe_url(driver).lower():
                        jd_handle = h; break
                except Exception: pass

            if jd_handle:
                driver.switch_to.window(jd_handle)
                keep_minimized(driver)
                log.info(f"  [热词] 已切换到京东标签: {safe_url(driver)[:60]}")
                # If not on search page, navigate within JD tab
                if "search.jd.com" not in safe_url(driver).lower():
                    search_url = f"https://search.jd.com/Search?keyword={quote(keyword)}&enc=utf-8"
                    driver.get(search_url)
                    keep_minimized(driver)
                    time.sleep(3)
            else:
                # No JD tab - navigate current tab to JD search
                log.info("  [热词] 未找到京东标签，在当前标签打开...")
                search_url = f"https://search.jd.com/Search?keyword={quote(keyword)}&enc=utf-8"
                driver.get(search_url)
                keep_minimized(driver)
                time.sleep(3)
                log.info(f"  [热词] 已打开京东搜索页")
        except Exception: pass

        # 方式1：在京东搜索框输入关键词触发下拉联想
        try:
            from selenium.webdriver.common.keys import Keys
            time.sleep(2)  # 等页面加载完成
            log.info(f"  [热词] 京东当前URL: {safe_url(driver)[:60]}")
            box = None
            for sel in ["#key", "input[name='keyword']", ".search-input input",
                        "[class*='search'] input[type='text']",
                        "input.input", "form input[type='text']",
                        "input[type='text']"]:
                try:
                    el = driver.find_element(By.CSS_SELECTOR, sel)
                    if el.is_displayed():
                        box = el
                        log.info(f"  [热词] 京东搜索框找到: {sel}")
                        break
                except Exception: pass

            if box:
                log.info(f"  [热词] 京东搜索框已找到，开始输入...")
                try:
                    # 点击激活
                    box.click()
                    time.sleep(0.3)
                    driver.execute_script("arguments[0].focus();", box)
                    time.sleep(0.2)
                    # 全选清空
                    box.send_keys(Keys.CONTROL + 'a')
                    time.sleep(0.1)
                    box.send_keys(Keys.DELETE)
                    time.sleep(0.3)
                    # 逐字输入
                    for ch in keyword:
                        box.send_keys(ch)
                        time.sleep(0.18)
                    time.sleep(3)

                    # 扫描可视区域所有文字元素，取含关键词的联想词
                    all_visible = driver.execute_script("""
                        var found = [];
                        var all = document.querySelectorAll('div,span,a,li,p');
                        for(var i=0;i<all.length;i++){
                            var el=all[i];
                            var rect=el.getBoundingClientRect();
                            if(rect.top>50 && rect.top<500 && rect.height>5 && rect.height<60){
                                var t=(el.innerText||el.textContent||'').trim();
                                if(t.length>1 && t.length<35 && found.indexOf(t)<0) found.push(t);
                            }
                        }
                        return found.slice(0,50);
                    """)
                    kw_core = keyword[:2]
                    for w in (all_visible or []):
                        if kw_core in w and w not in jd_words:
                            jd_words.append(w)
                    if jd_words:
                        log.info(f"  [热词] 京东抓取成功: {len(jd_words)} 个")
                except Exception as _je:
                    log.warning(f"  [热词] 京东输入失败: {_je}")
                try:
                    box.send_keys(Keys.ESCAPE)
                except Exception: pass
            else:
                log.warning("  [热词] 京东未找到搜索框")
        except Exception as _e:
            log.warning(f"  [热词] 京东搜索框方式失败: {_e}")

        # 方式2：注入JS动态脚本获取京东suggest（规避CORS限制）
        if not jd_words:
            try:
                kw_enc = quote(keyword)
                src = driver.execute_script(f"""
                    return new Promise(function(resolve) {{
                        var s = document.createElement('script');
                        window.__jd_suggest_cb__ = function(d) {{
                            window.__jd_suggest_result__ = d;
                        }};
                        s.src = 'https://suggest.jd.com/su?keyword={kw_enc}&enc=utf-8&area=1&callback=__jd_suggest_cb__';
                        s.onload = function() {{ resolve(JSON.stringify(window.__jd_suggest_result__ || {{}})); }};
                        s.onerror = function() {{ resolve(''); }};
                        document.head.appendChild(s);
                        setTimeout(function(){{ resolve(JSON.stringify(window.__jd_suggest_result__ || {{}})); }}, 3000);
                    }});
                """)
                log.info(f"  [热词] 京东JSONP返回: {(src or '')[:80]}")
                if src and src.strip().startswith("{"):
                    data = json.loads(src)
                    result = data.get("result", {})
                    if isinstance(result, dict):
                        for grp in result.get("groupList", []):
                            for w in grp.get("data", []):
                                if isinstance(w, str) and len(w) > 1 and w not in jd_words:
                                    jd_words.append(w)
                    elif isinstance(result, list):
                        jd_words = [w for w in result if isinstance(w, str) and len(w) > 1]
                    if jd_words:
                        log.info(f"  [热词] 京东JSONP成功: {len(jd_words)} 个")
            except Exception as _e:
                log.warning(f"  [热词] 京东JSONP失败: {_e}")

        results["京东"] = jd_words[:20]
        log.info(f"  [热词] 京东: {len(results['京东'])} 个词")
    except Exception as e:
        log.warning(f"  [热词] 京东异常: {e}")

    return results

def save_hotwords_to_excel(filepath, keyword, hotwords, log):
    """将热词写入Excel的"热门关键词"sheet"""
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = load_workbook(filepath)

        # 删旧sheet
        if "热门关键词" in wb.sheetnames:
            del wb["热门关键词"]
        ws = wb.create_sheet("热门关键词")

        RED = PatternFill("solid", fgColor="C0392B")
        GRN = PatternFill("solid", fgColor="1A5276")
        WHT = PatternFill("solid", fgColor="FFFFFF")
        GRAY= PatternFill("solid", fgColor="F5F5F5")
        WF  = Font(bold=True, color="FFFFFF", size=10, name="Arial")
        NF  = Font(size=9, name="Arial")
        CA  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        LA  = Alignment(horizontal="left",   vertical="center", wrap_text=True)
        BD  = Border(
            left=Side(style="thin", color="DDDDDD"), right=Side(style="thin", color="DDDDDD"),
            top=Side(style="thin", color="DDDDDD"),  bottom=Side(style="thin", color="DDDDDD"))

        _HW_TT = {
            "排名":           "搜索热度排名，数字越小越热门",
            "淘宝热门关键词": "淘宝搜索框联想词，代表买家实际搜索习惯",
            "京东热门关键词": "京东搜索框联想词，反映京东平台买家的搜索偏好",
        }
        def hdr(row, col, val, fill=RED):
            from openpyxl.comments import Comment as _Cmt
            c = ws.cell(row=row, column=col, value=val)
            c.fill = fill; c.font = WF; c.alignment = CA; c.border = BD
            if val in _HW_TT:
                cm = _Cmt(_HW_TT[val], "📋 列说明")
                cm.width = 260; cm.height = 55
                c.comment = cm

        def cell(row, col, val, is_gray=False):
            c = ws.cell(row=row, column=col, value=val)
            c.fill = GRAY if is_gray else WHT
            c.font = NF; c.alignment = LA; c.border = BD

        # 标题行
        ws.merge_cells("A1:C1")
        title_cell = ws.cell(row=1, column=1, value=f"🔥 热门搜索关键词  —  搜索词：{keyword}")
        title_cell.fill = PatternFill("solid", fgColor="1B2631")
        title_cell.font = Font(bold=True, color="FFFFFF", size=12, name="Arial")
        title_cell.alignment = CA

        # 列头
        hdr(2, 1, "排名");  hdr(2, 2, "淘宝热门关键词", RED);  hdr(2, 3, "京东热门关键词", GRN)

        tb_words = hotwords.get("淘宝", [])
        jd_words = hotwords.get("京东", [])
        max_len = max(len(tb_words), len(jd_words), 1)

        for i in range(max_len):
            row = i + 3
            gray = (i % 2 == 0)
            cell(row, 1, i + 1, gray)
            cell(row, 2, tb_words[i] if i < len(tb_words) else "", gray)
            cell(row, 3, jd_words[i] if i < len(jd_words) else "", gray)

        ws.column_dimensions["A"].width = 8
        ws.column_dimensions["B"].width = 35
        ws.column_dimensions["C"].width = 35
        ws.freeze_panes = "A3"
        apply_col_comments(ws, "热门关键词")

        wb.save(filepath)
        log.info(f"  [热词] 已写入Excel：淘宝{len(tb_words)}个 / 京东{len(jd_words)}个")
    except Exception as e:
        log.warning(f"  [热词] 写入Excel失败: {e}")



def scrape_jd(driver, keyword, pages, log, stop_event=None, use_current=False):
    """采集京东搜索结果"""
    keep_minimized(driver)
    log.info(f"\n{'='*50}")
    log.info(f"  [京东] 开始采集: {keyword}  Pages: {pages}")
    log.info(f"{'='*50}")
    items = []
    try:
        if not use_current:
            url = f"https://search.jd.com/Search?keyword={quote(keyword)}&enc=utf-8&page=1"
            driver.get(url)
            keep_minimized(driver)
            rsleep(5, 8, stop_event)
        # use_current=True: 已切换到京东标签页，无论当前什么页面，直接跳转到搜索页
        log.info(f"  [京东] 当前页面: {safe_title(driver)[:60]}")
        log.info(f"  [京东] URL: {safe_url(driver)[:80]}")

        # 检测登录页
        if "欢迎登录" in safe_title(driver) or "passport.jd.com" in safe_url(driver):
            log.warning("  [京东] 检测到登录页，跳过采集")
            return items

        # 检测是否已在京东搜索结果页（保留用户手动设置的筛选条件）
        cur = safe_url(driver).lower()
        is_error = ("error" in cur or "global.jd" in cur or "passport" in cur)
        is_on_search = "search.jd.com" in cur and not is_error
        if is_on_search:
            # 校验当前页面的关键词是否与目标一致，防止采集到上次搜索的残留数据
            from urllib.parse import urlparse, parse_qs, unquote
            try:
                qs = parse_qs(urlparse(safe_url(driver)).query)
                cur_kw = unquote(qs.get("keyword", [""])[0]).strip()
            except Exception:
                cur_kw = ""
            if cur_kw and cur_kw == keyword.strip():
                log.info(f"  [京东] ✅ 已在搜索结果页且关键词匹配，直接采集（保留筛选条件）")
                time.sleep(2)
            else:
                log.info(f"  [京东] ⚠️ 当前关键词「{cur_kw}」≠ 目标「{keyword}」，重新搜索")
                search_url = f"https://search.jd.com/Search?keyword={quote(keyword)}&enc=utf-8&page=1"
                driver.get(search_url)
                keep_minimized(driver)
                # 等待价格元素加载（最多15秒），检查容器文本含数字
                for _w in range(15):
                    time.sleep(1)
                    _price_els = driver.find_elements(By.CSS_SELECTOR, '[class*="price"], [class*="Price"]')
                    if any(any(c.isdigit() for c in (el.text or "")) for el in _price_els):
                        break
                log.info(f"  [京东] 跳转后: {safe_title(driver)[:50]}")
        else:
            search_url = f"https://search.jd.com/Search?keyword={quote(keyword)}&enc=utf-8&page=1"
            log.info(f"  [京东] 跳转到搜索页: {search_url[:60]}")
            driver.get(search_url)
            keep_minimized(driver)
            # 等待价格元素加载（最多15秒），检查容器文本含数字
            for _w in range(15):
                time.sleep(1)
                _price_els = driver.find_elements(By.CSS_SELECTOR, '[class*="price"], [class*="Price"]')
                if any(any(c.isdigit() for c in (el.text or "")) for el in _price_els):
                    break
            log.info(f"  [京东] 跳转后: {safe_title(driver)[:50]}")

        for page_i in range(pages):
            if stop_event and stop_event.is_set():
                break
            log.info(f"  [京东] 第 {page_i+1}/{pages} 页")

            # 充分滚动确保懒加载
            for _ in range(8):
                driver.execute_script("window.scrollBy(0, window.innerHeight * 0.7)")
                time.sleep(0.7)
            driver.execute_script("window.scrollTo(0, 0)")
            time.sleep(1.5)

            # 优先用 data-sku，最可靠
            cards = driver.find_elements(By.CSS_SELECTOR, "[data-sku]")
            if not cards:
                cards = driver.find_elements(By.CSS_SELECTOR, "li.gl-item")
            if not cards:
                cards = driver.find_elements(By.CSS_SELECTOR, "div[class*='gl-item']")
            if not cards:
                # 最后手段：找所有包含价格的容器
                cards = driver.find_elements(By.CSS_SELECTOR, "div[class*='goods-list'] > div, ul.gl-warp > li")
            log.info(f"  [京东] 找到 {len(cards)} 个商品卡片")

            if not cards:
                # 检测是否是京东风控验证页
                page_src = driver.page_source
                is_captcha = ("验证一下" in page_src or "快速验证" in page_src or
                              "risk_handler" in safe_url(driver) or
                              "京东验证" in safe_title(driver))
                if is_captcha:
                    log.warning("  [京东] ⚠️ 触发风控验证！请在浏览器中手动点击「快速验证」通过，工具将等待最多60秒后继续...")
                    for _w in range(60):
                        if stop_event and stop_event.is_set():
                            break
                        time.sleep(1)
                        # 检测是否已通过验证（回到搜索结果页）
                        if "search.jd.com" in safe_url(driver) and "risk_handler" not in safe_url(driver):
                            log.info("  [京东] ✅ 验证通过，继续采集...")
                            time.sleep(2)
                            break
                    else:
                        log.warning("  [京东] 等待超时，跳过本页")
                        break
                    # 验证通过后重新获取卡片
                    cards = driver.find_elements(By.CSS_SELECTOR,
                        "div.gl-item, li.gl-item, [class*='gl-item'], [class*='goods-item'], "
                        "[data-sku], [class*='ProductCard'], div[class*='item'][data-sku]")
                    if not cards:
                        break
                else:
                    log.warning(f"  [京东] 未找到卡片，保存页面HTML...")
                    try:
                        import os
                        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "jd_debug.html"), "w", encoding="utf-8") as _f:
                            _f.write(page_src)
                        log.info("  [京东] HTML已保存到 jd_debug.html")
                    except Exception: pass
                    break

            now = datetime.now().strftime("%Y-%m-%d %H:%M")
            page_items = []

            import re as _re
            parsed_count = 0
            for card in cards:
                if stop_event and stop_event.is_set():
                    break
                try:
                    raw = card.get_attribute("outerHTML") or ""
                    txt = card.get_attribute("textContent") or card.text or ""

                    # ── 标题：优先CSS，再用title属性，最后card.text ──
                    title = ""
                    # 方式1：传统class（gl-item模式）
                    for sel in ["div.p-name em", ".p-name em", "div.p-name a", ".p-name a",
                                "[class*='p-name'] em", "[class*='p-name'] a",
                                "[class*='title'] a", "[class*='name'] a",
                                "[class*='goods-name']", "[class*='item-name']"]:
                        try:
                            t = (card.find_element(By.CSS_SELECTOR, sel).text or "").strip()
                            if len(t) > 4: title = t; break
                        except Exception: pass
                    # 方式2：[title]属性（内容模式 - 取最长的title，即商品名）
                    if not title:
                        try:
                            titled_els = card.find_elements(By.CSS_SELECTOR, "[title]")
                            best = ""
                            for el in titled_els:
                                t = (el.get_attribute("title") or "").strip()
                                if len(t) > len(best):
                                    best = t
                            if len(best) > 4:
                                title = best
                        except Exception: pass
                    # 方式3：card.text第一行（通常是标题）
                    if not title:
                        lines = [l.strip() for l in txt.split("\n") if len(l.strip()) > 6]
                        if lines: title = lines[0][:80]

                    # ── 价格：优先CSS，再内容模式（¥+数字），最后HTML正则 ──
                    price_str = ""
                    for sel in ["div.p-price strong i", ".p-price i", "div.p-price strong",
                                ".p-price strong", "[class*='p-price'] i", "[class*='p-price'] em",
                                "[class*='goods-price']", "[class*='item-price']"]:
                        try:
                            t = card.find_element(By.CSS_SELECTOR, sel).text.strip()
                            if t and any(c.isdigit() for c in t):
                                price_str = t; break
                        except Exception: pass
                    # 新版CSS Module: [class*='price'] 匹配容器取完整文本
                    if not price_str:
                        for sel in ["[class*='price']", "[class*='Price']"]:
                            try:
                                for el in card.find_elements(By.CSS_SELECTOR, sel):
                                    t = (el.text or "").strip()
                                    pm = _re.search(r'(\d+\.?\d*)', t)
                                    if pm and float(pm.group(1)) > 0:
                                        price_str = pm.group(1); break
                                if price_str: break
                            except Exception: pass
                    # 内容模式：找包含¥的元素，取父元素完整文本（含¥+数字）
                    if not price_str:
                        try:
                            for el in card.find_elements(By.XPATH,
                                    ".//*[contains(text(),'¥') or contains(text(),'￥')]"):
                                try:
                                    parent = el.find_element(By.XPATH, "..")
                                    parent_text = (parent.text or "").strip()
                                except Exception:
                                    parent_text = (el.text or "").strip()
                                pm = _re.search(r'[¥￥]\s*(\d+\.?\d*)', parent_text)
                                if pm:
                                    price_str = pm.group(1); break
                        except Exception: pass
                    # card.text 正则兜底
                    if not price_str:
                        pm = _re.search(r'[¥￥]\s*(\d+\.?\d*)', txt)
                        if not pm:
                            pm = _re.search(r'到手价\D*(\d+\.?\d*)', txt)
                        if pm: price_str = pm.group(1)
                    if not price_str:
                        # HTML正则兜底：匹配 ¥</i>数字 或 price容器内数字
                        pm = _re.search(r'[¥￥]</[^>]+>\s*(\d+[\d,.]*)', raw)
                        if not pm:
                            pm = _re.search(r'(?:price|Price)[^"]*"[^>]*>\s*<[^>]+>\s*([\d,\.]+)', raw)
                        if pm: price_str = pm.group(1)

                    price = clean_price(price_str)

                    if not title or not price:
                        if not title: log.info(f"  [京东] 跳过：无标题  text前30={txt[:30]!r}")
                        if not price: log.info(f"  [京东] 跳过：无价格  raw_price={price_str!r}")
                        continue

                    # ── 店铺 ──
                    shop = ""
                    for sel in ["div.p-shop a", ".p-shop a", "[class*='shop'] a",
                                "[class*='shop'] span", "[class*='shop']",
                                "[class*='store']", "[class*='seller']"]:
                        try:
                            el = card.find_element(By.CSS_SELECTOR, sel)
                            t = (el.get_attribute("title") or el.get_attribute("textContent") or "").strip()
                            if t and len(t) < 40:
                                shop = t; break
                        except Exception: pass
                    # 内容模式：从textContent找店铺名（含"店"字的短文本）
                    if not shop:
                        for line in txt.split("\n"):
                            line = line.strip()
                            if 4 < len(line) < 30 and "店" in line and "搜" not in line:
                                shop = line; break

                    # ── 销量/评价数 ──
                    sales_str = ""
                    # 方式1：新版京东 - class含volume/sold的元素取title属性（不受窗口可见性影响）
                    for sel in ["[class*='volume'] span[title]", "[class*='volume'][title]",
                                "[class*='sold'] span[title]", "[class*='sold'][title]"]:
                        try:
                            t = (card.find_element(By.CSS_SELECTOR, sel).get_attribute("title") or "").strip()
                            if t and any(c.isdigit() for c in t):
                                sales_str = t; break
                        except Exception: pass
                    # 方式2：CSS选择器 + textContent（不受窗口可见性影响）
                    if not sales_str:
                        for sel in ["[class*='volume']", "[class*='sold']",
                                    "[class*='comment']", "[class*='commit']",
                                    "div.p-commit strong a", "div.p-commit strong",
                                    ".p-commit strong"]:
                            try:
                                t = (card.find_element(By.CSS_SELECTOR, sel).get_attribute("textContent") or "").strip()
                                if t and any(c.isdigit() for c in t):
                                    sales_str = t; break
                            except Exception: pass
                    # 方式3：textContent正则兜底
                    if not sales_str:
                        sm = _re.search(r'已售\s*([\d.]+[+万]*)', txt)
                        if not sm:
                            sm = _re.search(r'([\d.]+[+万]*)\s*(?:种草|评价|条|人好评)', txt)
                        if sm: sales_str = sm.group(1)
                    sales = parse_num(sales_str)

                    # ── 京东评价数（默认与销量同源） ──
                    comment_count = sales

                    # ── 链接：优先用 data-sku 拼接，回退 regex / <a> ──
                    link = ""
                    sku = (card.get_attribute("data-sku") or "").strip()
                    if sku and sku.isdigit():
                        link = f"https://item.jd.com/{sku}.html"
                    else:
                        lm = _re.search(r'href="((?:https?:)?//item\.jd\.com/[^"]+)"', raw)
                        if lm:
                            href = lm.group(1)
                            link = "https:" + href if href.startswith("//") else href
                        # 清理链接：截断到 .html 结尾，去掉后面的追踪参数
                        if link:
                            m_clean = _re.match(r'(https?://item\.jd\.com/\d+\.html)', link)
                            if m_clean: link = m_clean.group(1)

                    # ── 图片 ──
                    img_url = ""
                    try:
                        img = card.find_element(By.CSS_SELECTOR, "img")
                        loading_url = (img.get_attribute("data-loading") or "").strip()
                        for attr in ["data-lazy-img", "data-src", "src"]:
                            v = (img.get_attribute(attr) or "").strip()
                            if v and not v.startswith("data:") and "placeholder" not in v:
                                if loading_url and v == loading_url:
                                    continue  # 跳过京东 loading 占位图
                                if not v.startswith("http"): v = "https:" + v
                                img_url = v; break
                    except Exception: pass

                    page_items.append({
                        "商品主图":     img_url,
                        "采集时间":     now,
                        "平台":         "京东",
                        "商品标题":     title[:80],
                        "产品规格":     extract_spec(title[:80]),
                        "套装类型":     extract_bundle_type(title[:80]),
                        "当前售价(元)": price,
                        "月销量(件)":   sales,        # 京东以评价数代替销量展示
                        "评价数量":     comment_count, # 单独提取的评价数量
                        "店铺名称":     shop,
                        "店铺类型":     "JD",
                        "店铺地区":     "",
                        "商品链接":     link,
                    })
                    parsed_count += 1
                except Exception as e:
                    log.debug(f"  [京东] 卡片解析异常: {e}")
                    continue

            log.info(f"  [京东] 本页解析成功 {parsed_count}/{len(cards)} 条")

            log.info(f"  [京东] 本页解析 {len(page_items)} 条")
            # 如果还是0条，保存前3个卡片HTML供诊断
            if len(page_items) == 0 and cards:
                try:
                    import os
                    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "jd_card_debug.html"), "w", encoding="utf-8") as _f:
                        _f.write("\n\n<!-- NEXT CARD -->\n\n".join(c.get_attribute("outerHTML") for c in cards[:5]))
                    log.info("  [京东] 卡片HTML已保存到 jd_card_debug.html")
                except Exception: pass
            items.extend(page_items)

            # 翻页: 京东搜索页用滚动加载+点击下一页按钮
            if page_i < pages - 1:
                try:
                    # First try clicking "next page" button
                    clicked_next = False
                    for sel in ["a.pn-next", ".pn-next", "[class*='pn-next']",
                                "a[aria-label='下一页']", "a.fp-next",
                                "[class*='page'] [class*='next']"]:
                        try:
                            btns = driver.find_elements(By.CSS_SELECTOR, sel)
                            for btn in btns:
                                if btn.is_displayed() and btn.is_enabled():
                                    btn.click()
                                    clicked_next = True
                                    log.info(f"  [京东] 点击下一页按钮: {sel}")
                                    break
                            if clicked_next:
                                break
                        except Exception:
                            continue

                    if not clicked_next:
                        # Fallback: URL parameter approach
                        cur = safe_url(driver)
                        if "jd.com" not in cur.lower():
                            log.warning("  [京东] 翻页时不在京东页面，停止")
                            break
                        import re as _re
                        next_page = (page_i + 1) * 2 + 1
                        new_url = _re.sub(r"page=\d+", f"page={next_page}", cur) if "page=" in cur else cur + f"&page={next_page}"
                        driver.get(new_url)
                        log.info(f"  [京东] URL翻页: page={next_page}")

                    keep_minimized(driver)
                    rsleep(4, 6, stop_event)
                except Exception as e:
                    log.warning(f"  [京东] 翻页失败: {e}"); break

    except Exception as e:
        log.error(f"  [京东] 采集出错: {e}")
        traceback.print_exc()

    log.info(f"  [京东] 共采集 {len(items)} 条")
    return items


# ── 拼多多采集（暂停使用，保留代码以备将来恢复） ──
def _scrape_pdd_disabled(driver, keyword, pages, log, stop_event=None, use_current=False):
    """采集拼多多搜索结果（滚动加载）—— 暂停使用"""
    log.info(f"\n{'='*50}")
    log.info(f"  [拼多多] 开始采集: {keyword}  Pages: {pages}")
    log.info(f"{'='*50}")
    items = []
    seen_titles = set()
    try:
        if not use_current:
            url = f"https://mobile.yangkeduo.com/search_result.html?search_key={quote(keyword)}&source=search_history"
            driver.get(url)
            rsleep(5, 8, stop_event)
        else:
            # use_current=True: 在拼多多标签内，统一跳手机版搜索页
            log.info(f"  [拼多多] 当前页面: {safe_title(driver)[:50]}")
            search_url = f"https://mobile.yangkeduo.com/search_result.html?search_key={quote(keyword)}&source=search_history"
            log.info(f"  [拼多多] 跳转手机版搜索页...")
            driver.get(search_url)
            rsleep(5, 8, stop_event)
            log.info(f"  [拼多多] 搜索后: {safe_title(driver)[:50]}")
        time.sleep(2)

        # ── 先切换到手机模式，用手机UA访问搜索页 ──
        try:
            mobile_ua = ("Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) "
                         "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1")
            driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": mobile_ua})
            driver.execute_cdp_cmd("Emulation.setDeviceMetricsOverride",
                                   {"width": 390, "height": 844, "deviceScaleFactor": 3, "mobile": True})
            log.info("  [拼多多] 已切换手机模式")
        except Exception as _e:
            log.debug(f"  [拼多多] 切换手机模式: {_e}")

        # 跳转到手机版搜索页
        search_url = f"https://mobile.yangkeduo.com/search_result.html?search_key={quote(keyword)}"
        log.info(f"  [拼多多] 搜索URL: {search_url[:80]}")
        driver.get(search_url)
        time.sleep(5)
        log.info(f"  [拼多多] 跳转后: {safe_title(driver)[:50]} | {safe_url(driver)[:60]}")

        # 滚动加载所有商品
        scroll_rounds = pages * 6
        for i in range(scroll_rounds):
            if stop_event and stop_event.is_set(): break
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(1.5)
            if (i+1) % 6 == 0:
                log.info(f"  [拼多多] 滚动 {i+1}/{scroll_rounds}...")
        time.sleep(3)

        # 尝试多种卡片选择器
        cards = []
        for sel in [
                    # 手机版 yangkeduo 搜索结果标准选择器
                    "li.search-list-common-card",
                    "li[class*='search-list']",
                    "div[class*='goods-item']",
                    "div[class*='goodsItem']",
                    "div[class*='SearchItem']",
                    "div[class*='search-item']",
                    # 通用但排除纯导航li（导航li没有价格子元素）
                    "div[class*='card']",
                    "a[href*='goods.pinduoduo']",
                    "a[href*='yangkeduo']",
                ]:
            try:
                found = driver.find_elements(By.CSS_SELECTOR, sel)
                if len(found) > len(cards):
                    log.info(f"  [拼多多] {sel!r} 找到 {len(found)} 个")
                    cards = found
            except Exception: pass

        # 过滤：去掉没有价格子元素的卡片（排除导航菜单等误匹配）
        valid_cards = []
        for c in cards:
            try:
                html_c = c.get_attribute("innerHTML") or ""
                # 包含数字且卡片文本长度合理
                import re as _re
                if _re.search(r'\d{2,5}', html_c) and len(c.text.strip()) > 5:
                    valid_cards.append(c)
            except Exception: pass
        if valid_cards:
            cards = valid_cards
        log.info(f"  [拼多多] 有效卡片: {len(cards)} 个")

        if not cards:
            try:
                import os
                with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "pdd_debug.html"), "w", encoding="utf-8") as _f:
                    _f.write(driver.page_source)
                log.info("  [拼多多] 页面已保存到 pdd_debug.html")
            except Exception: pass

        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        for card in cards:
            if stop_event and stop_event.is_set(): break
            try:
                # ── 标题 ──
                title = ""
                for sel in ["div[class*='title']", "span[class*='title']",
                            "div[class*='name']", "span[class*='name']",
                            "h3", "p"]:
                    try:
                        t = card.find_element(By.CSS_SELECTOR, sel).text.strip()
                        if len(t) > 4: title = t; break
                    except Exception: pass
                if not title or title in seen_titles: continue

                # ── 价格 ──
                price_str = ""
                for sel in ["div[class*='price']", "span[class*='price']",
                            "div[class*='Price']", "span[class*='Price']", "strong"]:
                    try:
                        t = card.find_element(By.CSS_SELECTOR, sel).text.strip()
                        if t and any(c.isdigit() for c in t):
                            price_str = t; break
                    except Exception: pass
                if not price_str:
                    import re as _re
                    m = _re.search(r'(\d+\.\d{1,2}|\d{2,5})', card.get_attribute("outerHTML") or "")
                    if m: price_str = m.group(1)
                price = clean_price(price_str)
                if not price: continue

                # ── 销量 ──
                sales_str = ""
                for sel in ["[class*='sold']","[class*='sale']","[class*='buy']","[class*='count']"]:
                    try:
                        t = card.find_element(By.CSS_SELECTOR, sel).text.strip()
                        if t: sales_str = t; break
                    except Exception: pass
                sales = parse_num(sales_str)

                # ── 链接 ──
                link = ""
                try:
                    tag = card if card.tag_name == "a" else card.find_element(By.CSS_SELECTOR, "a")
                    href = tag.get_attribute("href") or ""
                    link = "https:" + href if href.startswith("//") else href
                except Exception: pass

                # ── 图片 ──
                img_url = ""
                try:
                    img = card.find_element(By.CSS_SELECTOR, "img")
                    for attr in ["src", "data-src", "data-lazy"]:
                        v = (img.get_attribute(attr) or "").strip()
                        if v and not v.startswith("data:"):
                            img_url = v; break
                except Exception: pass

                seen_titles.add(title)
                items.append({
                    "采集时间":     now,
                    "平台":         "拼多多",
                    "商品主图":     img_url,
                    "商品标题":     title[:80],
                    "当前售价(元)": price,
                    "月销量(件)":   sales,
                    "店铺名称":     "",
                    "商品链接":     link,
                })
            except Exception as e:
                log.debug(f"  [拼多多] 解析: {e}")
                continue

        # 恢复桌面模式
        try:
            driver.execute_cdp_cmd("Emulation.clearDeviceMetricsOverride", {})
            driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": ""})
        except Exception: pass

        if len(items) == 0 and cards:
            try:
                import os
                with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "pdd_card_debug.html"), "w", encoding="utf-8") as _f:
                    _f.write("\n\n<!-- NEXT -->\n\n".join(c.get_attribute("outerHTML") for c in cards[:5]))
                log.info("  [拼多多] 卡片HTML已保存到 pdd_card_debug.html")
            except Exception: pass

    except Exception as e:
        log.error(f"  [拼多多] 采集出错: {e}")
        traceback.print_exc()

    log.info(f"  [拼多多] 共采集 {len(items)} 条")
    return items


def analyse_reviews_from_items(tb_items, log):
    """基于已采集的评价数量数据做分析，不需要跳转页面"""
    if not tb_items:
        return {"total": 0, "pos": [], "neg": [], "top_words": [], "review_stats": []}

    import pandas as pd
    df = pd.DataFrame(tb_items)
    df["_reviews"] = pd.to_numeric(df.get("评价数量", pd.Series()), errors="coerce").fillna(0)
    df["_sales"]   = pd.to_numeric(df.get("月销量(件)", pd.Series()), errors="coerce").fillna(0)
    df["_price"]   = pd.to_numeric(df.get("当前售价(元)", pd.Series()), errors="coerce").fillna(0)

    # Top10评价最多的商品
    top_review = df.nlargest(10, "_reviews")[
        ["商品标题", "店铺名称", "当前售价(元)", "月销量(件)", "评价数量"]
    ].to_dict("records")

    # 评价/销量比（口碑转化率）：评价数越接近销量说明买家越愿意评价
    df["_ratio"] = df.apply(
        lambda r: round(r["_reviews"] / r["_sales"] * 100, 1) if r["_sales"] > 0 else 0, axis=1)
    top_ratio = df[df["_sales"] > 10].nlargest(5, "_ratio")[
        ["商品标题", "店铺名称", "评价数量", "月销量(件)"]
    ].to_dict("records")

    total_reviews = int(df["_reviews"].sum())
    avg_reviews   = round(float(df["_reviews"].mean()), 1)
    max_reviews   = int(df["_reviews"].max())

    log.info(f"  [评价分析] 共{len(tb_items)}条商品，总评价数{total_reviews}，"
             f"均值{avg_reviews}，最高{max_reviews}")

    return {
        "total": total_reviews,
        "avg":   avg_reviews,
        "max":   max_reviews,
        "top_review": top_review,   # 评价最多的Top10
        "top_ratio":  top_ratio,    # 口碑转化率Top5
        "pos": [], "neg": [], "top_words": [],
        "review_stats": top_review,
    }



def scrape_reviews(driver, item_url, platform, max_reviews, log, stop_event=None):
    """采集单个商品评价（淘宝/京东），返回评价文本列表"""
    reviews = []
    try:
        if platform == "taobao":
            driver.get(item_url)
            rsleep(3, 5, stop_event)
            try:
                btn = driver.find_element(By.CSS_SELECTOR,
                    "[class*='comment'],[class*='review'],[class*='评价']")
                btn.click()
                rsleep(2, 3, stop_event)
            except Exception:
                pass
            for sel in ["[class*='reviewContent']", "[class*='content--']",
                        "[class*='comment-content']", "p[class*='comment']"]:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
                for el in els[:max_reviews]:
                    t = el.text.strip()
                    if t and len(t) > 5:
                        reviews.append(t)
                if reviews:
                    break

        elif platform == "jd":
            # 京东评价直接用API
            try:
                import re as _re
                sku_match = _re.search(r'/(\d+)\.html', item_url)
                if sku_match:
                    sku = sku_match.group(1)
                    api_url = (f"https://club.jd.com/comment/productPageComments.action"
                               f"?productId={sku}&score=0&sortType=5&page=0&pageSize=10")
                    driver.get(api_url)
                    rsleep(2, 3, stop_event)
                    import json
                    body = driver.find_element(By.TAG_NAME, "body").text
                    data = json.loads(body)
                    for c in data.get("comments", [])[:max_reviews]:
                        t = c.get("content", "").strip()
                        if t:
                            reviews.append(t)
            except Exception as e:
                log.debug(f"  [评价] 京东API: {e}")

    except Exception as e:
        log.debug(f"  [评价] 采集出错: {e}")
    return reviews


def analyse_reviews(reviews, log):
    """统计评价关键词，返回正面/负面/高频词"""
    if not reviews:
        return {"total": 0, "pos": [], "neg": [], "top_words": []}

    pos_kws = ["好用", "效果好", "物超所值", "推荐", "满意", "不错", "很好",
               "喜欢", "正品", "值得", "补水", "保湿", "滋润", "白", "嫩"]
    neg_kws = ["差", "失望", "假货", "不好用", "无效", "刺激", "过敏", "贵",
               "发货慢", "态度差", "退货", "后悔", "坑", "劣质"]

    pos_count = {k: 0 for k in pos_kws}
    neg_count = {k: 0 for k in neg_kws}
    word_freq  = {}

    for review in reviews:
        for k in pos_kws:
            if k in review:
                pos_count[k] += 1
        for k in neg_kws:
            if k in review:
                neg_count[k] += 1
        # 简单分词：按2-4字切分
        for n in [2, 3, 4]:
            for i in range(len(review) - n + 1):
                w = review[i:i+n]
                if re.search(r'[\u4e00-\u9fff]{2,}', w):
                    word_freq[w] = word_freq.get(w, 0) + 1

    pos_result  = sorted([(k, v) for k, v in pos_count.items() if v > 0],
                          key=lambda x: -x[1])[:8]
    neg_result  = sorted([(k, v) for k, v in neg_count.items() if v > 0],
                          key=lambda x: -x[1])[:8]
    # 过滤停用词
    stopwords = {"的", "了", "是", "在", "我", "有", "和", "就", "不", "人",
                 "都", "一", "一个", "上", "也", "很", "到", "说", "要", "去",
                 "你", "会", "着", "没有", "看", "好的", "这个", "感觉", "收到",
                 "东西", "快递", "包装", "质量", "产品", "使用", "购买", "买"}
    top_words = sorted(
        [(w, c) for w, c in word_freq.items()
         if c >= 2 and w not in stopwords and len(w) >= 2],
        key=lambda x: -x[1]
    )[:15]

    log.info(f"  [评价] 共分析 {len(reviews)} 条，正面词 {len(pos_result)} 个，负面词 {len(neg_result)} 个")
    return {
        "total":     len(reviews),
        "pos":       pos_result,
        "neg":       neg_result,
        "top_words": top_words,
    }


def _save_platform_excel(items, filepath, platform_name, header_color, keyword, review_data, log):
    """保存单平台数据到独立Excel文件"""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    import os

    wb  = Workbook()
    ws  = wb.active
    ws.title = f"{platform_name}数据"

    HDR  = PatternFill("solid", fgColor=header_color)
    GRAY = PatternFill("solid", fgColor="F5F6FA")
    WH   = PatternFill("solid", fgColor="FFFFFF")
    WF   = Font(bold=True, color="FFFFFF", name="微软雅黑", size=10)
    NF   = Font(name="微软雅黑", size=9)
    CA   = Alignment(horizontal="center", vertical="center")
    LA   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    cols = ["商品主图", "采集时间", "商品标题", "产品规格", "套装类型", "当前售价(元)", "月销量(件)", "评价数量", "店铺名称", "店铺类型", "店铺地区", "商品链接"]
    for ci, col in enumerate(cols, 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.fill = HDR; c.font = WF; c.alignment = CA

    for ri, item in enumerate(items, 2):
        fill = GRAY if ri % 2 == 0 else WH
        for ci, col in enumerate(cols, 1):
            c = ws.cell(row=ri, column=ci, value=item.get(col, ""))
            c.fill = fill; c.font = NF; c.alignment = LA

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 50
    ws.freeze_panes = "B2"

    # 评价分析单独sheet（仅淘宝文件含此sheet）
    if review_data and review_data.get("total", 0) > 0:
        RED2  = PatternFill("solid", fgColor="2A5298")
        GRN   = PatternFill("solid", fgColor="27AE60")
        ORG   = PatternFill("solid", fgColor="E67E22")
        PURP  = PatternFill("solid", fgColor="8E44AD")
        ws_r  = wb.create_sheet("评价分析")

        ws_r.cell(row=1, column=1, value=f"评价分析 - {keyword}").fill = PURP
        ws_r.cell(row=1, column=1).font = WF
        ws_r.cell(row=1, column=1).alignment = CA
        ws_r.merge_cells("A1:B1")
        ws_r.cell(row=2, column=1, value=f"共分析评价：{review_data['total']} 条").font = NF

        r = 4
        for label, fill_c, data_key in [
            ("✅ 正面高频词", GRN, "pos"),
            ("❌ 负面高频词", ORG, "neg"),
            ("📊 全部高频词", RED2, "top_words"),
        ]:
            ws_r.cell(row=r, column=1, value=label).fill = fill_c
            ws_r.cell(row=r, column=1).font = WF
            ws_r.cell(row=r, column=2, value="出现次数").fill = fill_c
            ws_r.cell(row=r, column=2).font = WF
            r += 1
            for word, cnt in review_data.get(data_key, []):
                ws_r.cell(row=r, column=1, value=word).font = NF
                ws_r.cell(row=r, column=2, value=cnt).font = NF
                r += 1
            r += 1

        ws_r.column_dimensions["A"].width = 20
        ws_r.column_dimensions["B"].width = 12

    wb.save(filepath)
    log.info(f"  [{platform_name}] 已保存: {filepath}  ({len(items)} 条)")


def save_multiplatform_excel(tb_items, jd_items, review_data,
                              filepath, keyword, log):
    """把京东、评价分析写入淘宝同一个Excel文件的新sheet"""
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    import os

    if not os.path.exists(filepath):
        log.warning(f"  [多平台] 找不到淘宝数据文件: {filepath}")
        return

    try:
        wb = load_workbook(filepath)
    except Exception as e:
        log.warning(f"  [多平台] 打开文件失败: {e}")
        return

    from openpyxl.styles import Border, Side
    GRAY = PatternFill("solid", fgColor="F5F6FA")
    WH   = PatternFill("solid", fgColor="FFFFFF")
    WF   = Font(bold=True, color="FFFFFF", name="微软雅黑", size=10)
    NF   = Font(name="微软雅黑", size=9)
    CA   = Alignment(horizontal="center", vertical="center")
    LA   = Alignment(horizontal="left", vertical="center", wrap_text=True)
    BD   = Border(
        left=Side(style="thin", color="DDDDDD"), right=Side(style="thin", color="DDDDDD"),
        top=Side(style="thin", color="DDDDDD"),  bottom=Side(style="thin", color="DDDDDD"))

    cols = ["商品主图", "采集时间", "商品标题", "产品规格", "套装类型", "当前售价(元)", "月销量(件)", "评价数量", "店铺名称", "店铺类型", "店铺地区", "商品链接"]

    _SHEET_TT = {
        "采集时间":     "本条数据的采集时间",
        "平台":         "数据来源平台（京东/淘宝）",
        "商品主图":     "商品主图URL",
        "商品标题":     "京东商品完整标题",
        "产品规格":     "从标题提取的规格，如容量/件数/装型",
        "当前售价(元)": "京东商品当前售价（元）",
        "月销量(件)":   "近30天评价数，作为销量参考",
        "店铺名称":     "京东卖家店铺名称",
        "商品链接":     "京东商品详情页URL",
        "商品数":       "该平台采集到的商品总件数",
        "均价(元)":     "该平台所有商品的平均售价",
        "最低价":       "该平台采集到的最低售价",
        "最高价":       "该平台采集到的最高售价",
        "Top3均销量":   "该平台月销量前3名商品的平均销量",
    }
    def write_sheet(wb, sheet_name, items, hdr_color):
        from openpyxl.comments import Comment as _Cmt2
        from openpyxl.utils import get_column_letter as _gcl
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(sheet_name)
        HDR = PatternFill("solid", fgColor=hdr_color)
        IMG_COL = 1   # 商品主图在第1列
        IMG_SZ  = 80

        # 表头
        for ci, col in enumerate(cols, 1):
            c = ws.cell(row=1, column=ci, value=col)
            c.fill = HDR; c.font = WF; c.alignment = CA; c.border = BD
            if col in _SHEET_TT:
                cm = _Cmt2(_SHEET_TT[col], "📋 列说明")
                cm.width = 260; cm.height = 55
                c.comment = cm

        # 数据行（商品主图列写URL，后面再替换成图片）
        for ri, item in enumerate(items, 2):
            fill = GRAY if ri % 2 == 0 else WH
            ws.row_dimensions[ri].height = 65
            for ci, col in enumerate(cols, 1):
                c = ws.cell(row=ri, column=ci, value=item.get(col, ""))
                c.fill = fill; c.font = NF
                c.alignment = CA if ci == IMG_COL else LA
                c.border = BD

        # 嵌入图片（替换商品主图列）
        img_col_name = "商品主图"
        if img_col_name in cols:
            img_ci = cols.index(img_col_name) + 1
            ok = 0
            for ri, item in enumerate(items, 2):
                url = str(item.get(img_col_name, "") or "").strip()
                if not url or url == "nan": continue
                try:
                    buf = download_image(url, IMG_SZ)
                    if buf:
                        xl = XLImage(buf)
                        xl.width = IMG_SZ; xl.height = IMG_SZ
                        ws.add_image(xl, _gcl(img_ci) + str(ri))
                        ws.cell(ri, img_ci).value = ""  # 清掉URL文字
                        ok += 1
                except Exception as _ie:
                    log.debug(f"  [{sheet_name}] 图片嵌入失败 row{ri}: {_ie}")
            log.info(f"  [{sheet_name}] 图片嵌入 {ok}/{len(items)} 张")

        # 列宽
        col_widths = [12, 16, 45, 14, 12, 12, 12, 12, 22, 10, 12, 50]
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[_gcl(i)].width = w
        ws.freeze_panes = "C2"
        apply_col_comments(ws, sheet_name)
        log.info(f"  [{sheet_name}] 写入 {len(items)} 条")

    if jd_items:
        write_sheet(wb, "京东数据", jd_items, "27AE60")

    # ── 跨平台对比摘要sheet ──
    if jd_items:
        cmp_name = "跨平台对比"
        if cmp_name in wb.sheetnames:
            del wb[cmp_name]
        ws_c = wb.create_sheet(cmp_name)
        HDR_C = PatternFill("solid", fgColor="2A5298")

        cmp_headers = ["平台", "商品数", "均价(元)", "最低价", "最高价", "Top3均销量"]
        for ci, h in enumerate(cmp_headers, 1):
            from openpyxl.comments import Comment as _CmtC
            c = ws_c.cell(row=1, column=ci, value=h)
            c.fill = HDR_C; c.font = WF; c.alignment = CA
            if h in _SHEET_TT:
                cm = _CmtC(_SHEET_TT[h], "📋 列说明")
                cm.width = 260; cm.height = 55
                c.comment = cm

        platform_fills = [
            PatternFill("solid", fgColor="D5E8F0"),
            PatternFill("solid", fgColor="D5F0D5"),
            PatternFill("solid", fgColor="F0E5D5"),
        ]
        price_col = "当前售价(元)"
        sales_col = "月销量(件)"

        row = 2
        for idx, (name, items) in enumerate([("淘宝", tb_items), ("京东", jd_items)]):
            if not items:
                continue
            prices = [float(i.get(price_col, 0) or 0) for i in items if i.get(price_col)]
            sales  = sorted([int(i.get(sales_col, 0) or 0) for i in items if i.get(sales_col)], reverse=True)
            avg_p  = round(sum(prices)/len(prices), 1) if prices else 0
            min_p  = round(min(prices), 1) if prices else 0
            max_p  = round(max(prices), 1) if prices else 0
            top3   = round(sum(sales[:3])/min(3,len(sales)), 0) if sales else 0
            for ci, val in enumerate([name, len(items), avg_p, min_p, max_p, int(top3)], 1):
                c = ws_c.cell(row=row, column=ci, value=val)
                c.fill = platform_fills[idx]; c.font = NF; c.alignment = CA
            row += 1

        for ci in range(1, 7):
            ws_c.column_dimensions[get_column_letter(ci)].width = 14
        log.info(f"  [跨平台对比] 写入完成")

    # ── 评价分析sheet ──
    if review_data and review_data.get("total", 0) > 0:
        rname = "评价分析"
        if rname in wb.sheetnames:
            del wb[rname]
        ws_r = wb.create_sheet(rname)
        PURP = PatternFill("solid", fgColor="8E44AD")
        GRN  = PatternFill("solid", fgColor="27AE60")
        ORG  = PatternFill("solid", fgColor="E67E22")
        RED2 = PatternFill("solid", fgColor="2A5298")

        ws_r.cell(row=1, column=1, value=f"评价数量分析 - {keyword}").fill = PURP
        ws_r.cell(row=1, column=1).font = WF
        ws_r.merge_cells("A1:D1")

        # 总览统计
        r = 2
        for label, val in [
            ("总评价数", review_data.get("total", 0)),
            ("平均评价数", review_data.get("avg", 0)),
            ("最高评价数", review_data.get("max", 0)),
        ]:
            ws_r.cell(r, 1, label).font = WF
            ws_r.cell(r, 2, val).font = NF
            r += 1

        # Top10 评价最多商品
        r += 1
        ws_r.cell(r, 1, "📊 评价数量 Top10").fill = GRN
        ws_r.cell(r, 1).font = WF
        ws_r.merge_cells(f"A{r}:D{r}")
        r += 1
        for h, ci in [("商品标题", 1), ("店铺名称", 2), ("售价(元)", 3), ("评价数", 4)]:
            ws_r.cell(r, ci, h).fill = RED2
            ws_r.cell(r, ci).font = WF
        r += 1
        for item in review_data.get("top_review", []):
            ws_r.cell(r, 1, str(item.get("商品标题",""))[:30]).font = NF
            ws_r.cell(r, 2, str(item.get("店铺名称",""))).font = NF
            ws_r.cell(r, 3, item.get("当前售价(元)","")).font = NF
            ws_r.cell(r, 4, item.get("评价数量","")).font = NF
            r += 1

        # 口碑转化率 Top5
        r += 1
        ws_r.cell(r, 1, "⭐ 口碑转化率 Top5（评价/销量）").fill = ORG
        ws_r.cell(r, 1).font = WF
        ws_r.merge_cells(f"A{r}:D{r}")
        r += 1
        for h, ci in [("商品标题", 1), ("店铺名称", 2), ("评价数", 3), ("月销量", 4)]:
            ws_r.cell(r, ci, h).fill = RED2
            ws_r.cell(r, ci).font = WF
        r += 1
        for item in review_data.get("top_ratio", []):
            ws_r.cell(r, 1, str(item.get("商品标题",""))[:30]).font = NF
            ws_r.cell(r, 2, str(item.get("店铺名称",""))).font = NF
            ws_r.cell(r, 3, item.get("评价数量","")).font = NF
            ws_r.cell(r, 4, item.get("月销量(件)","")).font = NF
            r += 1

        ws_r.column_dimensions["A"].width = 35
        ws_r.column_dimensions["B"].width = 20
        ws_r.column_dimensions["C"].width = 12
        ws_r.column_dimensions["D"].width = 12
        log.info(f"  [评价分析] 写入完成（Top10评价商品 + 口碑转化率）")

    wb.save(filepath)
    log.info(f"  [多平台] 所有数据已写入: {filepath}")


def run_multiplatform(driver, cfg, log, stop_event=None, wait_login_callback=None):
    """多平台采集主函数，返回各平台数据"""
    keep_minimized(driver)
    keyword   = cfg["keyword"]
    pages     = cfg.get("pages", 3)
    filepath  = cfg["output_file"]
    platforms = cfg.get("platforms", [])

    jd_items  = []
    reviews   = []
    review_data = {}

    # 记录当前淘宝数据（已存在）
    tb_items = []
    try:
        import pandas as pd
        if os.path.exists(filepath):
            df = pd.read_excel(filepath, sheet_name="采集数据")
            if not df.empty:
                # 把所有可能的NaN列转成安全类型，避免"in"操作崩溃
                for col in ["商品链接", "商品标题", "店铺名称", "店铺类型"]:
                    if col in df.columns:
                        df[col] = df[col].fillna("").astype(str)
                for col in ["当前售价(元)", "月销量(件)", "原价(元)"]:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
                tb_items = df.to_dict("records")
    except Exception as e:
        log.warning(f"  [多平台] 读取淘宝数据失败: {e}")

    def _switch_and_check(platform_name, handle, domain_keywords, login_keywords):
        """切换标签页，检测是否登录，返回是否成功"""
        if not handle:
            log.warning(f"  [{platform_name}] 没有handle，跳过")
            return False
        try:
            driver.switch_to.window(handle)
            keep_minimized(driver)
            time.sleep(1)
            cur_url   = safe_url(driver).lower()
            cur_title = safe_title(driver)
            log.info(f"  [{platform_name}] 切换后: {cur_title[:50]} | {cur_url[:55]}")
            # 确认在正确平台
            if not any(k in cur_url or k in cur_title for k in domain_keywords):
                log.warning(f"  [{platform_name}] 不是目标平台，跳过")
                return False
            # 检测错误页/404
            if "error2.aspx" in cur_url or "您所访问的页面不存在" in cur_title:
                log.warning(f"  [{platform_name}] 检测到错误页，尝试跳转搜索页...")
                return True   # 继续执行，scrape_jd内部会跳转搜索页
            # 检测登录页
            if any(k in cur_title or k in cur_url for k in login_keywords):
                log.warning(f"  [{platform_name}] 检测到登录页，等待用户登录...")
                if wait_login_callback:
                    wait_login_callback(platform_name)
                time.sleep(3)
            log.info(f"  [{platform_name}] ✅ 准备采集")
            return True
        except Exception as e:
            log.warning(f"  [{platform_name}] 切换失败: {e}")
            return False

    # 京东采集（直接使用已有的京东标签页）
    if "jd" in platforms:
        log.info("\n[多平台] 切换到京东标签页...")

        # 扫描找到京东标签页
        jd_handle_fresh = None
        for h in driver.window_handles:
            try:
                driver.switch_to.window(h)
                time.sleep(0.3)
                u = safe_url(driver).lower()
                t = safe_title(driver)
                if "jd.com" in u or "360buy" in u:
                    jd_handle_fresh = h
                    log.info(f"  [京东] 找到京东标签: {t[:40]} | {u[:50]}")
                    break
            except Exception:
                pass

        if jd_handle_fresh:
            driver.switch_to.window(jd_handle_fresh)
            keep_minimized(driver)
            time.sleep(0.5)
            jd_items = scrape_jd(driver, keyword, pages, log, stop_event, use_current=True)
        else:
            # 没有京东标签，在当前标签跳转
            log.info("  [京东] 未找到京东标签，直接跳转...")
            jd_items = scrape_jd(driver, keyword, pages, log, stop_event, use_current=False)

        rsleep(3, 5, stop_event)



    # 热门关键词采集
    if cfg.get("scrape_hotwords"):
        log.info("\n[多平台] 开始采集热门关键词...")
        hotwords = scrape_hot_keywords(driver, keyword, log, stop_event)
        cfg["_hotwords"] = hotwords   # 暂存，save时写入
        # Hot keywords may have navigated tabs - no need to restore,
        # the JD tab URL doesn't matter after scraping is done
    else:
        cfg["_hotwords"] = {}

    # 评价分析：基于已采集数据的评价数量，不跳转页面
    if cfg.get("scrape_reviews") and tb_items:
        log.info("\n[多平台] 开始评价数量分析...")
        review_data = analyse_reviews_from_items(tb_items, log)

    # 保存到Excel
    # 写热词sheet
    if cfg.get("_hotwords"):
        save_hotwords_to_excel(filepath, keyword, cfg["_hotwords"], log)

    save_multiplatform_excel(tb_items, jd_items, review_data,
                              filepath, keyword, log)

    # 京东数据也写入历史归档（与淘宝的 save_history 共用同一个历史文件）
    if jd_items:
        try:
            # 给京东数据加上店铺类型标记，方便历史记录区分平台
            for item in jd_items:
                if not item.get("店铺类型"):
                    item["店铺类型"] = "JD"
            save_history(jd_items, filepath, keyword, log, platform="JD")
        except Exception as _he:
            log.warning(f"  [历史] 京东数据归档异常: {_he}")

    return jd_items, [], review_data
