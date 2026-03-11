# TaobaoScraper API 部署说明

## 一、环境要求

| 项目 | 要求 | 下载地址 |
|------|------|---------|
| 操作系统 | Windows 10 / 11（含 Windows Server） | — |
| Python | 3.10 及以上 | https://www.python.org/downloads/ |
| Google Chrome | 任意版本（程序自动匹配驱动） | https://www.google.com/chrome/ |
| 网络 | 需连接互联网 | — |

> **注意：** 本工具依赖 Windows API 和 Chrome 浏览器，暂不支持 Linux 部署。

---

## 二、快速部署（3 步）

### 第 1 步：配置环境变量

复制 `.env.example` 为 `.env`，修改 API Key：

```
SCRAPER_API_KEY=your-secret-key-here
```

**必须修改！** 这是调用 API 的唯一凭证。

### 第 2 步：双击启动

双击 `START_API.bat`，脚本会自动：
- 加载 `.env` 配置
- 安装 Python 依赖（首次）
- 启动 Chrome 浏览器（调试端口 9222）
- 启动 API 服务（端口 8000）

### 第 3 步：登录淘宝/京东

启动后会弹出 Chrome 窗口，**必须手动登录淘宝/京东账号**，否则无法采集。
登录成功后即可通过 API 调用采集功能。

---

## 三、API 使用

### 3.1 认证方式

所有接口都需要 API Key，支持两种传递方式：

```
# 方式一：Header（推荐）
X-API-Key: your-secret-key-here

# 方式二：Query 参数
?api_key=your-secret-key-here
```

### 3.2 交互式文档

启动后访问 http://localhost:8000/docs 可看到 Swagger UI，支持在线测试所有接口。

### 3.3 接口一览

#### 系统管理

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v1/system/health` | 健康检查 |
| GET | `/api/v1/system/chrome` | Chrome 状态 |
| POST | `/api/v1/system/chrome/restart` | 重启 Chrome |

#### 数据采集（异步，返回 task_id）

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/v1/scrape/taobao` | 淘宝采集 |
| POST | `/api/v1/scrape/jd` | 京东采集 |
| POST | `/api/v1/scrape/multiplatform` | 多平台采集（淘宝+京东+评价+热词） |
| POST | `/api/v1/reviews/scrape` | 采集单商品评价 |
| POST | `/api/v1/hotwords/scrape` | 热门关键词采集 |

#### 任务管理

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v1/tasks` | 列出所有任务 |
| GET | `/api/v1/tasks/{task_id}` | 查询任务状态 |
| GET | `/api/v1/tasks/{task_id}/result` | 获取任务结果 |
| POST | `/api/v1/tasks/{task_id}/cancel` | 取消任务 |

#### 文件导出

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/v1/export/files` | 列出 Excel 文件 |
| GET | `/api/v1/export/download/{filename}` | 下载文件 |

#### 通知推送

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/v1/notify/wechat` | 微信推送 |
| POST | `/api/v1/notify/email` | 邮件发送 |
| POST | `/api/v1/notify/test` | 测试通知配置 |

#### 数据分析（同步，直接返回结果）

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/v1/reviews/analyse` | 评价关键词分析 |

### 3.4 调用示例

#### 提交淘宝采集任务

```bash
curl -X POST http://localhost:8000/api/v1/scrape/taobao \
  -H "X-API-Key: your-secret-key-here" \
  -H "Content-Type: application/json" \
  -d '{"keyword": "天气丹两件套", "pages": 3, "sort_by": "sale"}'
```

返回：
```json
{"task_id": "taobao_20260310_093000_a1b2c3", "message": "任务已提交"}
```

#### 查询任务状态

```bash
curl http://localhost:8000/api/v1/tasks/taobao_20260310_093000_a1b2c3 \
  -H "X-API-Key: your-secret-key-here"
```

返回：
```json
{
  "task_id": "taobao_20260310_093000_a1b2c3",
  "task_type": "taobao",
  "status": "completed",
  "progress": "采集完成，共 120 条",
  "result_file": "天气丹两件套_data.xlsx"
}
```

#### 下载采集结果

```bash
curl -O http://localhost:8000/api/v1/export/download/天气丹两件套_data.xlsx \
  -H "X-API-Key: your-secret-key-here"
```

#### 发送微信推送

```bash
curl -X POST http://localhost:8000/api/v1/notify/wechat \
  -H "X-API-Key: your-secret-key-here" \
  -H "Content-Type: application/json" \
  -d '{"title": "采集完成", "content": "共采集120条数据", "wx_type": "pushplus", "wx_key": "your-token"}'
```

---

## 四、典型使用流程

```
1. 提交采集任务    →  POST /api/v1/scrape/taobao
2. 轮询任务状态    →  GET  /api/v1/tasks/{task_id}   （等待 status=completed）
3. 下载结果文件    →  GET  /api/v1/export/download/{filename}
4. （可选）推送通知 →  POST /api/v1/notify/wechat
```

---

## 五、注意事项

1. **必须登录账号：** Chrome 启动后需手动登录淘宝/京东，API 不会自动登录。
2. **单任务执行：** 同一时间只能运行一个采集任务（Chrome 单实例限制），后续任务会排队等待。
3. **保持 Chrome 运行：** 不要手动关闭 Chrome 窗口，API 服务依赖它工作。
4. **修改 API Key：** 务必修改默认 Key，否则任何人都能调用你的接口。
5. **防火墙放行：** 如需外网访问，需在 Windows 防火墙中放行 8000 端口。
6. **验证码处理：** 淘宝偶尔弹验证码时需手动完成，建议降低采集页数和频率。

---

## 六、文件清单

```
TaobaoScraper_v5.3/
├── START_API.bat          ← 一键启动 API 服务
├── START_GUI.bat          ← 一键启动 GUI 界面
├── .env.example           ← 环境变量模板（复制为 .env）
├── requirements_api.txt   ← API 全部依赖
├── run_api.py             ← API 启动脚本
├── scraper_core.py        ← 采集核心引擎
├── scraper_gui.py         ← GUI 界面（API 模式下不需要）
├── toggle_chrome.py       ← F9 快捷键工具（可选）
└── api/                   ← API 服务代码
    ├── main.py
    ├── config.py
    ├── auth.py
    ├── chrome_manager.py
    ├── models.py
    ├── task_manager.py
    └── routers/
        ├── scrape.py
        ├── reviews.py
        ├── hotwords.py
        ├── export.py
        ├── notify.py
        ├── tasks.py
        └── system.py
```
