# TaobaoScraper MCP Server

An MCP (Model Context Protocol) server for scraping product data from **Taobao/Tmall** and **JD.com** (Jingdong). Provides 8 tools that can be used directly in Claude Desktop or Claude Code.

> **Platform**: Windows (requires Chrome browser)
> **Language**: Python 3.10+

## Features

- **Multi-platform scraping** - Taobao, Tmall, JD.com in one tool
- **8 MCP tools** - Scraping, task management, notifications, system control
- **Hot search words** - Discover trending keywords and market insights
- **Excel export** - Auto-saves results as `.xlsx` files
- **Notification push** - WeChat (ServerChan / PushPlus) and Email alerts
- **Dual transport** - stdio (local) and SSE (remote) protocols
- **Trilingual GUI** - Chinese / English / Korean interface (optional)

## Architecture

```
Claude Desktop / Claude Code
        |
   MCP Server (stdio or SSE)
        |
   FastAPI Backend (:8000)
        |
   Selenium + Chrome (:9222)
```

## Quick Start

### 1. Prerequisites

- Python 3.10+
- Google Chrome browser
- Windows OS

### 2. Install

```bash
git clone https://github.com/jhongjun1981/taobao-scraper-mcp.git
cd taobao-scraper-mcp

# Install MCP server dependencies
pip install -r requirements_mcp.txt

# Install API backend dependencies
pip install -r requirements_api.txt
```

### 3. Configure

```bash
cp .env.example .env
# Edit .env to set your SCRAPER_API_KEY
```

### 4. Start Chrome with debug port

```bash
# Close all Chrome instances first, then:
START_GUI.bat
# Or manually:
chrome.exe --remote-debugging-port=9222 --user-data-dir="%LOCALAPPDATA%\Google\Chrome\Debug Profile"
```

### 5. Login to Taobao (first time only)

Open Chrome and login to your Taobao account. Cookies will persist in the debug profile.

### 6. Start API Backend

```bash
python run_api.py
```

### 7. Configure MCP Server

**For Claude Code** - Add to `.claude/settings.json`:

```json
{
  "mcpServers": {
    "taobao-scraper": {
      "command": "python",
      "args": ["run_mcp.py"],
      "cwd": "/path/to/taobao-scraper-mcp"
    }
  }
}
```

**For Claude Desktop** - Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "taobao-scraper": {
      "command": "python",
      "args": ["run_mcp.py"],
      "cwd": "C:\\path\\to\\taobao-scraper-mcp"
    }
  }
}
```

**SSE mode (remote)**:

```bash
python run_mcp.py --transport sse --port 8001
```

```json
{
  "mcpServers": {
    "taobao-scraper": {
      "type": "sse",
      "url": "http://your-server:8001/sse"
    }
  }
}
```

## Tools Reference

### Scraping Tools

#### `scrape_products`

Scrape product data from e-commerce platforms.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `keyword` | string | *required* | Search keyword |
| `platform` | string | `"taobao"` | `"taobao"` / `"jd"` / `"multi"` |
| `pages` | int | `3` | Number of pages to scrape |
| `sort_by` | string | `"sale"` | Sort method |
| `tmall_only` | bool | `false` | Tmall products only |
| `exact_match` | bool | `false` | Exact keyword match |
| `price_min` | float | `0` | Min price filter |
| `price_max` | float | `0` | Max price filter |
| `wait_for_result` | bool | `true` | Wait for completion |
| `timeout` | int | `300` | Timeout in seconds |

**Example**: "Scrape running shoes from both Taobao and JD"

#### `scrape_hotwords`

Scrape trending/related search keywords for market analysis.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `keyword` | string | *required* | Base keyword |
| `wait_for_result` | bool | `true` | Wait for completion |
| `timeout` | int | `120` | Timeout in seconds |

### Task Management Tools

#### `list_tasks`

List recent scraping tasks with status (pending/running/completed/failed).

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `limit` | int | `20` | Max tasks to return |

#### `get_task`

Get detailed status of a specific task. Set `include_result=true` to get full results.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `task_id` | string | *required* | Task ID |
| `include_result` | bool | `false` | Include result data |

#### `cancel_task`

Cancel a running or pending task.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `task_id` | string | *required* | Task ID to cancel |

### Export Tools

#### `list_files`

List all exported Excel data files with filename, size, and modification time.

### Notification Tools

#### `send_notification`

Send push notifications via WeChat or Email.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `channel` | string | *required* | `"wechat"` / `"email"` / `"test"` |
| `title` | string | `""` | Notification title |
| `content` | string | `""` | Notification body |
| `wx_type` | string | `"server_chan"` | `"server_chan"` or `"pushplus"` |

### System Tools

#### `system_status`

Check system health or restart Chrome.

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `action` | string | `"check"` | `"check"` or `"restart_chrome"` |

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `SCRAPER_API_KEY` | `changeme-your-secret-key` | API authentication key |
| `SCRAPER_API_URL` | `http://localhost:8000` | FastAPI backend URL |
| `SCRAPER_CHROME_PORT` | `9222` | Chrome debug port |
| `MCP_SSE_HOST` | `0.0.0.0` | SSE listen address |
| `MCP_SSE_PORT` | `8001` | SSE listen port |
| `MCP_HTTP_TIMEOUT` | `30` | HTTP request timeout |

## Data Output

Scraped data is automatically saved as Excel files containing:

- Product title, image URL, price
- Monthly sales volume, review count
- Shop name, shop type, shop region
- Product URL, scrape timestamp

## Important Notes

- **Taobao requires login** - You must login to your own Taobao account in Chrome first
- **JD works without login** - JD scraping does not require authentication
- **One task at a time** - Chrome can only handle one scraping task concurrently
- **Windows only** - The Chrome automation layer uses Windows-specific APIs
- **Respect rate limits** - Excessive scraping may trigger anti-bot protection

## License

MIT License - see [LICENSE](LICENSE) file.
