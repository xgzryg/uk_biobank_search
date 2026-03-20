# UK Biobank 文献检索与Excel导出

## 概述
从UK Biobank Showcase网站(https://biobank.ndph.ox.ac.uk/showcase/search.cgi)检索特定主题的文献和申请信息，并输出为格式化Excel文件。

## 适用场景
- 检索UK Biobank相关的学术文献
- 检索UK Biobank数据申请项目
- 整理生物医学研究主题的相关资料

## 前置要求
- Python环境已安装以下库：
  - `playwright` - 用于网页自动化
  - `beautifulsoup4` - 用于HTML解析
  - `openpyxl` - 用于Excel文件生成

## 操作步骤

### 1. 安装依赖
```bash
pip install playwright beautifulsoup4 openpyxl
playwright install chromium
```

### 2. 创建检索脚本
创建Python脚本 `uk_biobank_search.py`：

```python
import asyncio
import re
from html import unescape
from pathlib import Path
from playwright.async_api import async_playwright
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

async def search_uk_biobank(topic: str, output_path: str) -> dict:
    """
    Search UK Biobank for a topic and return results

    Args:
        topic: Search term (e.g., "diabetes", "cancer", "cardiovascular")
        output_path: Output Excel file path

    Returns:
        Dictionary with publications and applications counts
    """
    url = "https://biobank.ndph.ox.ac.uk/showcase/search.cgi"

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        # Navigate to search page
        await page.goto(url)

        # Fill in search term
        await page.fill('input[name="searchterm"]', topic)

        # Click search button
        await page.click('input[type="submit"][value="Search"]')

        # Wait for results to load
        await page.wait_for_load_state("networkidle")

        # Get page content
        html_content = await page.content()

        await browser.close()

    # Decode HTML entities
    html_content = unescape(html_content)
    soup = BeautifulSoup(html_content, 'html.parser')

    # Parse publications
    publications = []
    pub_rows = soup.find_all('tr', id=re.compile(r'^p\d+$'))

    for row in pub_rows:
        cells = row.find_all('td')
        if len(cells) >= 5:
            pub_id = cells[0].get_text(strip=True)
            title = cells[1].get_text(strip=True)
            authors = cells[2].get_text(strip=True)
            year = cells[3].get_text(strip=True)
            journal = cells[4].get_text(strip=True)
            publications.append([pub_id, title, authors, year, journal])

    # Parse applications
    applications = []
    app_rows = soup.find_all('tr', id=re.compile(r'^a\d+$'))

    for row in app_rows:
        cells = row.find_all('td')
        if len(cells) >= 2:
            app_id = cells[0].get_text(strip=True)
            title = cells[1].get_text(strip=True)
            applications.append([app_id, title])

    # Create Excel workbook
    wb = Workbook()

    # Header styling
    header_font_white = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

    # Publications sheet
    ws_pub = wb.active
    ws_pub.title = "Publications"
    ws_pub.append(["Pub ID", "Title", "Authors", "Year", "Journal"])
    for cell in ws_pub[1]:
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for pub in publications:
        ws_pub.append(pub)
    ws_pub.column_dimensions['A'].width = 10
    ws_pub.column_dimensions['B'].width = 80
    ws_pub.column_dimensions['C'].width = 40
    ws_pub.column_dimensions['D'].width = 8
    ws_pub.column_dimensions['E'].width = 40

    # Applications sheet
    ws_app = wb.create_sheet("Applications")
    ws_app.append(["Application ID", "Title"])
    for cell in ws_app[1]:
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for app in applications:
        ws_app.append(app)
    ws_app.column_dimensions['A'].width = 15
    ws_app.column_dimensions['B'].width = 100

    # Save workbook
    wb.save(output_path)

    return {
        "publications_count": len(publications),
        "applications_count": len(applications),
        "output_file": output_path
    }

# Command line interface
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python uk_biobank_search.py <topic> <output_excel_path>")
        print("Example: python uk_biobank_search.py diabetes C:\\Users\\Lenovo\\Desktop\\diabetes.xlsx")
        sys.exit(1)

    topic = sys.argv[1]
    output_path = sys.argv[2]

    print(f"Searching UK Biobank for: {topic}")
    result = asyncio.run(search_uk_biobank(topic, output_path))

    print(f"\nResults:")
    print(f"  Publications: {result['publications_count']}")
    print(f"  Applications: {result['applications_count']}")
    print(f"  Output file: {result['output_file']}")
```

### 3. 使用方法

**命令行方式：**
```bash
python uk_biobank_search.py "diabetes" "C:\Users\Lenovo\Desktop\diabetes.xlsx"
python uk_biobank_search.py "cancer" "C:\Users\Lenovo\Desktop\cancer.xlsx"
python uk biobank_search.py "cardiovascular" "C:\Users\Lenovo\Desktop\cardio.xlsx"
```

**Python代码方式：**
```python
import asyncio
from uk_biobank_search import search_uk_biobank

result = asyncio.run(search_uk_biobank(
    topic="diabetes",
    output_path=r"C:\Users\Lenovo\Desktop\diabetes.xlsx"
))

print(f"Found {result['publications_count']} publications")
print(f"Found {result['applications_count']} applications")
```

## 输出格式
Excel文件包含两个工作表：

### Publications（文献）
| 列名 | 说明 |
|------|------|
| Pub ID | 出版物ID |
| Title | 文献标题 |
| Authors | 作者 |
| Year | 出版年份 |
| Journal | 期刊名称 |

### Applications（申请）
| 列名 | 说明 |
|------|------|
| Application ID | 申请ID |
| Title | 申请标题 |

## 注意事项
1. UK Biobank网站可能需要几秒钟加载搜索结果
2. 某些特殊字符可能需要额外的HTML实体解码
3. 搜索结果数量可能有限制，默认返回约100条文献和部分申请

## 错误处理
如果遇到错误，请检查：
- 网络连接是否正常
- Playwright是否正确安装
- 输出路径是否有写入权限
