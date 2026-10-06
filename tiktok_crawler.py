"""
TikTok Keyword Crawler v3
Search hashtags -> collect videos, author, likes, comments -> export Excel
Edit keywords.txt to change search terms (hashtags without #).
"""
import asyncio
import re
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from playwright.async_api import async_playwright

# ============ Config ============
KEYWORDS_FILE = Path(__file__).parent / "keywords.txt"
MAX_RESULTS = 100
OUTPUT_DIR = Path(__file__).parent / "导出结果"
HEADLESS = False
# ================================


def load_keywords():
    if KEYWORDS_FILE.exists():
        text = KEYWORDS_FILE.read_text(encoding="utf-8").strip()
        return [k.strip() for k in re.split(r"[,，\n]", text) if k.strip()]
    return ["leatherbag"]


async def crawl_tiktok(page, tag):
    url = f"https://www.tiktok.com/tag/{tag}"
    print(f"  Opening {url}", flush=True)
    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(12)

    # Close login popup
    try:
        close = await page.query_selector('[aria-label="Close"], button:has-text("Not now")')
        if close:
            await close.click()
            await asyncio.sleep(2)
    except:
        pass

    # Scroll to load more
    for _ in range(15):
        await page.evaluate("window.scrollBy(0, 1200)")
        await asyncio.sleep(1.5)

    # Extract from DOM: each video card
    cards = await page.evaluate("""
        () => {
            const links = document.querySelectorAll('a[href*="/video/"]');
            const seen = new Set();
            const out = [];
            for (const a of links) {
                const href = a.href.split('?')[0];
                if (seen.has(href)) continue;
                seen.add(href);
                // climb up to find card container
                let node = a;
                let cardText = '';
                for (let j = 0; j < 8; j++) {
                    if (node.parentElement) node = node.parentElement;
                    if (node.innerText && node.innerText.length > 20) {
                        cardText = node.innerText;
                        if (node.querySelectorAll('a[href*="/video/"]').length > 2) break;
                    }
                }
                out.push({href, cardText: cardText.slice(0, 600)});
            }
            return out;
        }
    """)

    results = []
    for c in cards[:MAX_RESULTS]:
        href = c["href"]
        # author from URL
        m = re.search(r"tiktok\.com/@([^/]+)/video/", href)
        author = m.group(1) if m else ""

        # Parse card text for likes (heart icon number)
        text = c.get("cardText", "")
        likes = 0
        lk = re.search(r"([\d.,]+[KM]?)\s*$", text.split("\n")[-1])
        if lk:
            num = lk.group(1)
            likes = float(num.replace(",", "").rstrip("K")) * (1000 if "K" in num.upper() else 1)
            if "M" in num.upper():
                likes = float(num.replace(",", "").rstrip("M")) * 1_000_000
            likes = int(likes)

        # Caption: longest line after author
        lines = [l.strip() for l in text.split("\n") if l.strip() and l.strip() != author]
        caption = " ".join(lines[:3])[:300]

        results.append({
            "author": author,
            "caption": caption,
            "likes": likes,
            "url": href,
        })

    print(f"  Found {len(results)} videos", flush=True)
    return results


def save_to_excel(all_data, keyword):
    OUTPUT_DIR.mkdir(exist_ok=True)
    all_data = sorted(all_data, key=lambda x: x.get("likes", 0), reverse=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "TikTok Videos"

    headers = ["Keyword", "Author", "Caption", "Likes", "URL", "Collected At"]
    ws.append(headers)
    for v in all_data:
        ws.append([keyword, v["author"], v["caption"], v["likes"], v["url"],
                    datetime.now().strftime("%Y-%m-%d %H:%M")])

    for i, w in enumerate([15, 25, 60, 10, 50, 18], 1):
        ws.column_dimensions[chr(64 + i)].width = w

    fname = OUTPUT_DIR / f"TikTok_{keyword}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(str(fname))
    print(f"  Saved: {fname}", flush=True)


async def main():
    keywords = load_keywords()
    print("=" * 50, flush=True)
    print("  TikTok Keyword Crawler v3", flush=True)
    print("=" * 50, flush=True)
    print(f"Keywords: {', '.join(keywords)}", flush=True)
    print()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS, channel="chrome")
        ctx = await browser.new_context(
            locale="en-US",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 900},
        )
        page = await ctx.new_page()

        total = 0
        for kw in keywords:
            print(f"=== #{kw} ===", flush=True)
            try:
                videos = await crawl_tiktok(page, kw)
                if videos:
                    save_to_excel(videos, kw)
                    for v in videos[:3]:
                        print(f"    @{v['author']} | {v['likes']} likes", flush=True)
                    total += len(videos)
            except Exception as e:
                print(f"  Error: {e}", flush=True)

        print(f"\nDone! Total {total} videos.", flush=True)
        print(f"Output: {OUTPUT_DIR.resolve()}", flush=True)
        await browser.close()
        input("Press Enter to exit...")


if __name__ == "__main__":
    asyncio.run(main())
