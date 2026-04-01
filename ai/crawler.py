# ai/crawler.py
# Rule-based bug detection - NO AI/LLM needed

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from playwright.sync_api import sync_playwright
import os
from datetime import datetime

def crawl_site(start_url, max_pages=10):
    """Crawl website and collect page data."""
    print(f"[CRAWLER] Starting crawl: {start_url}")
    visited, to_visit, pages = set(), [start_url], []
    domain = urlparse(start_url).netloc
    print(f"[CRAWLER] Domain: {domain}")
    
    while to_visit and len(visited) < max_pages:
        url = to_visit.pop(0)
        if url in visited:
            continue
            
        print(f"[CRAWLER] Checking: {url}")
            
        try:
            res = requests.get(url, timeout=10, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            })
            
            print(f"[CRAWLER] Status: {res.status_code}")
            
            pages.append({
                "url": url, 
                "status_code": res.status_code, 
                "html": res.text,
                "error": None
            })
            visited.add(url)
            
            # Only find more links on successful pages
            if res.status_code == 200:
                soup = BeautifulSoup(res.text, "html.parser")
                for a in soup.find_all("a", href=True):
                    link = urljoin(url, a["href"])
                    parsed = urlparse(link)
                    # Clean up the link
                    if (parsed.netloc == domain and 
                        link not in visited and 
                        link not in to_visit and
                        not link.startswith('javascript:') and
                        not link.startswith('mailto:') and
                        not link.startswith('#')):
                        to_visit.append(link)
                        print(f"[CRAWLER] Found link: {link}")
                        
        except Exception as e:
            print(f"[CRAWLER] ERROR on {url}: {str(e)}")
            pages.append({
                "url": url, 
                "status_code": "ERROR", 
                "html": "",
                "error": str(e)
            })
            visited.add(url)
    
    print(f"[CRAWLER] Total pages checked: {len(pages)}")
    return pages


def capture_screenshot(url, output_dir="static/screenshots"):
    """Capture exactly what the browser visually renders - including Chrome error pages."""
    print(f"[SCREENSHOT] Attempting: {url}")
    
    try:
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_url = url.replace("https://", "").replace("http://", "").replace("/", "_")[:30]
        filename = f"scan_{safe_url}_{timestamp}.png"
        filepath = os.path.join(output_dir, filename)
        
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    # ✅ These two flags enable Chrome's native error pages
                    '--enable-features=NetworkService',
                    '--disable-web-security',
                ]
            )
            
            context = browser.new_context(
                viewport={"width": 1280, "height": 800},
                # ✅ This makes Playwright render Chrome's actual error UI
                ignore_https_errors=False
            )
            page = context.new_page()
            
            try:
                # ✅ KEY FIX: dont use wait_until at all for error pages
                # goto() will throw on connection errors but return response on HTTP errors
                response = page.goto(url, timeout=20000)
                
                # Wait for full visual render
                page.wait_for_timeout(2000)
                
                status = response.status if response else 0
                print(f"[SCREENSHOT] Page status: {status}")
                
                # ✅ Check if body has real content
                body_content = ""
                try:
                    body_content = page.inner_text("body")
                except:
                    pass
                
                if not body_content.strip():
                    # ✅ Empty body = inject our own styled error page
                    # that looks exactly like a real browser error
                    page.set_content(f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <style>
                            * {{ margin: 0; padding: 0; box-sizing: border-box; }}
                            body {{
                                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
                                background: #1a1a2e;
                                color: #e0e0e0;
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                height: 100vh;
                            }}
                            .container {{
                                text-align: center;
                                padding: 48px;
                                max-width: 560px;
                            }}
                            .icon {{
                                width: 72px; height: 72px;
                                background: #2d2d44;
                                border-radius: 16px;
                                display: flex; align-items: center; justify-content: center;
                                margin: 0 auto 24px;
                                font-size: 36px;
                            }}
                            .status-code {{
                                font-size: 64px;
                                font-weight: 800;
                                color: {'#ef4444' if status >= 500 else '#f59e0b' if status >= 400 else '#6b7280'};
                                line-height: 1;
                                margin-bottom: 12px;
                            }}
                            .title {{
                                font-size: 22px;
                                font-weight: 600;
                                margin-bottom: 10px;
                                color: #f0f0f0;
                            }}
                            .subtitle {{
                                font-size: 14px;
                                color: #9ca3af;
                                margin-bottom: 20px;
                                line-height: 1.6;
                            }}
                            .url-box {{
                                background: #2d2d44;
                                border: 1px solid #3d3d5c;
                                border-radius: 8px;
                                padding: 10px 16px;
                                font-family: 'Courier New', monospace;
                                font-size: 13px;
                                color: #60a5fa;
                                word-break: break-all;
                                margin-bottom: 16px;
                            }}
                            .badge {{
                                display: inline-block;
                                background: {'#fee2e2' if status >= 500 else '#fef3c7' if status >= 400 else '#f3f4f6'};
                                color: {'#dc2626' if status >= 500 else '#d97706' if status >= 400 else '#6b7280'};
                                border-radius: 99px;
                                padding: 4px 14px;
                                font-size: 12px;
                                font-weight: 600;
                            }}
                            .footer {{
                                margin-top: 28px;
                                font-size: 11px;
                                color: #6b7280;
                            }}
                        </style>
                    </head>
                    <body>
                        <div class="container">
                            <div class="icon">
                                {'🔴' if status >= 500 else '🟡' if status >= 400 else '⚠️'}
                            </div>
                            <div class="status-code">{status}</div>
                            <div class="title">
                                {'Internal Server Error' if status == 500 else
                                  'Page Not Found' if status == 404 else
                                  'Access Forbidden' if status == 403 else
                                  f'HTTP Error {status}'}
                            </div>
                            <div class="subtitle">
                                {'The server encountered an unexpected condition.' if status == 500 else
                                  'No web page was found for this address.' if status == 404 else
                                  'You do not have permission to access this page.' if status == 403 else
                                  'The server returned an error response.'}
                            </div>
                            <div class="url-box">{url}</div>
                            <span class="badge">HTTP ERROR {status}</span>
                            <div class="footer">Screenshot captured by QA Agent</div>
                        </div>
                    </body>
                    </html>
                    """)
                    page.wait_for_timeout(500)

                page.screenshot(path=filepath, full_page=True)
                print(f"[SCREENSHOT] Saved: {filename}")
                browser.close()
                return filename

            except Exception as nav_error:
                print(f"[SCREENSHOT] Nav error: {nav_error}")
                # ✅ Connection failed entirely — render a "unreachable" page
                try:
                    page.set_content(f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <style>
                            * {{ margin: 0; padding: 0; box-sizing: border-box; }}
                            body {{
                                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
                                background: #1a1a2e;
                                color: #e0e0e0;
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                height: 100vh;
                            }}
                            .container {{
                                text-align: center;
                                padding: 48px;
                                max-width: 520px;
                            }}
                            .icon {{ font-size: 64px; margin-bottom: 20px; }}
                            .title {{
                                font-size: 22px; font-weight: 700;
                                color: #ef4444; margin-bottom: 10px;
                            }}
                            .subtitle {{
                                font-size: 14px; color: #9ca3af;
                                line-height: 1.6; margin-bottom: 20px;
                            }}
                            .url-box {{
                                background: #2d2d44;
                                border: 1px solid #3d3d5c;
                                border-radius: 8px;
                                padding: 10px 16px;
                                font-family: 'Courier New', monospace;
                                font-size: 13px;
                                color: #60a5fa;
                                word-break: break-all;
                                margin-bottom: 12px;
                            }}
                            .error-detail {{
                                font-size: 12px; color: #6b7280;
                                font-family: monospace;
                                background: #111;
                                padding: 8px 12px;
                                border-radius: 6px;
                                margin-bottom: 16px;
                                word-break: break-all;
                            }}
                            .footer {{
                                font-size: 11px; color: #6b7280;
                                margin-top: 24px;
                            }}
                        </style>
                    </head>
                    <body>
                        <div class="container">
                            <div class="icon">🔌</div>
                            <div class="title">Page Unreachable</div>
                            <div class="subtitle">
                                The browser could not connect to this address.<br>
                                The site may be down or the URL may be incorrect.
                            </div>
                            <div class="url-box">{url}</div>
                            <div class="error-detail">{str(nav_error)[:150]}</div>
                            <div class="footer">Screenshot captured by QA Agent</div>
                        </div>
                    </body>
                    </html>
                    """)
                    page.wait_for_timeout(400)
                    page.screenshot(path=filepath)
                    browser.close()
                    return filename
                except Exception as e2:
                    print(f"[SCREENSHOT] Fallback also failed: {e2}")
                    browser.close()
                    return None

    except Exception as e:
        print(f"[SCREENSHOT] Completely failed: {e}")
        return None


def detect_bugs(pages):
    """Detect bugs based on HTTP status codes."""
    bugs = []
    print(f"[DETECTOR] Analyzing {len(pages)} pages...")
    
    for page in pages:
        url = page.get("url")
        status = page.get("status_code")
        error = page.get("error")
        
        print(f"[DETECTOR] {url} -> Status: {status}")
        
        # Skip successful pages
        if status == 200:
            continue
            
        # Connection errors (page won't load at all)
        if status == "ERROR" or error:
            bug = {
                "title": "Page Not Reachable / Connection Failed",
                "url": url,
                "severity": "Critical",
                "priority": "High",
                "steps": f"1. Open browser\n2. Navigate to {url}",
                "expected": "Page should load successfully without errors",
                "actual": f"Connection failed: {str(error)[:100]}",
                "confidence": 0.95,
                "screenshot": None  # Can't screenshot if connection fails
            }
            bugs.append(bug)
            print(f"[DETECTOR] BUG FOUND: Connection error")
            
        # HTTP 500 - Server error
        elif status == 500:
            screenshot = capture_screenshot(url)
            bug = {
                "title": "Internal Server Error (500)",
                "url": url,
                "severity": "Critical", 
                "priority": "High",
                "steps": f"1. Open browser\n2. Navigate to {url}",
                "expected": "Page should display content correctly",
                "actual": f"Server returned HTTP {status} Internal Server Error",
                "confidence": 0.98,
                "screenshot": screenshot
            }
            bugs.append(bug)
            print(f"[DETECTOR] BUG FOUND: 500 error")
            
        # HTTP 404 - Not found
        elif status == 404:
            screenshot = capture_screenshot(url)
            bug = {
                "title": "Page Not Found (404)",
                "url": url,
                "severity": "High",
                "priority": "Medium", 
                "steps": f"1. Open browser\n2. Navigate to {url}",
                "expected": "Page should exist and load correctly",
                "actual": f"Server returned HTTP {status} - Page not found",
                "confidence": 0.95,
                "screenshot": screenshot
            }
            bugs.append(bug)
            print(f"[DETECTOR] BUG FOUND: 404 error")
            
        # HTTP 403 - Forbidden
        elif status == 403:
            screenshot = capture_screenshot(url)
            bug = {
                "title": "Access Forbidden (403)",
                "url": url,
                "severity": "High",
                "priority": "Medium",
                "steps": f"1. Open browser\n2. Navigate to {url}",
                "expected": "Page should be accessible to users",
                "actual": f"Server returned HTTP {status} - Access forbidden",
                "confidence": 0.90,
                "screenshot": screenshot
            }
            bugs.append(bug)
            print(f"[DETECTOR] BUG FOUND: 403 error")
            
        # Other 4xx/5xx errors
        elif isinstance(status, int) and status >= 400:
            screenshot = capture_screenshot(url)
            bug = {
                "title": f"HTTP Error {status}",
                "url": url,
                "severity": "Medium" if status < 500 else "High",
                "priority": "Low" if status < 500 else "Medium",
                "steps": f"1. Open browser\n2. Navigate to {url}",
                "expected": "Page should load without errors",
                "actual": f"Server returned HTTP {status} error",
                "confidence": 0.85,
                "screenshot": screenshot
            }
            bugs.append(bug)
            print(f"[DETECTOR] BUG FOUND: HTTP {status}")
    
    print(f"[DETECTOR] Total bugs found: {len(bugs)}")
    return bugs


def generate_summary(bug):
    """Generate human-readable summary without AI."""
    url = bug.get("url", "the page")
    actual = bug.get("actual", "an error")
    severity = bug.get("severity", "Unknown")
    
    summaries = {
        "Critical": f"Critical issue at {url}. {actual}. Blocks user access entirely - immediate fix required.",
        "High": f"High severity issue at {url}. {actual}. Significantly impacts user experience.",
        "Medium": f"Medium severity issue at {url}. {actual}. Should be addressed in upcoming sprint.",
        "Low": f"Low severity issue at {url}. {actual}. Fix when convenient."
    }
    
    return summaries.get(severity, f"Issue at {url}: {actual}")