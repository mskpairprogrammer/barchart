"""
Barchart Screenshot Utility
===========================

Captures screenshots of stock put-call ratio pages from Barchart.com.

WORKFLOW:
1. Find & focus browser window containing "barchart" in title
2. Search for stock symbol (if STOCK_SYMBOL is set in .env)
3. Click center of screen to ensure page focus
4. Scroll down to show chart area
5. Capture screenshot (retry if blank)
6. Save to screenshots folder with timestamp

All configuration is read from .env file.
"""

import pyautogui
import win32gui
import win32con
from datetime import datetime
from PIL import Image, ImageStat
from dotenv import load_dotenv
import os
import time

# Load environment variables
load_dotenv()

# Configure pyautogui failsafe
pyautogui.FAILSAFE = os.getenv('DISABLE_FAILSAFE', 'true').lower() != 'true'


def get_browser_keywords() -> list:
    """Get list of browser name keywords from env or defaults."""
    default = 'chrome,firefox,edge,safari,opera,brave,vivaldi'
    return [k.strip().lower() for k in os.getenv('BROWSER_KEYWORDS', default).split(',')]


def is_screenshot_blank(image: Image.Image) -> bool:
    """Check if screenshot is mostly white/blank using RGB mean analysis."""
    threshold = float(os.getenv('BLANK_THRESHOLD', '240'))
    try:
        stat = ImageStat.Stat(image)
        return all(m > threshold for m in stat.mean[:3])
    except Exception:
        return False


def bring_window_to_front(keyword: str) -> bool:
    """
    Find browser window by keyword and bring to foreground.
    Returns True if successful, False otherwise.
    """
    browser_keywords = get_browser_keywords()
    found_windows = []
    
    def enum_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if keyword.lower() in title.lower():
                is_browser = any(b in title.lower() for b in browser_keywords)
                safe_title = title.encode('ascii', 'replace').decode('ascii')
                found_windows.append((hwnd, safe_title, is_browser))
        return True
    
    win32gui.EnumWindows(enum_callback, None)
    
    # Filter to browser windows only
    browser_windows = [(hwnd, title) for hwnd, title, is_browser in found_windows if is_browser]
    
    if not browser_windows:
        if found_windows:
            print(f"  [WARN] No browser window found with '{keyword}'. Non-browser matches:")
            for _, title, _ in found_windows:
                print(f"         - {title}")
        else:
            print(f"  [WARN] No window found with '{keyword}'")
        return False
    
    hwnd, title = browser_windows[0]
    print(f"  [OK] Found: {title}")
    
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        return True
    except Exception as e:
        print(f"  [ERROR] {e}")
        return False


def search_stock(symbol: str) -> None:
    """Navigate to stock's put-call ratio page via URL bar."""
    wait = float(os.getenv('SEARCH_WAIT', '15.0'))
    
    # Check if symbol needs $ prefix (index symbols like VIX)
    index_symbols = [s.strip().upper() for s in os.getenv('INDEX_SYMBOLS', '').split(',') if s.strip()]
    url_symbol = f"%24{symbol}" if symbol.upper() in index_symbols else symbol
    
    # Build the direct URL for put-call ratios page
    url = f"https://www.barchart.com/stocks/quotes/{url_symbol}/put-call-ratios"
    
    print(f"[SEARCH] Navigating to: {symbol}" + (" (index)" if symbol.upper() in index_symbols else ""))
    
    # Use Ctrl+L to focus address bar, then type URL
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'l')  # Focus address bar
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'a')  # Select all
    time.sleep(0.1)
    pyautogui.typewrite(url, interval=0.01)  # Type URL fast
    time.sleep(0.3)
    pyautogui.press('enter')
    
    print(f"[WAIT] Loading page ({wait}s)...")
    time.sleep(wait)


def scroll_page() -> None:
    """Scroll page down to show chart area."""
    count = int(os.getenv('SCROLL_DOWN_COUNT', '7'))
    distance = int(os.getenv('SCROLL_DISTANCE', '-90'))
    delay = float(os.getenv('SCROLL_DELAY', '0.005'))
    post_wait = float(os.getenv('POST_SCROLL_WAIT', '0.3'))
    
    if count > 0:
        print(f"[SCROLL] {count}x (distance: {distance})")
        for _ in range(count):
            pyautogui.scroll(distance)
            time.sleep(delay)
        time.sleep(post_wait)


def capture_screenshot() -> Image.Image:
    """Capture full screen or region screenshot based on env settings."""
    left = os.getenv('CHART_LEFT', '').strip()
    top = os.getenv('CHART_TOP', '').strip()
    width = os.getenv('CHART_WIDTH', '').strip()
    height = os.getenv('CHART_HEIGHT', '').strip()
    
    if left and top and width and height:
        region = (int(left), int(top), int(width), int(height))
        print(f"[SCREENSHOT] Region: {region}")
        return pyautogui.screenshot(region=region)
    else:
        print("[SCREENSHOT] Full screen")
        return pyautogui.screenshot()


def take_screenshot(stock_symbol: str | None = None) -> str | None:
    """
    Main function: find window, search stock, scroll, and capture screenshot.
    Args:
        stock_symbol: Stock symbol to capture. If None, reads from STOCK_SYMBOL env.
    Returns filepath on success, None on failure.
    """
    # Load config
    window_keyword = os.getenv('WINDOW_KEYWORD', 'barchart')
    base_output_dir = os.getenv('OUTPUT_DIR', 'screenshots')
    settle_delay = float(os.getenv('WINDOW_SETTLE_DELAY', '2.0'))
    max_retries = int(os.getenv('MAX_RETRIES', '3'))
    refresh_wait = float(os.getenv('REFRESH_WAIT', '5.0'))
    click_wait = float(os.getenv('CLICK_WAIT', '0.3'))
    
    # Use provided symbol or fall back to env
    if stock_symbol is None:
        stock_symbol = os.getenv('STOCK_SYMBOL', '').strip()
    
    # Create output directory (subfolder per stock if symbol provided)
    if stock_symbol:
        output_dir = os.path.join(base_output_dir, stock_symbol)
    else:
        output_dir = base_output_dir
    os.makedirs(output_dir, exist_ok=True)
    
    # Build filename
    prefix = os.getenv('FILENAME_PREFIX', 'barchart')
    suffix = os.getenv('FILENAME_SUFFIX', 'put_call_ratios')
    symbol_part = f"{stock_symbol}_" if stock_symbol else ""
    filepath = os.path.join(output_dir, f"{prefix}_{symbol_part}{suffix}.png")

    # Step 1: Find and focus browser window
    print(f"\n[>>] Finding '{window_keyword}' window...")
    if not bring_window_to_front(window_keyword):
        print("[ERROR] Window not found. Open Barchart in a browser first.")
        return None

    print(f"[WAIT] Window settle ({settle_delay}s)...")
    time.sleep(settle_delay)

    # Step 2: Search for stock (if symbol provided)
    if stock_symbol:
        search_stock(stock_symbol)

    # Step 3: Click center to ensure focus
    w, h = pyautogui.size()
    pyautogui.click(w // 2, h // 2)
    time.sleep(click_wait)

    # Step 4: Scroll to chart area
    scroll_page()

    # Step 5: Capture with retry on blank
    for attempt in range(1, max_retries + 1):
        print(f"\n[Attempt {attempt}/{max_retries}]")
        screenshot = capture_screenshot()
        
        if is_screenshot_blank(screenshot):
            print("  [WARN] Blank screen detected")
            if attempt < max_retries:
                print(f"  Refreshing ({refresh_wait}s)...")
                pyautogui.press('f5')
                time.sleep(refresh_wait)
                continue
            print("  [WARN] Still blank, saving anyway")
        else:
            print("  [OK] Captured")
        
        screenshot.save(filepath)
        break

    print(f"\n[OK] Saved: {filepath}")
    return filepath


def load_symbols_from_file(filepath: str) -> list[str]:
    """Load stock symbols from a text file (one per line)."""
    try:
        with open(filepath, 'r') as f:
            symbols = [line.strip().upper() for line in f if line.strip()]
        return symbols
    except Exception as e:
        print(f"[ERROR] Failed to read symbols file: {e}")
        return []


def process_batch(symbols: list[str]) -> dict:
    """Process multiple stock symbols and return results summary."""
    results = {'success': [], 'failed': [], 'skipped': []}
    total = len(symbols)
    
    # Get symbols to skip
    skip_symbols = [s.strip().upper() for s in os.getenv('SKIP_SYMBOLS', '').split(',') if s.strip()]
    
    for i, symbol in enumerate(symbols, 1):
        print(f"\n{'='*60}")
        print(f"[{i}/{total}] Processing: {symbol}")
        print('='*60)
        
        if symbol in skip_symbols:
            print(f"[SKIP] {symbol} is in skip list")
            results['skipped'].append(symbol)
            continue
        
        try:
            result = take_screenshot(stock_symbol=symbol)
            if result:
                results['success'].append(symbol)
            else:
                results['failed'].append(symbol)
        except Exception as e:
            print(f"[ERROR] {symbol} failed: {e}")
            results['failed'].append(symbol)
    
    return results


def main():
    """Entry point - supports single symbol or batch mode."""
    # Check for batch mode (symbols file)
    symbols_file = os.getenv('STOCK_SYMBOLS_FILE', '').strip()
    single_symbol = os.getenv('STOCK_SYMBOL', '').strip()
    
    if symbols_file and os.path.exists(symbols_file):
        # Batch mode
        print(f"[BATCH MODE] Loading symbols from: {symbols_file}")
        symbols = load_symbols_from_file(symbols_file)
        
        if not symbols:
            print("[ERROR] No symbols found in file.")
            return
        
        print(f"[INFO] Found {len(symbols)} symbols: {', '.join(symbols)}")
        
        results = process_batch(symbols)
        
        # Print summary
        print(f"\n{'='*60}")
        print("BATCH COMPLETE")
        print('='*60)
        print(f"Success: {len(results['success'])} - {', '.join(results['success']) or 'None'}")
        if results['skipped']:
            print(f"Skipped: {len(results['skipped'])} - {', '.join(results['skipped'])}")
        if results['failed']:
            print(f"Failed:  {len(results['failed'])} - {', '.join(results['failed'])}")
    
    elif single_symbol:
        # Single symbol mode
        result = take_screenshot(stock_symbol=single_symbol)
        if result:
            print(f"\nDone! Screenshot: {result}")
        else:
            print("\nFailed. Ensure Barchart is open in a browser.")
    
    else:
        print("[ERROR] No STOCK_SYMBOL or STOCK_SYMBOLS_FILE configured in .env")


if __name__ == "__main__":
    main()
