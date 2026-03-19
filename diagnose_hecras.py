"""HEC-RAS Preprocessing Diagnostic Script

Run this on the Windows VM with HEC-RAS 6.6 installed.
It discovers all window titles, menu items, and control IDs,
then attempts to open RAS Mapper and logs everything.

Usage:
    python diagnose_hecras.py "C:\path\to\project.prj" "C:\path\to\Ras.exe"

Output: diagnose_results.json (share this file back)
"""

import json
import os
import subprocess
import sys
import time
import traceback

try:
    import win32gui
    import win32con
    import win32api
    import win32process
    import ctypes
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32")
    sys.exit(1)

WM_COMMAND = 0x0111
results = {
    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
    "hecras_exe": "",
    "prj_file": "",
    "steps": [],
    "errors": [],
}


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}")
    results["steps"].append({"time": time.strftime("%H:%M:%S"), "msg": msg})


def get_menu_string(menu_handle, pos):
    buf_size = 256
    buf = ctypes.create_unicode_buffer(buf_size)
    ctypes.windll.user32.GetMenuStringW(menu_handle, pos, buf, buf_size, 0x00000400)
    return buf.value


def find_all_windows():
    """Find all visible windows with titles."""
    windows = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                cls = win32gui.GetClassName(hwnd)
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                windows.append({
                    "hwnd": hwnd,
                    "title": title,
                    "class": cls,
                    "pid": pid,
                })
        return True
    win32gui.EnumWindows(callback, None)
    return windows


def find_window_by_title(pattern):
    """Find first window whose title contains pattern."""
    for w in find_all_windows():
        if pattern.lower() in w["title"].lower():
            return w["hwnd"]
    return None


def get_all_menus(hwnd):
    """Get all menu items from a window."""
    menus = {}
    menu_bar = win32gui.GetMenu(hwnd)
    if not menu_bar:
        return {"error": "No menu bar found"}

    menu_count = win32gui.GetMenuItemCount(menu_bar)
    for i in range(menu_count):
        menu_text = get_menu_string(menu_bar, i).replace("&", "")
        submenu = win32gui.GetSubMenu(menu_bar, i)
        items = []
        if submenu:
            item_count = win32gui.GetMenuItemCount(submenu)
            for j in range(item_count):
                item_text = get_menu_string(submenu, j)
                menu_id = win32gui.GetMenuItemID(submenu, j)
                items.append({"text": item_text, "id": menu_id, "pos": j})

                # Check sub-submenus
                sub_submenu = win32gui.GetSubMenu(submenu, j)
                if sub_submenu:
                    sub_items = []
                    sub_count = win32gui.GetMenuItemCount(sub_submenu)
                    for k in range(sub_count):
                        sub_text = get_menu_string(sub_submenu, k)
                        sub_id = win32gui.GetMenuItemID(sub_submenu, k)
                        sub_items.append({"text": sub_text, "id": sub_id, "pos": k})
                    items[-1]["submenu"] = sub_items

        menus[menu_text] = items
    return menus


def get_child_controls(hwnd):
    """Get all child controls of a window."""
    controls = []
    def callback(child_hwnd, _):
        cls = win32gui.GetClassName(child_hwnd)
        text = win32gui.GetWindowText(child_hwnd)
        ctrl_id = win32gui.GetDlgCtrlID(child_hwnd)
        rect = win32gui.GetWindowRect(child_hwnd)
        controls.append({
            "hwnd": child_hwnd,
            "class": cls,
            "text": text[:100] if text else "",
            "id": ctrl_id,
            "rect": list(rect),
        })
        return True
    try:
        win32gui.EnumChildWindows(hwnd, callback, None)
    except Exception:
        pass
    return controls


def main():
    if len(sys.argv) < 3:
        print("Usage: python diagnose_hecras.py <project.prj> <Ras.exe>")
        print()
        print("Example:")
        print('  python diagnose_hecras.py "C:\\Models\\VA_2D_Test2\\VA_2D_Test2.prj" "C:\\Program Files\\HEC\\HEC-RAS\\6.6\\Ras.exe"')
        sys.exit(1)

    prj_file = sys.argv[1]
    hec_ras_exe = sys.argv[2]

    results["prj_file"] = prj_file
    results["hecras_exe"] = hec_ras_exe

    log(f"PRJ file: {prj_file}")
    log(f"HEC-RAS exe: {hec_ras_exe}")
    log(f"PRJ exists: {os.path.exists(prj_file)}")
    log(f"EXE exists: {os.path.exists(hec_ras_exe)}")

    # Step 1: Launch HEC-RAS
    log("Step 1: Launching HEC-RAS...")
    process = subprocess.Popen([hec_ras_exe, prj_file])
    pid = process.pid
    log(f"Process started (PID: {pid})")

    # Step 2: Wait for HEC-RAS window
    log("Step 2: Waiting for HEC-RAS main window...")
    time.sleep(8)

    # Handle "already running" dialog
    already_running = find_window_by_title("already running")
    if not already_running:
        already_running = find_window_by_title("Another instance")
    if already_running:
        log(f"Found 'already running' dialog — clicking Yes")
        # Try to find Yes button
        def find_button(hwnd, _):
            if win32gui.GetClassName(hwnd) == "Button":
                text = win32gui.GetWindowText(hwnd)
                if "yes" in text.lower() or "ok" in text.lower():
                    win32gui.SendMessage(hwnd, win32con.BM_CLICK, 0, 0)
            return True
        try:
            win32gui.EnumChildWindows(already_running, find_button, None)
        except Exception:
            pass
        time.sleep(3)

    # Find HEC-RAS main window
    hecras_hwnd = None
    for attempt in range(15):
        windows = find_all_windows()
        for w in windows:
            if "HEC-RAS" in w["title"] and w["pid"] == pid:
                if win32gui.GetMenu(w["hwnd"]):
                    hecras_hwnd = w["hwnd"]
                    break
        if hecras_hwnd:
            break
        time.sleep(2)

    if not hecras_hwnd:
        log("ERROR: Could not find HEC-RAS main window!")
        log("All visible windows:")
        for w in find_all_windows():
            log(f"  PID={w['pid']} HWND={w['hwnd']} class={w['class']} title='{w['title']}'")
        results["errors"].append("HEC-RAS window not found")
        save_results()
        process.terminate()
        return

    log(f"Found HEC-RAS window: HWND={hecras_hwnd}, Title='{win32gui.GetWindowText(hecras_hwnd)}'")

    # Step 3: Get all HEC-RAS menus
    log("Step 3: Discovering HEC-RAS menu structure...")
    menus = get_all_menus(hecras_hwnd)
    results["hecras_menus"] = menus

    for menu_name, items in menus.items():
        log(f"  Menu: [{menu_name}]")
        for item in items:
            log(f"    [{item['pos']}] {item['text']:40s} ID={item['id']}")
            if "submenu" in item:
                for sub in item["submenu"]:
                    log(f"      [{sub['pos']}] {sub['text']:38s} ID={sub['id']}")

    # Step 4: Find RAS Mapper menu item
    log("Step 4: Looking for RAS Mapper menu item...")
    mapper_menu_id = None
    mapper_menu_path = None

    for menu_name, items in menus.items():
        for item in items:
            item_text = item["text"].lower()
            if "mapper" in item_text or "ras m" in item_text:
                mapper_menu_id = item["id"]
                mapper_menu_path = f"{menu_name} -> {item['text']}"
                log(f"  FOUND: {mapper_menu_path} (ID={mapper_menu_id})")
            if "submenu" in item:
                for sub in item["submenu"]:
                    sub_text = sub["text"].lower()
                    if "mapper" in sub_text or "ras m" in sub_text:
                        mapper_menu_id = sub["id"]
                        mapper_menu_path = f"{menu_name} -> {item['text']} -> {sub['text']}"
                        log(f"  FOUND: {mapper_menu_path} (ID={mapper_menu_id})")

    if not mapper_menu_id:
        log("ERROR: Could not find RAS Mapper menu item!")
        log("Looking for any GIS-related menu items...")
        for menu_name, items in menus.items():
            for item in items:
                if any(kw in item["text"].lower() for kw in ["gis", "map", "spatial", "view"]):
                    log(f"  Possible: {menu_name} -> {item['text']} (ID={item['id']})")
        results["errors"].append("RAS Mapper menu not found")
        save_results()
        process.terminate()
        return

    results["mapper_menu_id"] = mapper_menu_id
    results["mapper_menu_path"] = mapper_menu_path

    # Step 5: Click RAS Mapper menu item
    log(f"Step 5: Opening RAS Mapper via menu ID {mapper_menu_id}...")
    win32gui.PostMessage(hecras_hwnd, WM_COMMAND, mapper_menu_id, 0)

    # Step 6: Wait for RAS Mapper window
    log("Step 6: Waiting for RAS Mapper window (up to 120s)...")
    mapper_hwnd = None
    start_wait = time.time()

    for attempt in range(60):
        # Check for any new windows
        for pattern in ["RAS Mapper", "Mapper", "Map"]:
            hwnd = find_window_by_title(pattern)
            if hwnd and hwnd != hecras_hwnd:
                mapper_hwnd = hwnd
                break
        if mapper_hwnd:
            break

        # Every 10 seconds, log all windows
        elapsed = int(time.time() - start_wait)
        if elapsed % 10 == 0 and elapsed > 0:
            log(f"  Still waiting ({elapsed}s)... Current windows:")
            for w in find_all_windows():
                if w["pid"] == pid or "ras" in w["title"].lower() or "map" in w["title"].lower():
                    log(f"    PID={w['pid']} HWND={w['hwnd']} title='{w['title']}'")
        time.sleep(2)

    if mapper_hwnd:
        elapsed = int(time.time() - start_wait)
        log(f"RAS Mapper found! HWND={mapper_hwnd}, Title='{win32gui.GetWindowText(mapper_hwnd)}', took {elapsed}s")

        # Step 7: Discover RAS Mapper controls
        log("Step 7: Discovering RAS Mapper menus and controls...")
        mapper_menus = get_all_menus(mapper_hwnd)
        results["mapper_menus"] = mapper_menus

        if isinstance(mapper_menus, dict) and "error" not in mapper_menus:
            for menu_name, items in mapper_menus.items():
                log(f"  Mapper Menu: [{menu_name}]")
                for item in items:
                    log(f"    [{item['pos']}] {item['text']:40s} ID={item['id']}")
        else:
            log(f"  No traditional menu bar: {mapper_menus}")

        # Get controls
        controls = get_child_controls(mapper_hwnd)
        results["mapper_controls"] = [
            {"class": c["class"], "text": c["text"], "id": c["id"]}
            for c in controls
        ]
        log(f"  Found {len(controls)} child controls")
        for c in controls:
            if c["text"] or c["class"] in ["SysTreeView32", "TreeView", "ToolbarWindow32", "Button"]:
                log(f"    class={c['class']:30s} text='{c['text'][:50]}' id={c['id']}")

        # Close RAS Mapper
        log("Closing RAS Mapper...")
        win32gui.PostMessage(mapper_hwnd, win32con.WM_CLOSE, 0, 0)
        time.sleep(3)
    else:
        elapsed = int(time.time() - start_wait)
        log(f"ERROR: RAS Mapper did not appear within {elapsed}s!")
        log("All windows at timeout:")
        for w in find_all_windows():
            log(f"  PID={w['pid']} HWND={w['hwnd']} class={w['class']} title='{w['title']}'")
        results["errors"].append(f"RAS Mapper window not found after {elapsed}s")

    # Close HEC-RAS
    log("Closing HEC-RAS...")
    win32gui.PostMessage(hecras_hwnd, win32con.WM_CLOSE, 0, 0)
    time.sleep(3)

    # Handle save dialog
    save_dlg = find_window_by_title("Save")
    if not save_dlg:
        save_dlg = find_window_by_title("HEC-RAS")
    if save_dlg and save_dlg != hecras_hwnd:
        def click_yes(hwnd, _):
            text = win32gui.GetWindowText(hwnd)
            if "yes" in text.lower() or "no" in text.lower():
                win32gui.SendMessage(hwnd, win32con.BM_CLICK, 0, 0)
            return True
        try:
            win32gui.EnumChildWindows(save_dlg, click_yes, None)
        except Exception:
            pass

    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        process.terminate()

    save_results()
    log("Done! Results saved to diagnose_results.json")


def save_results():
    output_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "diagnose_results.json")
    with open(output_file, "w") as f:
        json.dump(results, f, indent=2, default=str)
    print(f"\nResults saved to: {output_file}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        traceback.print_exc()
        results["errors"].append(str(e))
        save_results()
