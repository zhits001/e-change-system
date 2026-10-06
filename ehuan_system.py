# -*- coding: utf-8 -*-
"""
e换系统 —— 换系统前的迁移清单助手
================================
功能：
  1. 扫描磁盘，定位需要搬走的个人文件（用户目录 / 聊天记录 / 浏览器数据 /
     开发环境 / 邮件 / 字体 / 其他磁盘），列出目录、大小、文件数与迁移建议。
  2. 扫描已安装软件（注册表 + 商店应用），列出安装目录，并给出在
     新 Windows / macOS / Linux 上的兼容情况与迁移建议。
  3. 一键生成 HTML《换系统迁移清单》报告，可导出 CSV。

仅使用 Python 标准库，无第三方依赖。
用法：
  python ehuan_system.py            # 打开图形界面
  python ehuan_system.py --report   # 无界面直接生成迁移报告（保存到桌面）
  python ehuan_system.py --selftest # 自测扫描逻辑并打印结果
"""

import os
import sys
import csv
import heapq
import ctypes
import queue
import threading
import subprocess
import traceback
import re
import getpass
import platform
import datetime
import webbrowser

APP_NAME = "e换系统"
APP_VERSION = "1.0"

MB = 1024 * 1024
GB = 1024 * MB


def _utf8_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ----------------------------------------------------------------------------
# Windows 基础能力：已知文件夹、磁盘枚举
# ----------------------------------------------------------------------------

is_windows = sys.platform.startswith("win")

# 已知文件夹 GUID（桌面/文档/下载/图片/音乐/视频）
_KNOWN_GUIDS = {
    "desktop": "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}",
    "documents": "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}",
    "downloads": "{374DE290-123F-4565-9164-39C4925E467B}",
    "pictures": "{33E28130-4E1E-4676-835A-98395C95C3BC3BB}",
    "music": "{4BD8D571-6D19-48D3-BE97-422220080E43}",
    "videos": "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}",
}


def known_folder(key):
    """通过 SHGetKnownFolderPath 取真实路径（兼容 OneDrive 重定向），失败回退。"""
    if not is_windows:
        return None
    guid = _KNOWN_GUIDS.get(key)
    home = os.path.expanduser("~")
    fallback = {
        "desktop": os.path.join(home, "Desktop"),
        "documents": os.path.join(home, "Documents"),
        "downloads": os.path.join(home, "Downloads"),
        "pictures": os.path.join(home, "Pictures"),
        "music": os.path.join(home, "Music"),
        "videos": os.path.join(home, "Videos"),
    }.get(key)
    try:
        import uuid
        from ctypes import wintypes

        class _GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", wintypes.DWORD),
                ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        u = uuid.UUID(guid)
        g = _GUID()
        g.Data1 = u.time_low
        g.Data2 = u.time_mid
        g.Data3 = u.time_version
        for i, b in enumerate(u.bytes[8:16]):
            g.Data4[i] = b
        ptr = ctypes.c_wchar_p()
        res = ctypes.windll.shell32.SHGetKnownFolderPath(
            ctypes.byref(g), 0, None, ctypes.byref(ptr)
        )
        if res == 0 and ptr.value:
            path = ptr.value
            ctypes.windll.ole32.CoTaskMemFree(ptr)
            if os.path.exists(path):
                return path
    except Exception:
        pass
    return fallback if fallback and os.path.exists(fallback) else None


def list_fixed_drives():
    """返回 [(盘符, 是否系统盘, 是否可移动)]。"""
    drives = []
    if not is_windows:
        return drives
    try:
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        for i in range(26):
            if bitmask & (1 << i):
                letter = chr(ord("A") + i) + ":\\"
                dt = ctypes.windll.kernel32.GetDriveTypeW(ctypes.c_wchar_p(letter))
                if dt in (3, 4):  # 3=固定磁盘 4=网络磁盘
                    drives.append((letter, letter[:2] == os.environ.get("SystemDrive", "C:").upper()))
    except Exception:
        pass
    return drives


# ----------------------------------------------------------------------------
# 目录统计：迭代遍历，跳过符号链接/联接，可取消，顺带收集大文件
# ----------------------------------------------------------------------------

def dir_stats(path, stop_event=None, top_files=None, top_n=30, min_size=200 * MB):
    """返回 (总大小, 文件数, 目录数)。top_files: [(size, path)] 堆。"""
    total = 0
    files = 0
    dirs = 0
    if not os.path.isdir(path):
        return 0, 0, 0
    stack = [path]
    while stack:
        if stop_event is not None and stop_event.is_set():
            break
        p = stack.pop()
        try:
            it = os.scandir(p)
        except OSError:
            continue
        with it:
            try:
                entries = list(it)
            except OSError:
                continue
        for e in entries:
            try:
                if e.is_symlink():
                    continue  # 跳过符号链接/联接，防止循环（如 OneDrive、应用数据回环）
                if e.is_dir(follow_symlinks=False):
                    dirs += 1
                    stack.append(e.path)
                elif e.is_file(follow_symlinks=False):
                    s = e.stat(follow_symlinks=False).st_size
                    total += s
                    files += 1
                    if top_files is not None and s >= min_size:
                        item = (s, e.path)
                        if len(top_files) < top_n:
                            heapq.heappush(top_files, item)
                        elif s > top_files[0][0]:
                            heapq.heapreplace(top_files, item)
            except OSError:
                continue
    return total, files, dirs


def fmt_size(n):
    if n is None:
        return "—"
    if n >= GB:
        return "%.2f GB" % (n / GB)
    if n >= MB:
        return "%.1f MB" % (n / MB)
    if n >= 1024:
        return "%.1f KB" % (n / 1024)
    return "%d B" % n


# ----------------------------------------------------------------------------
# 个人文件扫描目标定义
# ----------------------------------------------------------------------------

def build_scan_targets(include_user=True, include_chat=True, include_browser=True,
                       include_dev=True, include_office=True):
    """返回 [(类别, 路径, 迁移建议)]，仅保留存在的路径。"""
    home = os.path.expanduser("~")
    appdata = os.environ.get("APPDATA", os.path.join(home, "AppData", "Roaming"))
    local = os.environ.get("LOCALAPPDATA", os.path.join(home, "AppData", "Local"))
    docs = known_folder("documents") or os.path.join(home, "Documents")

    targets = []

    def add(cat, path, tip, must=False):
        if path and (must or os.path.exists(path)):
            targets.append((cat, os.path.normpath(path), tip))

    if include_user:
        add("用户基础目录", known_folder("desktop"), "整个文件夹搬走，桌面常散落重要文件")
        add("用户基础目录", docs, "文档目录：办公文件、微信/QQ 数据默认都在这里")
        add("用户基础目录", known_folder("downloads"), "下载目录：安装包可不留，个人文件需挑出")
        add("用户基础目录", known_folder("pictures"), "图片/截图")
        add("用户基础目录", known_folder("videos"), "视频")
        add("用户基础目录", known_folder("music"), "音乐")

    if include_chat:
        add("聊天与通讯", os.path.join(docs, "WeChat Files"),
            "微信 3.x 聊天记录与接收文件；换机建议用微信自带“聊天记录备份与迁移”")
        add("聊天与通讯", os.path.join(docs, "xwechat_files"),
            "微信 4.x 聊天记录与接收文件")
        add("聊天与通讯", os.path.join(docs, "Tencent Files"), "QQ 聊天记录与接收文件")
        add("聊天与通讯", os.path.join(appdata, "Telegram Desktop", "tdata"),
            "Telegram 会话数据，整目录拷贝可保会话")
        add("聊天与通讯", os.path.join(appdata, "DingTalk"), "钉钉本地数据（聊天云端有存）")

    if include_browser:
        add("浏览器数据", os.path.join(local, "Google", "Chrome", "User Data"),
            "Chrome 书签/密码/扩展；建议登录账号同步，或先导出书签和密码")
        add("浏览器数据", os.path.join(local, "Microsoft", "Edge", "User Data"),
            "Edge 书签/密码；微软账号可同步")
        add("浏览器数据", os.path.join(appdata, "Mozilla", "Firefox", "Profiles"),
            "Firefox 配置；账号同步或用其导出功能")
        add("浏览器数据", os.path.join(local, "Tencent", "QQBrowser", "User Data"), "QQ 浏览器数据")
        add("浏览器数据", os.path.join(appdata, "360se6", "User Data"), "360 浏览器数据")

    if include_dev:
        add("开发环境", os.path.join(home, ".ssh"), "SSH 私钥/已知主机，务必备份且勿泄露")
        add("开发环境", os.path.join(home, ".gitconfig"), "Git 全局配置")
        add("开发环境", os.path.join(home, ".gnupg"), "GPG 密钥")
        add("开发环境", os.path.join(home, ".aws"), "AWS 凭证与配置")
        add("开发环境", os.path.join(home, ".kube"), "Kubernetes 配置")
        add("开发环境", os.path.join(home, ".m2"), "Maven 本地仓库与 settings.xml")
        add("开发环境", os.path.join(home, ".gradle"), "Gradle 缓存与配置")
        add("开发环境", os.path.join(home, ".npmrc"), "npm 配置与令牌")
        add("开发环境", os.path.join(appdata, "Code", "User"),
            "VS Code 配置；建议开启 Settings Sync")
        add("开发环境", os.path.join(appdata, "JetBrains"), "JetBrains IDE 配置；账号可同步")
        add("开发环境", os.path.join(appdata, "Sublime Text"), "Sublime 配置")
        add("开发环境", os.path.join(home, ".docker"), "Docker 配置")

    if include_office:
        add("邮件与办公", os.path.join(docs, "Outlook Files"), "Outlook PST 邮箱数据文件")
        add("邮件与办公", os.path.join(local, "Microsoft", "Outlook"), "Outlook OST 缓存（服务器有原件）")
        add("邮件与办公", os.path.join(appdata, "Microsoft", "Signatures"), "Outlook 签名")
        add("邮件与办公", os.path.join(home, "Documents", "NetSarang"), "Xshell/Ftp 会话配置")

    add("其他个人数据", os.path.join(local, "Microsoft", "Windows", "Fonts"), "自己安装的字体")

    return targets


SYSTEM_SKIP_DIRS = {
    "windows", "program files", "program files (x86)", "programdata",
    "$recycle.bin", "system volume information", "$windows.~bt", "$windows.~ws",
    "$getcurrent", "recovery", "perflogs", "msocache", "intel", "amd", "dell", "hp",
}


def scan_deep_drives(stop_event=None, progress=None):
    """扫描非系统固定磁盘：列出一级目录及其二级目录大小。"""
    results = []  # (类别, 路径, 大小, 文件数, 建议)
    sys_drive = os.environ.get("SystemDrive", "C:").upper()
    for letter, is_sys, _ in list_fixed_drives():
        if is_sys or letter[:2].upper() == sys_drive:
            continue
        try:
            entries = [e for e in os.scandir(letter)]
        except OSError:
            continue
        d1_list = []
        files_root = 0
        for e in entries:
            try:
                if e.is_symlink():
                    continue
                if e.is_dir(follow_symlinks=False):
                    if e.name.lower() in SYSTEM_SKIP_DIRS:
                        continue
                    d1_list.append(e.path)
                elif e.is_file(follow_symlinks=False):
                    files_root += e.stat(follow_symlinks=False).st_size
            except OSError:
                continue
        total_d1 = len(d1_list)
        for idx, d1 in enumerate(d1_list):
            if stop_event is not None and stop_event.is_set():
                return results
            if progress:
                progress(idx, total_d1, d1)
            # 一级目录：文件直接累计，二级目录单独统计
            d1_total, d1_files = 0, 0
            children = []  # (size, files, path)
            try:
                sub = list(os.scandir(d1))
            except OSError:
                sub = []
            for e in sub:
                if stop_event is not None and stop_event.is_set():
                    return results
                try:
                    if e.is_symlink():
                        continue
                    if e.is_dir(follow_symlinks=False):
                        s, f, _ = dir_stats(e.path, stop_event)
                        children.append((s, f, e.path))
                        d1_total += s
                        d1_files += f
                    elif e.is_file(follow_symlinks=False):
                        s = e.stat(follow_symlinks=False).st_size
                        d1_total += s
                        d1_files += 1
                except OSError:
                    continue
            if d1_total > 50 * MB:
                results.append(("其他磁盘-" + letter[:1], d1, d1_total, d1_files,
                                "数据盘目录，按需搬移"))
                children.sort(key=lambda x: -x[0])
                if d1_total > 2 * GB:
                    shown = 0
                    for s, f, p in children:
                        if s < 500 * MB or shown >= 8:
                            break
                        results.append(("其他磁盘-" + letter[:1], p, s, f, ""))
                        shown += 1
    return results


# ----------------------------------------------------------------------------
# 已安装软件扫描（注册表 + 商店应用）
# ----------------------------------------------------------------------------

def scan_installed_software(include_uwp=True):
    """返回 [{name, version, publisher, location, size_mb, kind}]"""
    apps = {}
    seen = set()

    def read_tree(root, access, kind):
        try:
            key = winreg.OpenKey(root, r"Software\Microsoft\Windows\CurrentVersion\Uninstall",
                                 0, access | winreg.KEY_READ)
        except OSError:
            return
        with key:
            i = 0
            while True:
                try:
                    sub_name = winreg.EnumKey(key, i)
                except OSError:
                    break
                i += 1
                try:
                    with winreg.OpenKey(key, sub_name, 0, winreg.KEY_READ) as sk:
                        def gv(name):
                            try:
                                v, _ = winreg.QueryValueEx(sk, name)
                                return v
                            except OSError:
                                return None

                        if gv("SystemComponent") == 1:
                            continue
                        name = gv("DisplayName")
                        if not name:
                            continue
                        dedup = (name.lower(), str(gv("DisplayVersion") or ""))
                        if dedup in seen:
                            continue
                        seen.add(dedup)
                        size = gv("EstimatedSize")  # KB
                        loc = gv("InstallLocation") or ""
                        if not loc:
                            us = gv("UninstallString") or ""
                            m = re.search(r'"?([A-Za-z]:\\[^"]+?)(?:\\[^\\"]+)?\.exe', us)
                            if m:
                                loc = os.path.dirname(m.group(1))
                        apps[dedup] = {
                            "name": str(name),
                            "version": str(gv("DisplayVersion") or ""),
                            "publisher": str(gv("Publisher") or ""),
                            "location": str(loc),
                            "size_mb": round(size / 1024, 1) if size else None,
                            "kind": kind,
                        }
                except OSError:
                    continue
        return

    import winreg
    read_tree(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY, "系统")
    read_tree(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY, "系统32位")
    read_tree(winreg.HKEY_CURRENT_USER, 0, "当前用户")

    if include_uwp:
        for app in _scan_uwp():
            dedup = (app["name"].lower(), app["version"])
            if dedup not in seen:
                seen.add(dedup)
                apps[dedup] = app

    return sorted(apps.values(), key=lambda a: a["name"].lower())


def _scan_uwp():
    cmd = (
        "Get-AppxPackage | Where-Object { -not $_.IsFramework -and $_.NonRemovable -ne $true } | "
        "Select-Object Name, Version, InstallLocation | ConvertTo-Json -Compress"
    )
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd],
            capture_output=True, text=True, timeout=40,
        )
        if out.returncode != 0 or not out.stdout.strip():
            return []
        import json
        data = json.loads(out.stdout)
        if isinstance(data, dict):
            data = [data]
        result = []
        for p in data:
            name = p.get("Name") or ""
            if not name or name.startswith("Microsoft.NET") or name.startswith("Microsoft.VCLibs"):
                continue
            result.append({
                "name": name,
                "version": str(p.get("Version") or ""),
                "publisher": "Microsoft Store",
                "location": p.get("InstallLocation") or "",
                "size_mb": None,
                "kind": "商店应用",
            })
        return result
    except Exception:
        return []


# ----------------------------------------------------------------------------
# 软件跨系统兼容性知识库
# 符号：✅ 官方支持  ❌ 无官方版（需找替代）  ➖ 情况特殊/未知
# ----------------------------------------------------------------------------

KB = [
    # (匹配子串, 显示分类名, 新Windows, macOS, Linux, 迁移建议)
    (["google chrome"], "浏览器", "✅", "✅", "✅", "登录 Google 账号可同步书签/密码/扩展；不登录请先导出书签"),
    (["microsoft edge"], "浏览器", "✅", "✅", "❌", "微软账号同步；Linux 官方版已停止支持"),
    (["firefox", "火狐"], "浏览器", "✅", "✅", "✅", "账号同步，或导出书签/密码"),
    (["opera"], "浏览器", "✅", "✅", "✅", "账号同步"),
    (["brave"], "浏览器", "✅", "✅", "✅", "账号同步"),
    (["360安全浏览器", "360极速浏览器", "360se"], "浏览器", "✅", "❌", "❌", "书签需在浏览器内导出为 HTML"),
    (["qq浏览器"], "浏览器", "✅", "✅", "❌", "书签可账号同步"),
    (["visual studio code", "vs code", "vscode"], "开发", "✅", "✅", "✅", "开启 Settings Sync（GitHub/微软账号）一步迁移"),
    (["android studio"], "开发", "✅", "✅", "✅", "项目代码目录需单独备份"),
    (["pycharm", "intellij idea", "webstorm", "clion", "goland", "rider", "datagrip", "phpstorm", "jetbrains"],
     "开发", "✅", "✅", "✅", "JetBrains 账号登录可同步设置；重装前正常退出登录"),
    (["eclipse"], "开发", "✅", "✅", "✅", "工作区在用户目录，需备份"),
    (["visual studio 2", "visual studio professional", "visual studio enterprise", "visual studio community"],
     "开发", "✅", "❌", "❌", "VS for Mac 已停更；登录微软账号保留授权"),
    (["git"], "开发", "✅", "✅", "✅", "配置在 ~/.gitconfig，密钥在 ~/.ssh，务必带走"),
    (["github desktop"], "开发", "✅", "✅", "➖", "登录账号即可"),
    (["sourcetree"], "开发", "✅", "✅", "❌", ""),
    (["tortoisegit", "tortoisesvn"], "开发", "✅", "❌", "❌", "macOS/Linux 用命令行或 Fork 等替代"),
    (["python"], "开发", "✅", "✅", "✅", "重装后用 pip freeze 清单恢复依赖"),
    (["node.js", "nodejs"], "开发", "✅", "✅", "✅", "全局包用 npm ls -g 导出清单"),
    (["java", "jdk", "jre", "openjdk", "zoralli"], "开发", "✅", "✅", "✅", "注意记录 JAVA_HOME 等环境变量"),
    (["go programming", "golang"], "开发", "✅", "✅", "✅", ""),
    (["rust", "cargo"], "开发", "✅", "✅", "✅", ""),
    (["anaconda", "miniconda", "conda"], "开发", "✅", "✅", "✅", "conda env export 导出环境"),
    (["docker desktop"], "开发", "✅", "✅", "✅", "镜像推送到仓库或 docker save 导出"),
    (["vmware"], "虚拟化", "✅", "✅", "✅", "mac 用 Fusion、Linux 用 Workstation；.vmdk 虚拟机文件单独备份"),
    (["virtualbox"], "虚拟化", "✅", "✅", "✅", "虚拟机文件单独备份"),
    (["wechat", "weixin", "微信"], "通讯", "✅", "✅", "✅", "聊天记录用微信“聊天记录备份与迁移”；新版已有 Linux 原生版"),
    (["qq", "腾讯qq"], "通讯", "✅", "✅", "✅", "NTQQ 有官方 Linux 版；聊天记录可迁移"),
    (["tim"], "通讯", "✅", "➖", "➖", ""),
    (["dingtalk", "钉钉"], "通讯", "✅", "✅", "✅", "聊天记录云端保存"),
    (["feishu", "lark", "飞书"], "通讯", "✅", "✅", "✅", "数据在云端"),
    (["wemeet", "腾讯会议"], "通讯", "✅", "✅", "✅", ""),
    (["zoom"], "通讯", "✅", "✅", "✅", ""),
    (["telegram"], "通讯", "✅", "✅", "✅", "会话数据在 %APPDATA%\\Telegram Desktop"),
    (["discord"], "通讯", "✅", "✅", "✅", ""),
    (["foxmail"], "通讯", "✅", "➖", "❌", "邮件数据在本地，整目录备份"),
    (["thunder", "迅雷"], "下载", "✅", "✅", "❌", ""),
    (["baidunetdisk", "百度网盘"], "网盘", "✅", "✅", "➖", "文件在云端，客户端重装即可"),
    (["onedrive"], "网盘", "✅", "➖", "❌", "换机前确认同步完成（托盘图标打勾）"),
    (["nutstore", "坚果云"], "网盘", "✅", "✅", "➖", ""),
    (["wps", "金山办公"], "办公", "✅", "✅", "✅", "WPS 账号同步配置与文档"),
    (["microsoft 365", "office 1", "microsoft office", "office 3", "office专业增强版", "office 家庭"],
     "办公", "✅", "✅", "❌", "记录微软账号/密钥；Linux 用网页版；mac 装 Office for Mac"),
    (["outlook"], "办公", "✅", "✅", "➖", "PST/OST 数据文件在文档目录，需单独备份"),
    (["adobe", "photoshop", "premiere", "after effects", "illustrator", "lightroom", "acrobat"],
     "创意", "✅", "✅", "❌", "换机前在旧机退出设备授权（Adobe 限 2 台）；保存账号"),
    (["typora"], "效率", "✅", "✅", "✅", "许可证通用"),
    (["obsidian"], "效率", "✅", "✅", "✅", "仓库即文件夹，整目录搬走"),
    (["notion"], "效率", "✅", "✅", "✅", "云端数据，登录即可"),
    (["notepad++"], "效率", "✅", "❌", "➖", "配置在 %APPDATA%\\Notepad++；mac/Linux 可用 VS Code 替代"),
    (["sublime"], "效率", "✅", "✅", "✅", "许可证通用"),
    (["utools"], "效率", "✅", "✅", "✅", "账号同步插件"),
    (["everything"], "工具", "✅", "❌", "❌", "仅 Windows；mac 用 Spotlight、Linux 用 locate 替代"),
    (["listary"], "工具", "✅", "❌", "❌", "仅 Windows"),
    (["7-zip", "7zip"], "工具", "✅", "➖", "✅", "开源；mac 可用 Keka，Linux 用 p7zip"),
    (["winrar"], "工具", "✅", "✅", "➖", "商业软件，保留许可证"),
    (["bandizip"], "工具", "✅", "✅", "✅", ""),
    (["360压缩", "360zip"], "工具", "✅", "❌", "❌", "仅 Windows；mac/Linux 可用 Keka/7-Zip 替代"),
    (["potplayer"], "影音", "✅", "❌", "❌", "mac 用 IINA、Linux 用 MPV/VLC 替代"),
    (["vlc"], "影音", "✅", "✅", "✅", ""),
    (["网易云音乐"], "影音", "✅", "✅", "✅", "歌曲缓存无需迁移"),
    (["qq音乐"], "影音", "✅", "✅", "❌", ""),
    (["酷狗", "kugou"], "影音", "✅", "✅", "❌", "本地下载歌曲在缓存目录，需要的话自行挑出"),
    (["steam"], "游戏", "✅", "✅", "✅", "steamapps 文件夹整体搬移可免重新下载；账号登录"),
    (["epic games"], "游戏", "✅", "✅", "❌", ""),
    (["uplay", "ubisoft connect"], "游戏", "✅", "➖", "❌", ""),
    (["网易uu", "uu加速器", "雷神加速器", "加速器"], "游戏", "✅", "❌", "❌", "账号通用，客户端需重装"),
    (["obs studio", "obs"], "工具", "✅", "✅", "✅", "场景配置在 %APPDATA%\\obs-studio"),
    (["teamviewer", "anydesk", "todesk", "向日葵"], "远程", "✅", "✅", "✅", "记录账号与无人值守密码"),
    (["搜狗输入法"], "输入法", "✅", "✅", "➖", "词库可账号同步；Linux 配 fcitx 使用"),
    (["百度输入法", "qq输入法", "讯飞输入法"], "输入法", "✅", "➖", "❌", "mac/Linux 用系统自带输入法"),
    (["360安全卫士", "360杀毒", "360安全", "360软件管家"], "安全", "✅", "❌", "❌", "建议新系统不重装；无数据迁移价值"),
    (["火绒"], "安全", "✅", "❌", "❌", "仅 Windows"),
    (["nvidia", "geforce", "amd software", "adrenalin", "intel 显卡", "驱动"],
     "驱动", "✅", "❌", "❌", "驱动无需备份，新系统装最新版即可"),
    (["autocad", "revit", "3ds max"], "设计", "✅", "✅", "❌", "AutoCAD 有 mac 版；许可证需在账号中转移"),
    (["solidworks", "ug nx", "creo", "catia"], "设计", "✅", "❌", "❌", "仅 Windows；许可需先转移再重装"),
    (["xmind"], "效率", "✅", "✅", "➖", "文件即 .xmind，搬文件即可"),
    (["postman"], "开发", "✅", "✅", "✅", "账号同步集合"),
    (["xshell", "xftp", "netsarang"], "开发", "✅", "❌", "❌", "会话配置在 文档\\NetSarang；家庭学校版免费"),
    (["finalshell"], "开发", "✅", "➖", "✅", "连接配置在 ~/.finalshell"),
    (["power toys", "powertoys"], "工具", "✅", "❌", "❌", "微软官方，仅 Windows"),
    (["itunes", "apple application", "apple mobile"], "工具", "✅", "✅", "❌",
     "iPhone 备份在 %APPDATA%\\Apple Computer\\MobileSync，需要则备份"),
    (["microsoft .net", ".net framework", ".net sdk", ".net core"], "运行库", "✅", "✅", "✅",
     ".NET(5+) 跨平台；.NET Framework 仅 Windows，新系统自带"),
    (["vmware horizon", "citrix workspace"], "远程", "✅", "✅", "➖", "记录服务器地址与账号"),
]


def _kb_compiled():
    """把子串模式编译成正则：ASCII 开头/结尾加词边界，避免 todesk 误中 autodesk、
    git 误中 github 等；中文模式保持纯子串（QQ浏览器 也要能命中 qq）。"""
    if getattr(_kb_compiled, "_cache", None) is None:
        compiled = []
        for pats, cat, w, m, l, tip in KB:
            regs = []
            for p in pats:
                pre = r"(?<![a-z0-9])" if p[:1].isascii() and p[0].isalnum() else ""
                suf = r"(?![a-z0-9])" if p[-1:].isascii() and p[-1].isalnum() else ""
                regs.append(re.compile(pre + re.escape(p) + suf))
            compiled.append((regs, cat, w, m, l, tip))
        _kb_compiled._cache = compiled
    return _kb_compiled._cache


def match_kb(name, publisher):
    s = (name + " " + publisher).lower()
    for regs, cat, w, m, l, tip in _kb_compiled():
        for r in regs:
            if r.search(s):
                return cat, w, m, l, tip
    return "未收录", "✅", "➖", "➖", "未收录软件：换系统后需重装，请自行记录许可证密钥与配置目录"


# ----------------------------------------------------------------------------
# 开发环境包清单（尽力而为）
# ----------------------------------------------------------------------------

def dev_package_lists():
    out = []

    def run(title, cmd):
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=30, shell=True)
            if r.returncode == 0 and r.stdout.strip():
                out.append((title, r.stdout.strip()))
        except Exception:
            pass

    run("pip 已安装的包（pip list）", ["pip", "list"])
    run("npm 全局包（npm ls -g --depth=0）", ["npm", "ls", "-g", "--depth=0"])
    return out


# ----------------------------------------------------------------------------
# HTML 迁移报告
# ----------------------------------------------------------------------------

CHECKLIST = [
    "浏览器：导出书签为 HTML；密码用账号同步或导出（Chrome: 设置→自动填充→密码）",
    "微信：设置→通用→聊天记录备份与迁移（或迁移到另一台设备）",
    "QQ：消息记录导出或直接搬 Tencent Files 目录",
    "许可证类软件（Office/Adobe/Autodesk/Xshell 等）：记录账号与密钥；Adobe 需先退出设备授权",
    "Outlook/foxmail 邮件：确认 PST 数据文件已备份，或改为 IMAP 自动同步",
    "SSH/GPG 密钥：备份 ~/.ssh 与 ~/.gnupg，私钥切勿上传网盘明文",
    "开发者：pip freeze / npm ls -g / conda env export 导出依赖清单（本报告已尝试自动生成）",
    "BitLocker 加密盘：先到微软账户保存恢复密钥，否则换系统后打不开数据盘",
    "Windows 激活：数字许可证绑定微软账号后重装自动激活",
    "OneDrive/坚果云/百度网盘：确认同步全部完成（托盘图标为绿色勾）",
    "字体：自己安装的字体目录（本工具已扫描）需拷走",
    "打印机/扫描仪：驱动不用备份，记下型号即可",
    "游戏：Steam 库文件夹整体搬移可免重新下载；记录各平台账号",
    "最后：清空回收站与 C 盘临时文件前，再检查一遍本清单",
]


def build_report(items, software, top_files, dev_lists, total_size, cancelled=False):
    """items: [(类别,路径,大小,文件数,建议,勾选)]  software: [dict...]"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    host = os.environ.get("COMPUTERNAME", platform.node())
    user = getpass.getuser()
    try:
        sysver = " ".join(platform.win32_ver()[:2]) or platform.platform()
    except Exception:
        sysver = platform.platform()

    def esc(s):
        return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

    rows_f = []
    for cat, path, size, files, tip, checked in items:
        color = "#8a8a8a" if not checked else "#111"
        mark = "" if checked else "（未勾选）"
        rows_f.append(
            f"<tr><td>{esc(cat)}</td><td style='color:{color}'>{esc(path)}{mark}</td>"
            f"<td style='text-align:right'>{fmt_size(size)}</td>"
            f"<td style='text-align:right'>{files:,}</td><td>{esc(tip)}</td></tr>"
        )

    rows_s = []
    for a in software:
        cat, w, m, l, tip = match_kb(a["name"], a["publisher"])
        loc = a["location"] or "—"
        rows_s.append(
            f"<tr><td>{esc(a['name'])}</td><td>{esc(a['version'])}</td>"
            f"<td>{esc(a['publisher'][:24] or '—')}</td>"
            f"<td style='text-align:right'>{a['size_mb'] and ('%.0f MB' % a['size_mb']) or '—'}</td>"
            f"<td>{esc(loc)}</td><td style='text-align:center'>{w}</td>"
            f"<td style='text-align:center'>{m}</td><td style='text-align:center'>{l}</td>"
            f"<td>{esc(tip)}</td></tr>"
        )

    rows_big = "".join(
        f"<tr><td>{esc(p)}</td><td style='text-align:right'>{fmt_size(s)}</td></tr>"
        for s, p in sorted(top_files or [], reverse=True)
    )
    rows_chk = "".join(f"<li><label><input type='checkbox'> {esc(c)}</label></li>" for c in CHECKLIST)
    dev_html = ""
    if dev_lists:
        parts = []
        for title, content in dev_lists:
            code = esc(content).replace("\n", "<br>")
            parts.append(f"<h4>{esc(title)}</h4><div class='code'>{code}</div>")
        dev_html = ("<div class='card'><h2>5. 开发环境依赖清单（自动生成）</h2>"
                    + "".join(parts) + "</div>")
    else:
        dev_html = ""

    cancel_note = "<p class='warn'>⚠ 本次扫描被提前停止，以下数据可能不完整。</p>" if cancelled else ""

    html = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>换系统迁移清单 - {esc(host)}</title>
<style>
 body{{font-family:"Microsoft YaHei UI",system-ui,sans-serif;margin:0;background:#f2f4f8;color:#111}}
 .head{{background:linear-gradient(120deg,#1e66f5,#7b2ff7);color:#fff;padding:28px 36px}}
 .head h1{{margin:0 0 6px;font-size:26px}} .head p{{margin:0;opacity:.9}}
 .wrap{{max-width:1180px;margin:0 auto;padding:20px}}
 .card{{background:#fff;border-radius:12px;padding:18px 22px;margin:16px 0;box-shadow:0 1px 6px rgba(0,0,0,.08)}}
 h2{{font-size:18px;border-left:4px solid #1e66f5;padding-left:10px;margin:4px 0 12px}}
 table{{border-collapse:collapse;width:100%;font-size:13px}}
 th,td{{border:1px solid #e3e6ec;padding:6px 9px;text-align:left;vertical-align:top}}
 th{{background:#f0f3fa;position:sticky;top:0}}
 tr:nth-child(even) td{{background:#fafbfd}}
 .code{{background:#0d1117;color:#9fe8a7;border-radius:8px;padding:12px;font-family:Consolas,monospace;font-size:12px;max-height:260px;overflow:auto}}
 .warn{{color:#c7000b;background:#fff3f3;border:1px solid #ffd9d9;border-radius:8px;padding:8px 12px}}
 ul{{padding-left:0;list-style:none}} li{{margin:6px 0}}
 .muted{{color:#777;font-size:12px}}
</style></head><body>
<div class="head"><h1>换系统迁移清单</h1>
<p>生成于 {now} · 主机 {esc(host)} · 用户 {esc(user)} · 系统 {esc(sysver)}</p></div>
<div class="wrap">
{cancel_note}
<div class="card"><h2>1. 概览</h2>
<p>个人数据 <b>{len(items)}</b> 处，合计约 <b>{fmt_size(total_size)}</b>；
已安装软件 <b>{len(software)}</b> 个（其中商店应用已按需包含）。
勾选项即建议搬移的内容；未勾选项仅供参考。</p></div>

<div class="card"><h2>2. 个人文件清单</h2>
<table><tr><th>类别</th><th>位置</th><th>大小</th><th>文件数</th><th>迁移建议</th></tr>
{''.join(rows_f)}</table></div>

<div class="card"><h2>3. 大文件 TOP 榜（≥200MB，便于挑着搬）</h2>
{'<table><tr><th>文件</th><th>大小</th></tr>' + rows_big + '</table>' if rows_big else '<p class="muted">未发现超过 200MB 的单文件（或扫描已停止）。</p>'}
</div>

<div class="card"><h2>4. 已安装软件与跨系统兼容情况</h2>
<p class="muted">✅ 官方支持 · ❌ 无官方版需找替代 · ➖ 特殊/未知</p>
<table><tr><th>软件</th><th>版本</th><th>发布者</th><th>大小</th><th>安装目录</th>
<th>新Windows</th><th>macOS</th><th>Linux</th><th>迁移建议</th></tr>
{''.join(rows_s)}</table></div>

{dev_html}

<div class="card"><h2>6. 换系统通用检查清单（打印或照做）</h2>
<ul>{rows_chk}</ul></div>

<p class="muted" style="text-align:center">由 e换系统 v{APP_VERSION} 生成 · 报告中的路径以扫描时为准</p>
</div></body></html>"""
    return html


def report_path():
    base = known_folder("desktop") or os.path.expanduser("~")
    return os.path.join(base, "换系统迁移清单_%s.html" % datetime.datetime.now().strftime("%Y%m%d_%H%M"))


# ----------------------------------------------------------------------------
# CSV 导出
# ----------------------------------------------------------------------------

def export_files_csv(path, items):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["类别", "位置", "大小", "字节数", "文件数", "是否勾选迁移", "迁移建议"])
        for cat, p, size, files, tip, checked in items:
            w.writerow([cat, p, fmt_size(size), size, files, "是" if checked else "否", tip])


def export_software_csv(path, software):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["软件", "版本", "发布者", "大小MB", "安装目录", "来源", "新Windows", "macOS", "Linux", "迁移建议"])
        for a in software:
            cat, win, mac, lin, tip = match_kb(a["name"], a["publisher"])
            w.writerow([a["name"], a["version"], a["publisher"], a["size_mb"] or "",
                        a["location"], a["kind"], win, mac, lin, tip])


# ----------------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------------

def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.title(f"{APP_NAME} — 换系统迁移助手 v{APP_VERSION}")
    root.geometry("1180x760")
    root.minsize(980, 620)

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    base_font = ("Microsoft YaHei UI", 10)
    root.option_add("*Font", base_font)
    style.configure("Treeview", rowheight=26, font=base_font)
    style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))
    style.configure("Big.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(14, 6))
    style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"))
    style.configure("Sub.TLabel", foreground="#666")

    q = queue.Queue()
    app = {"file_rows": [], "software": [], "top_files": [], "cancelled": False,
           "report_file": None, "stop": threading.Event(),
           "files_busy": False, "sw_busy": False}

    # ---- 通用：底部状态栏 -------------------------------------------------
    status_var = tk.StringVar(value="就绪。建议先在「① 个人文件」点击 快速扫描。")
    progress = ttk.Progressbar(root, mode="determinate", length=280)

    # ---- 头部 -------------------------------------------------------------
    head = ttk.Frame(root, padding=(16, 12, 16, 4))
    head.pack(fill="x")
    ttk.Label(head, text="e换系统", style="Title.TLabel",
              foreground="#1e66f5").pack(side="left")
    ttk.Label(head, style="Sub.TLabel",
              text="  换系统前，快速定位哪些内容需要搬 · 个人文件 + 软件清单 + 跨系统兼容"
              ).pack(side="left", padx=10)

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=12, pady=6)

    # ======================= Tab1 个人文件 ================================
    tab1 = ttk.Frame(nb, padding=8)
    nb.add(tab1, text="  ① 个人文件  ")

    opt = ttk.Frame(tab1)
    opt.pack(fill="x", pady=(0, 6))
    opts = {}
    for key, label, default in [
        ("user", "用户目录(桌面/文档/下载等)", True),
        ("chat", "微信/QQ/Telegram 数据", True),
        ("browser", "浏览器书签密码", True),
        ("dev", "开发环境(SSH密钥/配置)", True),
        ("office", "邮件与办公数据", True),
        ("deep", "深度扫描其他磁盘(较慢)", False),
    ]:
        var = tk.BooleanVar(value=default)
        opts[key] = var
        ttk.Checkbutton(opt, text=label, variable=var).pack(side="left", padx=(0, 10))

    btnbar = ttk.Frame(tab1)
    btnbar.pack(fill="x", pady=(0, 6))
    scan_btn = ttk.Button(btnbar, text="🚀 快速扫描", style="Big.TButton")
    scan_btn.pack(side="left")
    stop_btn = ttk.Button(btnbar, text="🛑 停止", state="disabled")
    stop_btn.pack(side="left", padx=8)
    files_sum = tk.StringVar(value="")
    ttk.Label(btnbar, textvariable=files_sum, foreground="#0a7d32").pack(side="left", padx=16)

    cols = ("sel", "cat", "path", "size", "files", "tip")
    ftv = ttk.Treeview(tab1, columns=cols, show="headings", selectmode="browse")
    for cid, text, w, anchor in [
        ("sel", "选", 40, "center"), ("cat", "类别", 120, "w"),
        ("path", "位置", 380, "w"), ("size", "大小", 90, "e"),
        ("files", "文件数", 80, "e"), ("tip", "迁移建议", 330, "w"),
    ]:
        ftv.heading(cid, text=text)
        ftv.column(cid, width=w, anchor=anchor, stretch=(cid == "path"))
    fsb = ttk.Scrollbar(tab1, orient="vertical", command=ftv.yview)
    ftv.configure(yscrollcommand=fsb.set)
    ftv.pack(side="left", fill="both", expand=True)
    fsb.pack(side="right", fill="y")
    ftv.tag_configure("parent", background="#eef2fb", font=("Microsoft YaHei UI", 10, "bold"))
    ftv.tag_configure("off", foreground="#999")

    def toggle_row(event=None):
        sel = ftv.selection()
        if not sel:
            return
        iid = sel[0]
        for row in app["file_rows"]:
            if row["iid"] == iid:
                row["checked"] = not row["checked"]
                ftv.set(iid, "sel", "☑" if row["checked"] else "☐")
                if not row["checked"]:
                    ftv.item(iid, tags=("off",))
                else:
                    ftv.item(iid, tags=())
                update_sum()

    def open_selected():
        for iid in ftv.selection():
            vals = ftv.item(iid, "values")
            p = vals[2]
            if p and os.path.isdir(p):
                os.startfile(p)

    def copy_path():
        for iid in ftv.selection():
            root.clipboard_clear()
            root.clipboard_append(ftv.item(iid, "values")[2])
            break

    fmenu = tk.Menu(ftv, tearoff=0)
    fmenu.add_command(label="打开位置", command=open_selected)
    fmenu.add_command(label="复制路径", command=copy_path)
    fmenu.add_command(label="勾选/取消勾选（双击同效）", command=toggle_row)

    def popup_f(event):
        iid = ftv.identify_row(event.y)
        if iid:
            ftv.selection_set(iid)
            fmenu.tk_popup(event.x_root, event.y_root)

    ftv.bind("<Double-1>", toggle_row)
    ftv.bind("<space>", toggle_row)
    ftv.bind("<Button-3>", popup_f)

    def update_sum():
        total = sum(r["size"] for r in app["file_rows"] if r["checked"])
        n = sum(1 for r in app["file_rows"] if r["checked"])
        files_sum.set(f"已勾选 {n} 项，共 {fmt_size(total)}")

    def start_files_scan():
        if app["files_busy"]:
            messagebox.showinfo(APP_NAME, "正在扫描中，请稍候或先停止。")
            return
        app["files_busy"] = True
        app["cancelled"] = False
        app["stop"].clear()
        app["file_rows"] = []
        app["top_files"] = []
        ftv.delete(*ftv.get_children())
        scan_btn.config(state="disabled")
        stop_btn.config(state="normal")
        progress.config(mode="indeterminate")
        progress.start(12)
        kwargs = dict(
            include_user=opts["user"].get(), include_chat=opts["chat"].get(),
            include_browser=opts["browser"].get(), include_dev=opts["dev"].get(),
            include_office=opts["office"].get(), deep=opts["deep"].get(),
        )
        threading.Thread(target=scan_files_worker, args=(kwargs,), daemon=True).start()

    def scan_files_worker(kwargs):
        try:
            stop = app["stop"]
            targets = build_scan_targets(
                include_user=kwargs["include_user"], include_chat=kwargs["include_chat"],
                include_browser=kwargs["include_browser"], include_dev=kwargs["include_dev"],
                include_office=kwargs["include_office"])
            q.put(("msg", f"开始扫描 {len(targets)} 个已知位置…"))
            rows = []
            for cat, path, tip in targets:
                if stop.is_set():
                    break
                q.put(("msg", f"统计 {path}"))
                size, files, _ = dir_stats(path, stop, app["top_files"])
                rows.append({"cat": cat, "path": path, "size": size,
                             "files": files, "tip": tip, "checked": size > 0})
                q.put(("files_row", rows[-1]))
            cancelled = stop.is_set()
            if kwargs.get("deep") and not stop.is_set():
                q.put(("msg", "深度扫描其他磁盘…"))
                for cat, path, size, files, tip in scan_deep_drives(
                        stop, progress=lambda i, n, p: q.put(("msg", f"深度扫描 {p}"))):
                    rows.append({"cat": cat, "path": path, "size": size,
                                 "files": files, "tip": tip, "checked": size > 0})
                    q.put(("files_row", rows[-1]))
            q.put(("files_done", rows, cancelled))
        except Exception:
            q.put(("error", traceback.format_exc()))

    scan_btn.config(command=start_files_scan)
    stop_btn.config(command=lambda: (app["stop"].set(), status_var.set("正在停止…")))

    def export_files():
        if not app["file_rows"]:
            messagebox.showinfo(APP_NAME, "请先扫描个人文件。")
            return
        p = filedialog.asksaveasfilename(
            defaultextension=".csv", initialfile="个人文件清单.csv",
            filetypes=[("CSV 文件", "*.csv")], parent=root)
        if p:
            items = [(r["cat"], r["path"], r["size"], r["files"], r["tip"], r["checked"])
                     for r in app["file_rows"]]
            export_files_csv(p, items)
            status_var.set("已导出：" + p)

    ttk.Button(btnbar, text="📁 打开位置", command=open_selected).pack(side="right", padx=4)
    ttk.Button(btnbar, text="📤 导出CSV", command=export_files).pack(side="right", padx=4)

    # ======================= Tab2 软件清单 =================================
    tab2 = ttk.Frame(nb, padding=8)
    nb.add(tab2, text="  ② 软件清单  ")

    bar2 = ttk.Frame(tab2)
    bar2.pack(fill="x", pady=(0, 6))
    sw_btn = ttk.Button(bar2, text="🔍 扫描已安装软件", style="Big.TButton")
    sw_btn.pack(side="left")
    uwp_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(bar2, text="包含商店应用", variable=uwp_var).pack(side="left", padx=10)
    ttk.Label(bar2, text="过滤:").pack(side="left", padx=(18, 4))
    search_var = tk.StringVar()
    search_ent = ttk.Entry(bar2, textvariable=search_var, width=22)
    search_ent.pack(side="left")
    sw_sum = tk.StringVar(value="")
    ttk.Label(bar2, textvariable=sw_sum).pack(side="left", padx=12)
    ttk.Button(bar2, text="📤 导出CSV",
               command=lambda: export_sw()).pack(side="right", padx=4)

    scols = ("name", "ver", "pub", "size", "loc", "w", "m", "l", "tip")
    stv = ttk.Treeview(tab2, columns=scols, show="headings", selectmode="browse")
    for cid, text, w, anchor in [
        ("name", "软件", 210, "w"), ("ver", "版本", 90, "w"),
        ("pub", "发布者", 140, "w"), ("size", "大小", 75, "e"),
        ("loc", "安装目录", 300, "w"), ("w", "新Win", 55, "center"),
        ("m", "macOS", 62, "center"), ("l", "Linux", 55, "center"),
        ("tip", "迁移建议", 320, "w"),
    ]:
        stv.heading(cid, text=text)
        stv.column(cid, width=w, anchor=anchor, stretch=(cid in ("name", "loc", "tip")))
    ssb = ttk.Scrollbar(tab2, orient="vertical", command=stv.yview)
    stv.configure(yscrollcommand=ssb.set)
    stv.pack(side="left", fill="both", expand=True)
    ssb.pack(side="right", fill="y")

    def render_software():
        kw = search_var.get().strip().lower()
        stv.delete(*stv.get_children())
        n = 0
        for a in app["software"]:
            if kw and kw not in (a["name"] + a["publisher"] + a["location"]).lower():
                continue
            _, w, m, l, tip = match_kb(a["name"], a["publisher"])
            stv.insert("", "end", values=(
                a["name"], a["version"], a["publisher"][:20],
                f"{a['size_mb']:.0f} MB" if a["size_mb"] else "—",
                a["location"] or "—", w, m, l, tip))
            n += 1
        sw_sum.set(f"共 {n} 个")

    search_var.trace_add("write", lambda *a: render_software())

    def open_sw_dir():
        for iid in stv.selection():
            p = stv.item(iid, "values")[4]
            if p and os.path.isdir(p):
                os.startfile(p)
                return
        messagebox.showinfo(APP_NAME, "该项没有记录到安装目录。")

    smenu = tk.Menu(stv, tearoff=0)
    smenu.add_command(label="打开安装目录", command=open_sw_dir)
    smenu.add_command(label="复制软件名", command=lambda: (
        root.clipboard_clear(), root.clipboard_append(stv.item(stv.selection()[0], "values")[0])
        if stv.selection() else None))

    def popup_s(event):
        iid = stv.identify_row(event.y)
        if iid:
            stv.selection_set(iid)
            smenu.tk_popup(event.x_root, event.y_root)

    stv.bind("<Button-3>", popup_s)

    def start_sw_scan():
        if app["sw_busy"]:
            messagebox.showinfo(APP_NAME, "正在扫描中，请稍候。")
            return
        app["sw_busy"] = True
        sw_btn.config(state="disabled")
        status_var.set("正在读取软件清单…")
        threading.Thread(target=sw_worker, daemon=True).start()

    def sw_worker():
        try:
            sw = scan_installed_software(include_uwp=uwp_var.get())
            q.put(("sw_done", sw))
        except Exception:
            q.put(("error", traceback.format_exc()))

    sw_btn.config(command=start_sw_scan)

    def export_sw():
        if not app["software"]:
            messagebox.showinfo(APP_NAME, "请先扫描软件清单。")
            return
        p = filedialog.asksaveasfilename(
            defaultextension=".csv", initialfile="软件清单.csv",
            filetypes=[("CSV 文件", "*.csv")], parent=root)
        if p:
            export_software_csv(p, app["software"])
            status_var.set("已导出：" + p)

    # ======================= Tab3 迁移报告 =================================
    tab3 = ttk.Frame(nb, padding=8)
    nb.add(tab3, text="  ③ 迁移报告  ")

    bar3 = ttk.Frame(tab3)
    bar3.pack(fill="x", pady=(0, 6))
    ttk.Button(bar3, text="📝 生成迁移报告 (HTML)", style="Big.TButton",
               command=lambda: gen_report()).pack(side="left")
    ttk.Button(bar3, text="🌐 打开已生成的报告", command=lambda: (
        webbrowser.open(app["report_file"]) if app["report_file"]
        else messagebox.showinfo(APP_NAME, "还没有生成报告。"))).pack(side="left", padx=8)
    ttk.Label(bar3, style="Sub.TLabel",
              text="汇总个人文件 + 软件兼容情况 + 通用检查清单，保存到桌面").pack(side="left", padx=8)

    import tkinter.scrolledtext as st
    preview = st.ScrolledText(tab3, height=18, state="disabled",
                              font=("Microsoft YaHei UI", 10))
    preview.pack(fill="both", expand=True)

    def gen_report():
        if not app["file_rows"]:
            messagebox.showinfo(APP_NAME, "请先在「① 个人文件」完成扫描。")
            return
        software = app["software"]
        if not software:
            if messagebox.askyesno(APP_NAME, "软件清单还没扫描。现在扫描并写进报告吗？"):
                status_var.set("正在读取软件清单…")
                root.update_idletasks()
                try:
                    software = scan_installed_software(include_uwp=uwp_var.get())
                    app["software"] = software
                    render_software()
                except Exception as e:
                    software = []
                    status_var.set("软件清单读取失败：" + str(e))
        status_var.set("正在生成报告（含开发环境依赖清单）…")
        root.update_idletasks()
        dev = dev_package_lists() if any(r["cat"] == "开发环境" for r in app["file_rows"]) else []
        items = [(r["cat"], r["path"], r["size"], r["files"], r["tip"], r["checked"])
                 for r in app["file_rows"]]
        total = sum(r["size"] for r in app["file_rows"] if r["checked"])
        html = build_report(items, software, app["top_files"], dev, total, app["cancelled"])
        try:
            p = report_path()
            with open(p, "w", encoding="utf-8") as f:
                f.write(html)
        except PermissionError:
            from tkinter import filedialog
            p = filedialog.asksaveasfilename(defaultextension=".html",
                                             initialfile="换系统迁移清单.html", parent=root)
            if not p:
                return
            with open(p, "w", encoding="utf-8") as f:
                f.write(html)
        app["report_file"] = p
        preview.config(state="normal")
        preview.delete("1.0", "end")
        preview.insert("end",
                       f"✅ 迁移报告已生成：\n{p}\n\n"
                       f"个人文件：{len(items)} 处，勾选合计 {fmt_size(total)}\n"
                       f"大文件 TOP：{len(app['top_files'])} 个\n"
                       f"软件：{len(software)} 个，其中已给出兼容建议 "
                       f"{sum(1 for a in software if match_kb(a['name'], a['publisher'])[0] != '未收录')} 个\n"
                       f"开发依赖清单：{len(dev)} 份\n\n"
                       f"通用检查清单 {len(CHECKLIST)} 条已附在报告末尾。\n"
                       f"点击「打开已生成的报告」在浏览器查看/打印。")
        preview.config(state="disabled")
        status_var.set("报告已生成：" + p)
        webbrowser.open(p)

    # ---- 队列轮询 ---------------------------------------------------------
    def poll():
        try:
            while True:
                kind, *payload = q.get_nowait()
                if kind == "msg":
                    status_var.set(payload[0])
                elif kind == "files_row":
                    row = payload[0]
                    iid = ftv.insert("", "end", values=(
                        "☑" if row["checked"] else "☐", row["cat"], row["path"],
                        fmt_size(row["size"]), f"{row['files']:,}", row["tip"]))
                    row["iid"] = iid
                    if not row["checked"]:
                        ftv.item(iid, tags=("off",))
                    ftv.see(iid)
                elif kind == "files_done":
                    rows, cancelled = payload
                    app["file_rows"] = rows
                    app["cancelled"] = cancelled
                    app["files_busy"] = False
                    progress.stop()
                    progress.config(mode="determinate")
                    scan_btn.config(state="normal")
                    stop_btn.config(state="disabled")
                    # 按类别分组展示
                    cats = {}
                    for r in rows:
                        cats.setdefault(r["cat"], []).append(r)
                    ftv.delete(*ftv.get_children())
                    for cat, list_ in cats.items():
                        csum = sum(x["size"] for x in list_)
                        cfiles = sum(x["files"] for x in list_)
                        parent = ftv.insert("", "end", values=(
                            "", cat, f"（{len(list_)} 项）", fmt_size(csum),
                            f"{cfiles:,}", ""), tags=("parent",))
                        for r in sorted(list_, key=lambda x: -x["size"]):
                            iid = ftv.insert(parent, "end", values=(
                                "☑" if r["checked"] else "☐", r["cat"], r["path"],
                                fmt_size(r["size"]), f"{r['files']:,}", r["tip"]))
                            r["iid"] = iid
                            if not r["checked"]:
                                ftv.item(iid, tags=("off",))
                        ftv.item(parent, open=True)
                    update_sum()
                    status_var.set("扫描完成" + ("（已提前停止，结果可能不全）" if cancelled else "")
                                   + f"。共 {len(rows)} 项，双击行可勾选/取消。")
                elif kind == "sw_done":
                    app["software"] = payload[0]
                    app["sw_busy"] = False
                    sw_btn.config(state="normal")
                    render_software()
                    status_var.set(f"软件清单就绪：{len(app['software'])} 个。可切换到「③ 迁移报告」。")
                elif kind == "error":
                    app["files_busy"] = False
                    app["sw_busy"] = False
                    progress.stop()
                    progress.config(mode="determinate")
                    scan_btn.config(state="normal")
                    sw_btn.config(state="normal")
                    stop_btn.config(state="disabled")
                    status_var.set("发生错误，详见弹窗。")
                    messagebox.showerror(APP_NAME, "扫描出错：\n\n" + payload[0])
        except queue.Empty:
            pass
        root.after(80, poll)

    root.after(80, poll)

    # 底部状态栏
    bottom = ttk.Frame(root, padding=(16, 2, 16, 8))
    bottom.pack(fill="x")
    ttk.Label(bottom, textvariable=status_var, style="Sub.TLabel").pack(side="left")
    progress.pack(side="right")

    # 自动测试模式：自动扫描便于截图
    if os.environ.get("EHS_AUTOTEST"):
        root.after(600, start_files_scan)
        root.after(2500, start_sw_scan)
        root.after(22000, lambda: nb.select(tab2))
        root.after(45000, root.destroy)

    root.mainloop()


# ----------------------------------------------------------------------------
# 无头模式：--report / --selftest
# ----------------------------------------------------------------------------

def headless_report():
    _utf8_stdout()
    print("e换系统 无界面模式：全量扫描并生成报告…")
    stop = threading.Event()
    targets = build_scan_targets()
    rows = []
    top_files = []
    for cat, path, tip in targets:
        print(f"  统计 {path} …")
        size, files, _ = dir_stats(path, stop, top_files)
        rows.append({"cat": cat, "path": path, "size": size, "files": files,
                     "tip": tip, "checked": True})
    print("  深度扫描其他磁盘…")
    for cat, path, size, files, tip in scan_deep_drives(stop):
        rows.append({"cat": cat, "path": path, "size": size, "files": files,
                     "tip": tip, "checked": True})
    print("  读取软件清单…")
    software = scan_installed_software()
    dev = dev_package_lists()
    items = [(r["cat"], r["path"], r["size"], r["files"], r["tip"], r["checked"]) for r in rows]
    total = sum(r["size"] for r in rows)
    html = build_report(items, software, top_files, dev, total, stop.is_set())
    p = report_path()
    with open(p, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n✅ 报告已生成: {p}")
    print(f"   个人文件 {len(rows)} 处 / 软件 {len(software)} 个")
    return p


def selftest():
    _utf8_stdout()
    ok = True

    print("== 1. 已知文件夹 ==")
    for k in ("desktop", "documents", "downloads", "pictures", "videos", "music"):
        p = known_folder(k)
        print(f"  {k}: {p}")
        if not p:
            ok = False

    print("== 2. 磁盘 ==")
    drives = list_fixed_drives()
    print("  驱动器:", drives)
    if not drives:
        ok = False

    print("== 3. 扫描目标存在性 ==")
    targets = build_scan_targets()
    print(f"  命中 {len(targets)} 个已存在的目标位置")
    for cat, p, tip in targets[:40]:
        print(f"   [{cat}] {p}")
    if not targets:
        ok = False

    print("== 4. 目录统计（取第一个目标） ==")
    if targets:
        s, f, d = dir_stats(targets[0][1], None, None)
        print(f"  {targets[0][1]} -> {fmt_size(s)}, {f} 文件, {d} 子目录")
        assert isinstance(s, int)

    print("== 5. 软件枚举（仅注册表） ==")
    sw = scan_installed_software(include_uwp=False)
    print(f"  注册表软件 {len(sw)} 个")
    for a in sw[:15]:
        cat, w, m, l, tip = match_kb(a["name"], a["publisher"])
        print(f"   {a['name'][:36]:38} {w}{m}{l} [{cat}] {a['location'][:40]}")
    if not sw:
        ok = False

    print("== 6. 兼容性知识库抽查 ==")
    samples = [("Google Chrome", "Google LLC"), ("微信", "Tencent"), ("WeChat", "Tencent"),
               ("7-Zip 23.01", "Igor Pavlov"), ("Steam", "Valve"),
               ("PotPlayer", "Kakao"), ("WPS Office", "Kingsoft"),
               ("Microsoft Visual Studio Code", "Microsoft"), ("不存在的软件XYZ", "")]
    for name, pub in samples:
        cat, w, m, l, tip = match_kb(name, pub)
        hit = cat != "未收录"
        print(f"  {'OK ' if hit else '-- '} {name:36} Win{w} Mac{m} Linux{l} [{cat}]")
        if not hit and name != "不存在的软件XYZ":
            ok = False

    print("== 7. 报告生成 ==")
    items = [(c, p, 123 * MB, 10, t, True) for c, p, t in targets[:5]]
    html = build_report(items, sw[:5], [(2 * GB, "D:\\big\\video.mp4")], [], 615 * MB)
    assert "换系统迁移清单" in html and len(html) > 3000
    tmp = os.path.join(os.environ.get("TEMP", "."), "ehs_selftest.html")
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  样例报告 {len(html)} 字符 -> {tmp}")

    print("== 8. CSV 导出 ==")
    csv1 = tmp.replace(".html", "_f.csv")
    csv2 = tmp.replace(".html", "_s.csv")
    export_files_csv(csv1, items)
    export_software_csv(csv2, sw[:5])
    print(f"  {csv1}\n  {csv2}")

    print("\n" + ("✅ 自测全部通过" if ok else "⚠ 自测存在失败项"))
    return 0 if ok else 1


def main():
    if "--selftest" in sys.argv:
        sys.exit(selftest())
    if "--report" in sys.argv:
        sys.exit(0 if headless_report() else 1)
    run_gui()


if __name__ == "__main__":
    main()
