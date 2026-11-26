#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分卷压缩包分批解压工具 v1.3
===========================

功能：按顺序解压分卷，在确认后续分卷已开始读取、当前分卷数据已写入输出文件后，
      自动删除“上一个已处理完毕”的分卷，避免需要双倍磁盘空间。

支持：
  * RAR 分卷：xxx.part1.rar / xxx.part01.rar / xxx.rar.001 / xxx.r00+xxx.rar
      —— 使用随程序附带的官方 UnRAR.exe（RARLAB），解析其 "Extracting from" 行作为分卷切换信号
  * ZIP 分卷：xxx.zip.001 / xxx.001 / xxx.z01+xxx.zip
      —— 纯 Python 流式读取（多卷虚拟流 + zipfile），按成员hash校验后逐卷回收

依赖：Python 3.9+（仅标准库），RAR 解压依赖随包附带的 tools/UnRAR.exe（x64 Windows）

命令行模式（便于批处理/测试）：
  python 分卷解压工具.py --cli <第一分卷或目录> --output <输出目录> [--precheck] [--auto-delete] [--password 密码]

右键菜单集成：
  分卷解压工具.exe --install-menu     注册资源管理器右键菜单（当前用户，无需管理员）
  分卷解压工具.exe --uninstall-menu   卸载右键菜单
  右键菜单内部调用：分卷解压工具.exe --path "<选中的路径>"
"""

from __future__ import annotations

import bisect
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import zipfile
import winreg
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional, Tuple

# 打包成单文件 exe 后，__file__ 指向临时解包目录，资源从 _MEIPASS 读取；
# 审计日志仍写到 exe 所在目录，避免退出后被清理。
FROZEN = getattr(sys, "frozen", False)
if FROZEN:
    BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    APP_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
    APP_DIR = BASE_DIR
TOOLS_DIR = BASE_DIR / "tools"
AUDIT_DIR = APP_DIR / "logs"
APP_TITLE = "分卷压缩包分批解压工具 v1.3"
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
SIGNATURE = "by东风敬我桂花酒"


def prepare_console() -> None:
    """GUI 版 exe 不带控制台；命令行模式需要时再申请一个，避免启动黑框。"""
    is_cli = any(a in sys.argv for a in ("--cli", "--install-menu", "--uninstall-menu"))
    if os.name == "nt" and is_cli:
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            if not kernel32.GetConsoleWindow():
                # 从终端启动时直接挂到父终端，避免多弹一个窗口
                kernel32.AttachConsole(-1)  # ATTACH_PARENT_PROCESS
            if not kernel32.GetConsoleWindow() and kernel32.AllocConsole():
                    kernel32.SetConsoleOutputCP(65001)  # UTF-8，保证中文日志正常
                    sys.stdout = open("CONOUT$", "w", encoding="utf-8",
                                      errors="replace", buffering=1)
                    sys.stderr = open("CONOUT$", "w", encoding="utf-8",
                                      errors="replace", buffering=1)
                    try:
                        sys.stdin = open("CONIN$", "r", encoding="utf-8", errors="replace")
                    except Exception:
                        pass
        except Exception:
            pass
    # 无控制台的窗口模式下 stdout/stderr 可能是 None，给个空接收器防止异常
    for name in ("stdout", "stderr"):
        if getattr(sys, name, None) is None:
            setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))


# ---------------------------------------------------------------------------
# 资源管理器右键菜单集成（安装/卸载）
# ---------------------------------------------------------------------------

MENU_KEY_NAME = "VolumeExtractTool"
MENU_TITLE = "分卷解压工具"

# 级联子菜单：verb -> (显示文字, 命令行参数模板)
# {exe} 与 {path} 在注册时替换；%1/%V 由资源管理器传入
MENU_ITEMS = [
    ("VE.Open",     "用分卷解压工具打开...",        '--path "{path}"'),
    ("VE.Here",     "解压到此(E)",                  '--quick here "{path}"'),
    ("VE.Smart",    "智能解压到此(S)",              '--quick smart "{path}"'),
    ("VE.Subfolder", "解压到同名文件夹(&T)",        '--quick subfolder "{path}"'),
    ("VE.Browse",   "解压到其他文件夹(B)...",        '--quick browse "{path}"'),
]
COMMAND_STORE = r"Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell"

MENU_BASES = [
    (r"Software\Classes\*\shell", "%1"),                     # 文件（含分卷压缩包）
    (r"Software\Classes\Directory\shell", "%1"),             # 文件夹
    (r"Software\Classes\Directory\Background\shell", "%V"),  # 文件夹空白处
]


def _verb_command(template: str, exe: str, param: str) -> str:
    return f'"{exe}" ' + template.replace("{path}", param)


def menu_install_status() -> str:
    """返回：'missing'（未安装）/ 'ok'（指向当前 exe）/ 'moved'（指向其他位置）。"""
    if os.name != "nt":
        return "missing"
    exe = str(Path(sys.executable).resolve()).lower()
    found = 0
    matched = 0
    for base, _ in MENU_BASES:
        try:
            k = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                               base + "\\" + MENU_KEY_NAME)
            val, _ = winreg.QueryValueEx(k, "SubCommands")
            winreg.CloseKey(k)
            found += 1
            if str(val):
                matched += 1
        except FileNotFoundError:
            pass
        except Exception:
            pass
    if found == 0:
        return "missing"
    return "ok" if matched == found else "moved"


def install_context_menu(emit=None) -> bool:
    """把级联右键菜单写入当前用户注册表（无需管理员权限）。"""
    say = emit or (lambda m: print(m))
    if os.name != "nt":
        say("右键菜单仅支持 Windows")
        return False
    if not FROZEN:
        say("请用【分卷解压工具.exe】操作：源码模式注册的是 Python 解释器，不能用于右键。")
        return False

    exe = str(Path(sys.executable).resolve())
    try:
        # 1) 先在 CommandStore 注册每个子命令
        for verb, label, template in MENU_ITEMS:
            store_key = COMMAND_STORE + "\\" + verb
            k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, store_key)
            winreg.SetValueEx(k, "", 0, winreg.REG_SZ, label)
            winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, exe)
            winreg.CloseKey(k)
            # 子命令的 command 里放占位符；实际路径参数通过下面的 per-shell 键覆盖
            cmd = winreg.CreateKey(winreg.HKEY_CURRENT_USER, store_key + "\\command")
            winreg.SetValueEx(cmd, "", 0, winreg.REG_SZ, _verb_command(template, exe, "%1"))
            winreg.CloseKey(cmd)

        # 2) 三处主菜单用 SubCommands 引用子命令
        verbs = ";".join(v for v, _, _ in MENU_ITEMS)
        for base, param in MENU_BASES:
            key_path = base + "\\" + MENU_KEY_NAME
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, MENU_TITLE)
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, exe)
            winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ, verbs)
            try:
                winreg.DeleteKey(key, "command")  # 清理旧版单层菜单残留
            except Exception:
                pass
            winreg.CloseKey(key)
        say("右键级联菜单安装成功，已注册：文件、文件夹、文件夹空白处三处。")
        say("菜单项：打开 / 解压到此 / 智能解压到此 / 解压到同名文件夹 / 解压到其他文件夹")
        return True
    except Exception as exc:
        say(f"安装失败：{exc}")
        return False


def uninstall_context_menu(emit=None) -> bool:
    """移除右键菜单条目，不影响程序本身。"""
    say = emit or (lambda m: print(m))
    if os.name != "nt":
        say("右键菜单仅支持 Windows")
        return False
    removed = 0
    for base, _ in MENU_BASES:
        key = base + "\\" + MENU_KEY_NAME
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key + "\\command")
        except Exception:
            pass
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key)
            removed += 1
        except FileNotFoundError:
            pass
        except Exception as exc:
            say(f"清理失败 {key}: {exc}")
    for verb, _, _ in MENU_ITEMS:
        store_key = COMMAND_STORE + "\\" + verb
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, store_key + "\\command")
        except Exception:
            pass
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, store_key)
        except Exception:
            pass
    say(f"右键菜单已卸载，移除 {removed} 处主菜单及子命令。")
    return True


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def now_str() -> str:
    return datetime.now().strftime("%H:%M:%S")


def stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def free_gb(path) -> float:
    """返回目标所在磁盘的剩余空间（GB）"""
    try:
        p = Path(path)
        if not p.exists():
            p = p.parent
        while p and not p.exists():
            p = p.parent
        return shutil.disk_usage(str(p) if p else ".").free / (1024 ** 3)
    except Exception:
        return -1.0


def human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{n:.2f} {unit}" if unit != "B" else f"{n} B"
        n /= 1024.0
    return f"{n:.2f} TB"


def decode_line(raw: bytes) -> str:
    """UnRAR 在中文 Windows 下输出 GBK，在其它系统可能输出 UTF-8"""
    if not raw:
        return ""
    for enc in ("utf-8", "gbk", "cp936", "latin-1"):
        try:
            return raw.decode(enc).rstrip("\r\n")
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace").rstrip("\r\n")


class StopRequested(Exception):
    pass


@dataclass
class TaskContext:
    emit: Callable[[str, str], None]
    stop: threading.Event
    error_seen: bool = False

    def log(self, msg: str, level: str = "INFO") -> None:
        self.emit(level, msg)

    def progress(self, percent: float) -> None:
        self.emit("__progress__", float(percent))

    def check_stop(self) -> None:
        if self.stop.is_set():
            raise StopRequested()


# ---------------------------------------------------------------------------
# 分卷识别
# ---------------------------------------------------------------------------

RE_PART_EXT = re.compile(r"^(?P<base>.+?)\.part(?P<num>\d{1,5})\.(?P<ext>[A-Za-z0-9]+)$", re.I)
RE_EXT_NUM = re.compile(r"^(?P<base>.+?)\.(?P<ext>[A-Za-z0-9]+)\.(?P<num>\d{1,5})$", re.I)
RE_PLAIN_NUM = re.compile(r"^(?P<base>.+?)\.(?P<num>\d{1,5})$", re.I)
RE_LEGACY_R = re.compile(r"^(?P<base>.+?)\.r(?P<num>\d{2})$", re.I)
RE_LEGACY_Z = re.compile(r"^(?P<base>.+?)\.z(?P<num>\d{2})$", re.I)
RE_LEGACY_MAIN_R = re.compile(r"^(?P<base>.+?)\.rar$", re.I)
RE_LEGACY_MAIN_Z = re.compile(r"^(?P<base>.+?)\.zip$", re.I)


@dataclass
class VolumeSet:
    base: str
    folder: Path
    volumes: List[Path]
    nums: List[int]
    kind: str                  # rar / zip / unknown
    note: str = ""

    @property
    def count(self) -> int:
        return len(self.volumes)

    def label(self, i: int) -> str:
        return f"part{self.nums[i]}"


def sniff_kind(path: Path) -> str:
    try:
        with open(path, "rb") as f:
            head = f.read(16)
    except Exception:
        return "unknown"
    if head.startswith(b"Rar!\x1a\x07"):
        return "rar"
    if head.startswith(b"PK\x03\x04") or head.startswith(b"PK\x05\x06"):
        return "zip"
    if head.startswith(b"7z\xbc\xaf\x27\x1c"):
        return "7z"
    return "unknown"


def _scan_groups(folder: Path) -> dict:
    """返回 key=(base.lower(), family) -> {num: Path}"""
    groups: dict = {}

    def put(base: str, family: str, num: int, path: Path) -> None:
        key = (base.lower(), family)
        groups.setdefault(key, {"base": base, "family": family, "items": {}})
        # 同名编号保留体积更大的那个，避免重复
        old = groups[key]["items"].get(num)
        if old is None or path.stat().st_size > old.stat().st_size:
            groups[key]["items"][num] = path

    for p in folder.iterdir():
        if not p.is_file():
            continue
        name = p.name
        try:
            m = RE_PART_EXT.match(name)
            if m:
                put(m.group("base"), "part." + m.group("ext").lower(), int(m.group("num")), p)
                continue
            m = RE_EXT_NUM.match(name)
            if m:
                put(m.group("base"), "ext." + m.group("ext").lower(), int(m.group("num")), p)
                continue
            m = RE_LEGACY_R.match(name)
            if m:
                put(m.group("base"), "legacy.rar", int(m.group("num")) + 1, p)
                continue
            m = RE_LEGACY_Z.match(name)
            if m:
                # 传统 ZIP 分卷：.z01 在前，.zip（含中央目录）在最后，先用 9999 占位
                put(m.group("base"), "legacy.zip", int(m.group("num")), p)
                continue
            m = RE_LEGACY_MAIN_R.match(name)
            if m:
                put(m.group("base"), "legacy.rar", 0, p)
                continue
            m = RE_LEGACY_MAIN_Z.match(name)
            if m:
                put(m.group("base"), "legacy.zip", 9999, p)
                continue
            m = RE_PLAIN_NUM.match(name)
            if m:
                put(m.group("base"), "plain", int(m.group("num")), p)
                continue
        except OSError:
            continue
    return groups


def _build_set(folder: Path, g: dict, prefer_kind: str = "") -> VolumeSet:
    items = dict(g["items"])
    family = g["family"]
    if family == "legacy.zip" and 9999 in items:
        main = items.pop(9999)
        others = [n for n in items if n != 9999]
        items[(max(others) + 1) if others else 0] = main
    nums = sorted(items)
    volumes = [items[n] for n in nums]

    if family == "legacy.rar":
        kind = "rar"
    elif family == "legacy.zip":
        kind = "zip"
    elif family.startswith("part.") or family.startswith("ext."):
        ext = family.split(".", 1)[1]
        kind = "rar" if ext == "rar" else ("zip" if ext == "zip" else sniff_kind(volumes[0]))
    else:
        kind = sniff_kind(volumes[0]) if volumes else "unknown"
        if kind == "unknown" and prefer_kind:
            kind = prefer_kind

    # 连续性检查
    note = ""
    if nums:
        expected = list(range(nums[0], nums[0] + len(nums)))
        if expected != nums:
            missing = sorted(set(expected) - set(nums))
            note = "分卷编号不连续，缺失：" + ", ".join(str(x) for x in missing)

    return VolumeSet(g["base"], folder, volumes, nums, kind, note)


def detect_volume_set(path_input, chooser: Optional[Callable[[List[VolumeSet]], VolumeSet]] = None) -> VolumeSet:
    """path_input 可以是分卷文件，也可以是分卷所在目录"""
    p = Path(path_input).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(f"路径不存在：{p}")
    folder = p if p.is_dir() else p.parent
    groups = _scan_groups(folder)
    sets = [_build_set(folder, g) for g in groups.values() if g["items"]]

    if not sets:
        raise ValueError("未在指定位置识别出分卷文件（支持 .part1.rar / .rar.001 / .001 / .z01 等命名）")

    if p.is_file():
        # 优先选择包含该文件的组
        for vs in sets:
            if p in vs.volumes:
                return vs
        return min(sets, key=lambda s: s.nums[0])

    sets.sort(key=lambda s: (s.nums[0], s.base))
    if len(sets) == 1:
        return sets[0]
    if chooser is not None:
        chosen = chooser(sets)
        if chosen is not None:
            return chosen
    # 无法交互时选卷数最多的一组
    return max(sets, key=lambda s: (s.count, -s.nums[0]))


# ---------------------------------------------------------------------------
# UnRAR 定位
# ---------------------------------------------------------------------------

def find_unrar() -> Optional[Path]:
    candidates = [TOOLS_DIR / "UnRAR.exe", TOOLS_DIR / "unrar.exe"]
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        root = os.environ.get(env)
        if root:
            candidates.append(Path(root) / "WinRAR" / "UnRAR.exe")
            candidates.append(Path(root) / "WinRAR" / "unrar.exe")
    found = None
    for c in candidates:
        try:
            if c.is_file():
                out = subprocess.run([str(c)], capture_output=True, timeout=8,
                                     creationflags=CREATE_NO_WINDOW).stdout
                if b"UNRAR" in out:
                    found = c
                    break
        except Exception:
            continue
    if found is None:
        for name in ("unrar", "unrar.exe"):
            try:
                r = subprocess.run([name], capture_output=True, timeout=8,
                                   creationflags=CREATE_NO_WINDOW)
                if b"UNRAR" in r.stdout:
                    found = Path(shutil.which(name) or name)
                    break
            except Exception:
                continue
    return found


# ---------------------------------------------------------------------------
# 多卷虚拟流（供 zipfile 使用）
# ---------------------------------------------------------------------------

class MultiVolumeStream:
    """把多个分卷伪装成一个连续可读、可 seek 的二进制流。只读，不回头依赖已消费的分卷。"""

    def __init__(self, volumes: List[Path]):
        self.volumes = list(volumes)
        if not self.volumes:
            raise ValueError("分卷列表为空")
        self.sizes = [p.stat().st_size for p in self.volumes]
        self.bounds: List[int] = []
        acc = 0
        for s in self.sizes:
            acc += s
            self.bounds.append(acc)
        self.total = acc
        self.index = 0
        self.fp = None
        self.pos = 0
        self.pushback = b""

    def seekable(self) -> bool:
        return True

    def readable(self) -> bool:
        return True

    def writable(self) -> bool:
        return False

    def tell(self) -> int:
        return self.pos

    def seek(self, offset: int, whence: int = 0) -> int:
        if whence == 0:
            new = offset
        elif whence == 1:
            new = self.pos + offset
        elif whence == 2:
            new = self.total + offset
        else:
            raise ValueError("invalid whence")
        new = max(0, int(new))
        idx = min(bisect.bisect_right(self.bounds, new), len(self.volumes) - 1)
        # 必须无条件关闭并重置句柄：即使仍在同一分卷，文件指针也停留在旧位置，
        # 直接读会得到空数据（zipfile 定位 EOCD 时必现）。
        if self.fp:
            self.fp.close()
            self.fp = None
        self.index = idx
        self.pos = new
        self.pushback = b""
        return new

    def read(self, n: int = -1) -> bytes:
        if n is None or n < 0:
            n = max(0, self.total - self.pos)
        out = bytearray()
        while n > 0 and self.pos < self.total:
            if self.pushback:
                take = min(n, len(self.pushback))
                out.extend(self.pushback[:take])
                self.pushback = self.pushback[take:]
                self.pos += take
                n -= take
                continue
            if self.fp is None:
                path = self.volumes[self.index]
                if not path.exists():
                    raise FileNotFoundError(f"分卷已丢失，无法继续读取：{path.name}")
                self.fp = open(path, "rb")
                vol_start = self.bounds[self.index - 1] if self.index > 0 else 0
                self.fp.seek(self.pos - vol_start)
            chunk = self.fp.read(min(n, 1024 * 1024))
            if not chunk:
                self.fp.close()
                self.fp = None
                self.index += 1
                if self.index >= len(self.volumes):
                    break
                continue
            out.extend(chunk)
            self.pos += len(chunk)
            n -= len(chunk)
        return bytes(out)

    def close(self) -> None:
        if self.fp:
            self.fp.close()
            self.fp = None


# ---------------------------------------------------------------------------
# 删除管理器（所有删除的唯一入口）
# ---------------------------------------------------------------------------

class DeletionManager:
    """只删除同组中 index < 当前已确认进度的分卷，绝不碰输出文件。"""

    def __init__(self, enabled: bool, vset: VolumeSet, ctx: TaskContext, dest: Path):
        self.enabled = enabled
        self.vset = vset
        self.ctx = ctx
        self.dest = dest.resolve()
        self.deleted = set()
        AUDIT_DIR.mkdir(parents=True, exist_ok=True)
        self.audit = AUDIT_DIR / f"删除审计_{datetime.now():%Y%m%d}.log"

    def delete_through(self, last_index: int, reason: str) -> None:
        if not self.enabled or self.ctx.stop.is_set():
            return
        for i in range(0, min(last_index, self.vset.count - 1) + 1):
            self._delete_one(i, reason)

    def delete_all(self, reason: str) -> None:
        self.delete_through(self.vset.count - 1, reason)

    def _delete_one(self, i: int, reason: str) -> None:
        if i in self.deleted or i < 0 or i >= self.vset.count:
            return
        path = self.vset.volumes[i]
        self.deleted.add(i)
        if not path.exists():
            self.ctx.log(f"校验: {path.name} 已不存在，跳过删除 | 剩余磁盘空间: {free_gb(path):.2f} GB", "OK")
            return

        # 硬性安全护栏：只允许删除分卷目录内、属于本组、且不在输出目录中的路径
        try:
            rp = path.resolve()
            if path.parent.resolve() != self.vset.folder.resolve():
                raise RuntimeError("路径不在分卷目录内")
            if rp not in {p.resolve() for p in self.vset.volumes}:
                raise RuntimeError("路径不属于当前分卷组")
            if rp == self.dest or self.dest in rp.parents:
                raise RuntimeError("路径位于输出目录内，疑似成品文件")
        except Exception as exc:
            self.ctx.log(f"安全拦截：拒绝删除 {path.name}（{exc}）", "ERR")
            return

        try:
            with open(self.audit, "a", encoding="utf-8") as f:
                f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 准备删除 {path} | 原因: {reason}\n")
        except Exception:
            pass

        err = None
        for attempt in range(8):
            try:
                os.remove(path)
                err = None
                break
            except PermissionError as exc:
                err = exc
                if self.ctx.stop.is_set():
                    break
                time.sleep(0.4)
            except Exception as exc:
                err = exc
                break

        if err is None and not path.exists():
            self.ctx.log(
                f"删除: {path.name} → 已删除 | 校验: {reason} | 剩余磁盘空间: {free_gb(path.parent):.2f} GB", "DEL")
        else:
            self.ctx.log(
                f"删除: {path.name} → 失败（{err}），保留文件 | 剩余磁盘空间: {free_gb(path.parent):.2f} GB", "WARN")


# ---------------------------------------------------------------------------
# RAR：解压 + 逐卷回收
# ---------------------------------------------------------------------------

RAR_ERROR_RE = re.compile(
    r"crc failed|bad archive|unexpected end|broken|corrupt|"
    r"cannot find|cannot open|wrong password|encrypted file is not supported|"
    r"unknown method|access denied|write error|disk full|not enough memory|"
    r"checksum error|^error|失败|损坏|校验错误",
    re.I,
)
RE_EXTRACTING_FROM = re.compile(r"Extracting from\s+(.+?)\s*$", re.I)
RE_TESTING_FROM = re.compile(r"(?:Testing from|Testing archive)\s+(.+?)\s*$", re.I)


def rar_precheck(vset: VolumeSet, ctx: TaskContext, unrar: Path) -> bool:
    if vset.kind != "rar":
        raise ValueError("当前分卷不是 RAR 分卷，不能使用 RAR 预校验")
    if unrar is None:
        raise RuntimeError("未找到 UnRAR.exe，无法校验 RAR 分卷")

    ctx.log(f"预校验：{vset.base}，共 {vset.count} 个分卷，从第 {vset.nums[0]} 卷开始", "STEP")
    if vset.note:
        ctx.log(f"警告：{vset.note}", "WARN")
    for p in vset.volumes:
        if not p.exists() or p.stat().st_size == 0:
            ctx.log(f"预校验失败：分卷缺失或为空 —— {p.name}", "ERR")
            return False

    cmd = [str(unrar), "t", "-y", str(vset.volumes[0])]
    ctx.log("执行命令：" + " ".join(cmd), "DBG")
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            creationflags=CREATE_NO_WINDOW)
    saw_ok = False
    error = False
    tested = set()
    for raw in iter(proc.stdout.readline, b""):
        if ctx.stop.is_set():
            proc.terminate()
            raise StopRequested()
        line = decode_line(raw)
        m = RE_TESTING_FROM.search(line) or RE_EXTRACTING_FROM.search(line)
        if m:
            name = Path(m.group(1).strip()).name
            tested.add(name.lower())
            idx = next((i for i, p in enumerate(vset.volumes)
                        if p.name.lower() == name.lower()), len(tested) - 1)
            ctx.progress((idx + 1) / vset.count * 100)
            ctx.log(f"正在校验分卷 {name} | 剩余磁盘空间: {free_gb(vset.folder):.2f} GB", "STEP")
        elif "All OK" in line:
            saw_ok = True
            ctx.log("校验结果：全部分卷 CRC 校验通过（All OK）", "OK")
        elif RAR_ERROR_RE.search(line) or "failed" in line.lower():
            error = True
            ctx.log(f"校验错误：{line}", "ERR")
        elif line.strip():
            ctx.log("  " + line[:160], "DBG")
    rc = proc.wait()
    ok = (rc == 0 and saw_ok and not error)
    ctx.log(f"预校验结束：退出码 {rc}，结果={'通过' if ok else '未通过'} | "
            f"剩余磁盘空间: {free_gb(vset.folder):.2f} GB", "OK" if ok else "ERR")
    return ok


def rar_extract(vset: VolumeSet, dest: Path, password: str, ctx: TaskContext,
                unrar: Path, dm: DeletionManager) -> bool:
    if unrar is None:
        raise RuntimeError("未找到 UnRAR.exe，无法解压 RAR 分卷")
    dest.mkdir(parents=True, exist_ok=True)

    cmd = [str(unrar), "x", "-y", "-o+"]
    if password:
        cmd.append("-p" + password)
    cmd += [str(vset.volumes[0]), str(dest) + os.sep]
    ctx.log("执行命令：" + " ".join(cmd[:4] + ["<第一分卷>", "<输出目录>"]), "DBG")

    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            creationflags=CREATE_NO_WINDOW)
    current_index = 0
    saw_data = False
    all_ok = False
    line_count = 0

    ctx.log(f"正在解压 {vset.label(0)}/{vset.count}: {vset.volumes[0].name} | "
            f"剩余磁盘空间: {free_gb(dest):.2f} GB", "STEP")

    for raw in iter(proc.stdout.readline, b""):
        if ctx.stop.is_set():
            proc.terminate()
            try:
                proc.wait(timeout=8)
            except Exception:
                proc.kill()
            raise StopRequested()

        line = decode_line(raw)
        line_count += 1
        m = RE_EXTRACTING_FROM.search(line)
        if m:
            name = Path(m.group(1).strip()).name
            idx = next((i for i, p in enumerate(vset.volumes)
                        if p.name.lower() == name.lower()), None)
            if idx is None:
                ctx.log(f"校验：UnRAR 切换到未知分卷 {name}，不执行删除", "WARN")
                current_index = idx
                saw_data = False
                continue

            # 核心：进入下一卷 == 上一卷已被 UnRAR 完整读取。
            # 额外校验：进程仍存活、本卷期间有写入进度输出、此前未出现错误。
            # 注意：解压很快时 UnRAR 可能已正常退出(0)而 stdout 仍在缓冲中，
            # 因此“进程存活或已成功退出”都算通过；非 0 退出一律不删除。
            verify_ok = (proc.poll() in (None, 0) and saw_data and not ctx.error_seen
                         and idx > current_index)
            if idx > 0:
                reason = (f"已切换到 {vset.label(idx)}，{vset.label(idx - 1)} 数据已写入输出文件且无错误"
                          if verify_ok else "校验未通过，暂不删除")
                if verify_ok:
                    dm.delete_through(idx - 1, reason)
                if not dm.enabled:
                    deleted_desc = f"{vset.label(idx - 1)} → 否（只解压模式，保留）"
                elif verify_ok:
                    deleted_desc = f"{vset.label(idx - 1)} 及更早分卷 → 已删除"
                else:
                    deleted_desc = f"{vset.label(idx - 1)} → 否（校验未通过，保留）"
                ctx.log(f"正在解压 {vset.label(idx)}/{vset.count}: {vset.volumes[idx].name} | "
                        f"校验: {'通过' if verify_ok else '未通过'} | 删除: {deleted_desc} | "
                        f"剩余磁盘空间: {free_gb(dest):.2f} GB", "STEP" if verify_ok else "WARN")
            else:
                ctx.log(f"正在解压 {vset.label(idx)}/{vset.count}: {vset.volumes[idx].name} | "
                        f"删除: 否（首卷） | 剩余磁盘空间: {free_gb(dest):.2f} GB", "STEP")
            ctx.progress(min(100.0, idx / vset.count * 100))
            current_index = idx
            saw_data = False
            continue

        if "All OK" in line:
            all_ok = True
        if RAR_ERROR_RE.search(line):
            if not ctx.error_seen:
                ctx.error_seen = True
            ctx.log(f"校验: 检测到错误 → {line}", "ERR")
        if re.search(r"\d{1,3}%|\bOK\b|Extracting\s", line):
            saw_data = True
        if line.strip() and line_count % 1 == 0:
            ctx.log("  " + line[:180], "DBG")

    rc = proc.wait()
    success = (rc == 0 and all_ok and not ctx.error_seen and not ctx.stop.is_set())
    if success:
        ctx.log(f"校验: UnRAR 退出码 0 且输出 All OK，解压完整 | "
                f"剩余磁盘空间: {free_gb(dest):.2f} GB", "OK")
        dm.delete_all("整套分卷解压完成且 All OK 校验通过")
    else:
        ctx.log(f"解压未完成：退出码 {rc}，All OK={'是' if all_ok else '否'}，"
                f"错误标记={'有' if ctx.error_seen else '无'} | 未继续删除任何分卷 | "
                f"剩余磁盘空间: {free_gb(dest):.2f} GB", "ERR")
    if not dm.enabled:
        ctx.log("当前为【只解压不删文件】模式，所有分卷均已保留", "OK")
    return success


# ---------------------------------------------------------------------------
# ZIP：流式解压 + 逐卷回收
# ---------------------------------------------------------------------------

def safe_target(dest: Path, name: str) -> Path:
    name = name.replace("\\", "/")
    parts = [x for x in name.split("/") if x not in ("", ".")]
    if any(x == ".." for x in parts):
        raise ValueError(f"压缩包内存在危险路径，已拒绝：{name}")
    root = dest.resolve()
    target = root.joinpath(*parts).resolve() if parts else root
    if target != root and root not in target.parents:
        raise ValueError(f"压缩包内路径越界，已拒绝：{name}")
    return target


def zip_precheck(vset: VolumeSet, ctx: TaskContext) -> bool:
    ctx.log(f"预校验：{vset.base}，共 {vset.count} 个分卷（ZIP 流式 CRC 校验）", "STEP")
    if vset.note:
        ctx.log(f"警告：{vset.note}", "WARN")
    for p in vset.volumes:
        if not p.exists() or p.stat().st_size == 0:
            ctx.log(f"预校验失败：分卷缺失或为空 —— {p.name}", "ERR")
            return False
    stream = MultiVolumeStream(vset.volumes)
    try:
        with zipfile.ZipFile(stream) as zf:
            infos = zf.infolist()
            ctx.log(f"压缩包内共 {len(infos)} 个条目，开始逐条 CRC 校验", "STEP")
            bad = zf.testzip()
            if ctx.stop.is_set():
                raise StopRequested()
            if bad is None:
                ctx.log(f"预校验结果：全部条目 CRC 通过 | "
                        f"剩余磁盘空间: {free_gb(vset.folder):.2f} GB", "OK")
                return True
            ctx.log(f"预校验失败：CRC 不匹配 —— {bad} | "
                    f"剩余磁盘空间: {free_gb(vset.folder):.2f} GB", "ERR")
            return False
    except StopRequested:
        raise
    except Exception as exc:
        ctx.log(f"预校验失败：{exc} | 剩余磁盘空间: {free_gb(vset.folder):.2f} GB", "ERR")
        return False
    finally:
        stream.close()


def zip_extract(vset: VolumeSet, dest: Path, ctx: TaskContext, dm: DeletionManager) -> bool:
    dest.mkdir(parents=True, exist_ok=True)
    stream = MultiVolumeStream(vset.volumes)
    extracted = 0
    try:
        with zipfile.ZipFile(stream) as zf:
            infos = zf.infolist()
            ctx.log(f"压缩包内共 {len(infos)} 个条目 | "
                    f"剩余磁盘空间: {free_gb(dest):.2f} GB", "STEP")
            for i, info in enumerate(infos):
                ctx.check_stop()
                target = safe_target(dest, info.filename)
                if info.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue

                ctx.log(f"正在解压条目 {i + 1}/{len(infos)}: {info.filename} | "
                        f"剩余磁盘空间: {free_gb(dest):.2f} GB", "STEP")
                target.parent.mkdir(parents=True, exist_ok=True)
                tmp = target.with_name(target.name + ".__extracting__")
                with zf.open(info) as src, open(tmp, "wb") as out:
                    while True:
                        ctx.check_stop()
                        chunk = src.read(1024 * 1024)
                        if not chunk:
                            break
                        out.write(chunk)
                os.replace(tmp, target)   # 条目完整落盘后才改正式名
                extracted += 1
                ctx.progress((i + 1) / max(1, len(infos)) * 100)

                # 当前条目的本地头已读、数据已写入且 CRC 由 zipfile 校验通过，
                # 因此本地头之前的分卷不会再被访问，可以安全回收。
                boundary = info.header_offset
                vol_idx = bisect.bisect_left(stream.bounds, boundary)
                if vol_idx > 0 and not ctx.error_seen:
                    reason = f"条目 {info.filename} 已完整写入且 CRC 通过"
                    dm.delete_through(vol_idx - 1, reason)
                    ctx.log(f"校验: 条目 {info.filename} 写入完成 | "
                            f"删除: {'已回收处理完毕的分卷' if dm.enabled else '否（只解压模式）'} | "
                            f"剩余磁盘空间: {free_gb(dest):.2f} GB", "STEP")
        # 先关闭虚拟流（否则 Windows 上最后一个分卷仍被占用，无法删除）
        stream.close()
        ctx.log(f"校验: {extracted} 个条目全部写入并 CRC 通过 | "
                f"剩余磁盘空间: {free_gb(dest):.2f} GB", "OK")
        dm.delete_all("ZIP 分卷全部读取完成，所有条目 CRC 校验通过")
        if not dm.enabled:
            ctx.log("当前为【只解压不删文件】模式，所有分卷均已保留", "OK")
        return True
    except StopRequested:
        ctx.log("用户停止操作，未继续删除任何分卷", "WARN")
        return False
    except Exception as exc:
        ctx.error_seen = True
        ctx.log(f"解压失败：{exc} | 未继续删除任何分卷 | "
                f"剩余磁盘空间: {free_gb(dest):.2f} GB", "ERR")
        ctx.log(traceback.format_exc(limit=3), "DBG")
        return False
    finally:
        stream.close()


# ---------------------------------------------------------------------------
# 统一任务入口（GUI 与 CLI 共用）
# ---------------------------------------------------------------------------

def flatten_single_dir(dest: Path, ctx: TaskContext) -> bool:
    """智能解压：输出目录内若只有一个子目录（且无散落文件），把子目录内容上移一层。"""
    try:
        entries = [p for p in dest.iterdir()]
        dirs = [p for p in entries if p.is_dir()]
        files = [p for p in entries if not p.is_dir()]
        if len(dirs) != 1 or files:
            ctx.log("智能解压：顶层结构已是扁平或多个条目，无需剥离", "INFO")
            return True
        only = dirs[0]
        inner = list(only.iterdir())
        ctx.log(f"智能解压：检测到唯一顶层文件夹「{only.name}」，正在上移其 {len(inner)} 个条目", "STEP")
        moved = 0
        conflicts = 0
        for item in inner:
            target = dest / item.name
            if target.exists():
                conflicts += 1
                ctx.log(f"智能解压：目标已存在，跳过「{item.name}」", "WARN")
                continue
            shutil.move(str(item), str(target))
            moved += 1
        if not any(only.iterdir()):
            only.rmdir()
        ctx.log(f"智能解压完成：上移 {moved} 个条目，跳过 {conflicts} 个冲突"
                + ("（外壳文件夹已保留）" if conflicts else "（已去掉外壳文件夹）"), "OK")
        return True
    except Exception as exc:
        ctx.log(f"智能解压失败：{exc}", "ERR")
        return False


def run_task(path_input, output_dir, auto_delete: bool, password: str,
             precheck_only: bool, ctx: TaskContext,
             chooser: Optional[Callable] = None,
             smart_flatten: bool = False) -> bool:
    vset = detect_volume_set(path_input, chooser)
    ctx.log(f"识别到分卷组：{vset.base}，类型 {vset.kind.upper()}，"
            f"共 {vset.count} 卷（{vset.label(0)} … {vset.label(vset.count - 1)}）", "STEP")
    if vset.note:
        ctx.log(f"警告：{vset.note}", "WARN")
    if vset.count < 2:
        ctx.log("提示：只识别到 1 个分卷，请确认是否还有其它分卷未放入同一目录", "WARN")

    unrar = find_unrar() if vset.kind == "rar" else None
    if vset.kind == "rar" and unrar is None:
        ctx.log("未找到 UnRAR.exe：请确认 tools/UnRAR.exe 存在，或安装 WinRAR", "ERR")
        return False
    ctx.log(f"解压引擎：{unrar.name if unrar else 'Python 内置流式 ZIP 读取器'}", "DBG")

    if precheck_only:
        if vset.kind == "rar":
            return rar_precheck(vset, ctx, unrar)
        return zip_precheck(vset, ctx)

    dest = Path(output_dir).expanduser().resolve()
    allow_same = (dest == vset.folder.resolve())
    if allow_same:
        ctx.log("输出目录与分卷目录相同，已切换为安全模式：只解压不删文件（避免误删分卷）", "WARN")
        auto_delete = False
    dest.mkdir(parents=True, exist_ok=True)

    dm = DeletionManager(auto_delete, vset, ctx, dest)
    if not auto_delete:
        ctx.log("安全模式：【只解压不删文件】，不会删除任何分卷", "OK")

    if vset.kind == "rar":
        ok = rar_extract(vset, dest, password, ctx, unrar, dm)
    elif vset.kind == "zip":
        ok = zip_extract(vset, dest, ctx, dm)
    else:
        ctx.log(f"暂不支持的分卷类型：{vset.kind}。RAR 请使用 .partN.rar，ZIP 请使用 .zip.001", "ERR")
        return False

    # 智能解压：若解压结果只有一个顶层文件夹，把它内部内容上移一层，去掉多余外壳
    if ok and smart_flatten:
        ok = flatten_single_dir(dest, ctx)
    return ok

    ctx.log(f"暂不支持的分卷类型：{vset.kind}。RAR 请使用 .partN.rar，ZIP 请使用 .zip.001", "ERR")
    return False


# ---------------------------------------------------------------------------
# 图形界面
# ---------------------------------------------------------------------------

def build_gui(initial_path: str = "", quick_mode: str = ""):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("1080x760")
    root.minsize(940, 620)

    q: queue.Queue = queue.Queue()
    stop_event = threading.Event()
    worker: dict = {"thread": None}

    # ---------- 样式 ----------
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    main = ttk.Frame(root, padding=12)
    main.pack(fill="both", expand=True)

    title = ttk.Label(main, text=APP_TITLE, font=("Microsoft YaHei UI", 16, "bold"))
    title.grid(row=0, column=0, columnspan=4, sticky="w")

    warning = (
        "风险提示：自动删除模式下，一旦分卷损坏或缺失，整套文件将报废且无法重来；"
        "删除不可撤销。强烈建议先点【预校验全部分卷】并自行备份重要数据。"
    )
    warn_lbl = tk.Label(main, text=warning, fg="#a1260d", bg="#fff2e8",
                        font=("Microsoft YaHei UI", 10),
                        wraplength=1030, justify="left", padx=8, pady=6, anchor="w")
    warn_lbl.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(6, 10))

    # ---------- 输入区 ----------
    ttk.Label(main, text="分卷路径：").grid(row=2, column=0, sticky="w")
    var_path = tk.StringVar(value=initial_path or "")
    ent_path = ttk.Entry(main, textvariable=var_path, width=88)
    ent_path.grid(row=2, column=1, columnspan=2, sticky="ew", padx=(4, 4))
    path_btns = ttk.Frame(main)
    path_btns.grid(row=2, column=3, sticky="e")
    btn_folder = ttk.Button(path_btns, text="选择分卷文件夹", width=15)
    btn_folder.pack(side="left", padx=(0, 4))
    btn_first = ttk.Button(path_btns, text="选择第一个分卷", width=15)
    btn_first.pack(side="left")

    ttk.Label(main, text="输出目录：").grid(row=3, column=0, sticky="w", pady=(4, 0))
    var_out = tk.StringVar()
    ent_out = ttk.Entry(main, textvariable=var_out, width=88)
    ent_out.grid(row=3, column=1, columnspan=2, sticky="ew", padx=(4, 4), pady=(4, 0))
    btn_out = ttk.Button(main, text="选择输出目录", width=16)
    btn_out.grid(row=3, column=3, sticky="w", pady=(4, 0))

    ttk.Label(main, text="解压密码：").grid(row=4, column=0, sticky="w", pady=(4, 0))
    var_pwd = tk.StringVar()
    ent_pwd = ttk.Entry(main, textvariable=var_pwd, show="*", width=26)
    ent_pwd.grid(row=4, column=1, sticky="w", padx=(4, 4), pady=(4, 0))
    var_auto = tk.BooleanVar(value=False)
    chk_auto = ttk.Checkbutton(main, text="启用自动删除已处理分卷（默认关闭 = 只解压不删）",
                               variable=var_auto)
    chk_auto.grid(row=4, column=2, columnspan=2, sticky="w", pady=(4, 0))

    # ---------- 操作区 ----------
    bar = ttk.Frame(main)
    bar.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(12, 6))
    btn_check = ttk.Button(bar, text="预校验全部分卷", width=20)
    btn_check.pack(side="left", padx=(0, 8))
    btn_start = ttk.Button(bar, text="开始解压", width=18)
    btn_start.pack(side="left", padx=(0, 8))
    btn_stop = ttk.Button(bar, text="停止", width=12, state="disabled")
    btn_stop.pack(side="left")
    var_status = tk.StringVar(value="就绪")
    ttk.Label(bar, textvariable=var_status, font=("Microsoft YaHei UI", 10)).pack(side="right")

    # 右键菜单安装/卸载按钮（放在状态标签左侧）
    def menu_button_text() -> str:
        st = menu_install_status()
        return {"missing": "安装右键菜单", "ok": "卸载右键菜单", "moved": "修复右键菜单"}[st]

    btn_menu = ttk.Button(bar, text=menu_button_text(), width=16)

    def on_menu_click():
        st = menu_install_status()
        if st == "ok":
            if not messagebox.askyesno("卸载右键菜单",
                                       "确定移除资源管理器的右键菜单吗？\n（不影响程序本身，可随时重新安装）"):
                return
            ok = uninstall_context_menu(lambda m: emit("INFO", m))
            if ok:
                messagebox.showinfo("完成", "右键菜单已卸载。")
        else:
            tip = ("检测到右键菜单指向的是其他位置的程序，将更新为当前 exe。" if st == "moved"
                   else "安装后，右键压缩包或文件夹即可直接用本工具打开，无需预先启动软件。\n\n是否安装？")
            if st == "missing" and not messagebox.askyesno("安装右键菜单", tip):
                return
            ok = install_context_menu(lambda m: emit("INFO", m))
            if ok:
                messagebox.showinfo("完成", "右键菜单已安装！\n现在右键压缩包或文件夹，选择【用分卷解压工具打开】即可。")
        btn_menu.configure(text=menu_button_text())

    btn_menu.configure(command=on_menu_click)
    btn_menu.pack(side="right", padx=(0, 12))

    var_progress = tk.DoubleVar(value=0)
    pb = ttk.Progressbar(main, variable=var_progress, maximum=100, mode="determinate")
    pb.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(2, 8))

    # ---------- 日志区 ----------
    log_box = scrolledtext.ScrolledText(main, height=24, wrap="word",
                                        font=("Consolas", 10))
    log_box.grid(row=7, column=0, columnspan=4, sticky="nsew")
    log_box.configure(state="disabled")
    log_box.tag_config("INFO", foreground="#24292f")
    log_box.tag_config("DBG", foreground="#7b7b7b")
    log_box.tag_config("STEP", foreground="#0b5fa5")
    log_box.tag_config("OK", foreground="#127a24")
    log_box.tag_config("DEL", foreground="#a1260d")
    log_box.tag_config("WARN", foreground="#9a6700")
    log_box.tag_config("ERR", foreground="#c02020")

    main.rowconfigure(7, weight=1)
    main.columnconfigure(1, weight=1)
    main.columnconfigure(2, weight=1)

    note = ("说明：RAR 分卷使用内置 UnRAR.exe（RARLAB 免费命令行版）；ZIP 分卷使用 Python 流式读取器。"
            "删除记录写入 logs/删除审计_日期.log。日志中的“剩余磁盘空间”按输出/分卷所在磁盘实时读取。\n"
            f"{SIGNATURE}")
    ttk.Label(main, text=note, foreground="#666666", font=("Microsoft YaHei UI", 9),
              wraplength=1030, justify="left").grid(row=8, column=0, columnspan=4, sticky="w", pady=(6, 0))

    # ---------- 日志队列 ----------
    def emit(level: str, msg: str) -> None:
        q.put((level, msg))

    def drain_queue() -> None:
        try:
            while True:
                level, msg = q.get_nowait()
                if level == "__progress__":
                    var_progress.set(max(0.0, min(100.0, float(msg))))
                    continue
                if level == "__done__":
                    set_busy(False)
                    if msg is None:
                        var_status.set("已停止")
                        emit("WARN", "任务已停止（未继续删除分卷）")
                    elif msg:
                        var_status.set("完成")
                        var_progress.set(100)
                        emit("OK", "任务完成")
                    else:
                        var_status.set("失败/未通过")
                        emit("ERR", "任务未通过，请检查日志")
                    continue
                log_box.configure(state="normal")
                log_box.insert("end", f"[{now_str()}] {msg}\n", level if level in
                               ("DBG", "STEP", "OK", "DEL", "WARN", "ERR") else "INFO")
                log_box.see("end")
                log_box.configure(state="disabled")
        except queue.Empty:
            pass
        root.after(120, drain_queue)

    def set_busy(busy: bool) -> None:
        for b in (btn_check, btn_start, btn_folder, btn_first, btn_out, chk_auto, ent_pwd):
            b.configure(state="disabled" if busy else "normal")
        btn_stop.configure(state="normal" if busy else "disabled")

    drain_queue()

    def log_choice(sets):
        result = {}

        def pick():
            result["v"] = lb.get(lb.curselection()[0])
            dlg.destroy()

        dlg = tk.Toplevel(root)
        dlg.title("选择分卷组")
        dlg.geometry("560x320")
        ttk.Label(dlg, text="检测到多组分卷，请选择要处理的一组：").pack(anchor="w", padx=12, pady=10)
        lb = tk.Listbox(dlg, font=("Microsoft YaHei UI", 10))
        for s in sets:
            lb.insert("end", f"{s.base}  ({s.count} 卷, {s.kind}, 从 {s.label(0)} 开始)")
        lb.pack(fill="both", expand=True, padx=12)
        if sets:
            lb.selection_set(0)
        ttk.Button(dlg, text="确定", command=pick).pack(pady=10)
        dlg.wait_window()
        return sets[result["v"]] if "v" in result else sets[0]

    def confirm_risky() -> bool:
        """自动删除开启时的强制二次确认：必须勾选才允许继续"""
        result = {"ok": False}
        dlg = tk.Toplevel(root)
        dlg.title("二次确认：启用自动删除")
        dlg.geometry("640x380")
        dlg.resizable(False, False)
        dlg.transient(root)
        dlg.grab_set()
        txt = (
            "你已勾选【启用自动删除已处理分卷】。\n\n"
            "• 程序会在确认后续分卷已开始读取、且当前分卷数据已写入输出文件后，删除已处理分卷。\n"
            "• 一旦某个分卷损坏或缺失，整套分卷将报废，且已删除的分卷不可恢复。\n"
            "• 本程序只删除同组中已处理完成的分卷，不会删除未解压分卷，也不会删除解压出来的成品文件。\n\n"
            "强烈建议：先点【预校验全部分卷】，并确认已有备份或可以承受数据丢失风险。"
        )
        tk.Label(dlg, text=txt, justify="left", anchor="w", wraplength=600,
                 font=("Microsoft YaHei UI", 10), fg="#a1260d").pack(padx=16, pady=14, anchor="w")
        var_ack = tk.BooleanVar(value=False)
        btn_ok = {}

        def toggle():
            btn_ok["b"].configure(state="normal" if var_ack.get() else "disabled")

        ttk.Checkbutton(dlg, text="我已知晓风险，确认启用自动删除", variable=var_ack,
                        command=toggle).pack(anchor="w", padx=16)
        btns = ttk.Frame(dlg)
        btns.pack(pady=16)
        btn_ok["b"] = ttk.Button(btns, text="确认并开始", state="disabled",
                                 command=lambda: (result.update(ok=True), dlg.destroy()))
        btn_ok["b"].pack(side="left", padx=6)
        ttk.Button(btns, text="取消", command=dlg.destroy).pack(side="left", padx=6)
        dlg.wait_window()
        return result["ok"]

    def on_auto_toggle():
        if var_auto.get():
            messagebox.showwarning(
                "风险提示",
                "自动删除模式会在解压过程中永久删除已处理分卷。\n"
                "若分卷损坏/缺失，整套文件将报废且无法恢复。\n"
                "建议先执行【预校验全部分卷】。开始解压时还会再次弹窗确认。",
                parent=root)

    chk_auto.configure(command=on_auto_toggle)

    def choose_folder():
        d = filedialog.askdirectory(parent=root, title="选择分卷所在文件夹")
        if d:
            var_path.set(d)
            if not var_out.get():
                var_out.set(os.path.join(d, "解压结果"))
            emit("INFO", f"已选择分卷文件夹：{d}")

    def choose_first():
        f = filedialog.askopenfilename(
            parent=root, title="选择第一个分卷",
            filetypes=[("分卷压缩包", "*.rar *.zip *.001 *.002 *.003 *.r00 *.r01 *.z01 *.part1.*"),
                       ("所有文件", "*.*")])
        if f:
            var_path.set(f)
            if not var_out.get():
                var_out.set(os.path.join(os.path.dirname(f), "解压结果"))
            emit("INFO", f"已选择第一个分卷：{f}")

    def choose_out():
        d = filedialog.askdirectory(parent=root, title="选择输出目录")
        if d:
            var_out.set(d)

    btn_folder.configure(command=choose_folder)
    btn_first.configure(command=choose_first)
    btn_out.configure(command=choose_out)

    def start_worker(precheck: bool, smart: bool = False) -> None:
        path_input = var_path.get().strip()
        output = var_out.get().strip()
        if not path_input:
            messagebox.showwarning("缺少参数", "请先选择分卷文件夹或第一个分卷", parent=root)
            return
        if not precheck and not output:
            messagebox.showwarning("缺少参数", "请选择输出目录", parent=root)
            return
        if not precheck and var_auto.get() and not confirm_risky():
            emit("WARN", "用户取消：未启用自动删除，任务未开始")
            return

        stop_event.clear()
        auto = var_auto.get() if not precheck else False
        pwd = var_pwd.get()
        ctx = TaskContext(emit, stop_event)
        var_progress.set(0)
        set_busy(True)
        var_status.set("正在校验…" if precheck else "正在解压…")
        emit("STEP", "=" * 30 + ("预校验" if precheck else "开始解压") + "=" * 30)

        def work():
            try:
                ok = run_task(path_input, output, auto, pwd, precheck, ctx, log_choice,
                              smart_flatten=smart)
                q.put(("__done__", ok))
            except StopRequested:
                q.put(("__done__", None))
            except Exception as exc:
                emit("ERR", f"任务异常：{exc}")
                emit("DBG", traceback.format_exc())
                q.put(("__done__", False))

        th = threading.Thread(target=work, daemon=True)
        worker["thread"] = th
        th.start()

    def do_check():
        start_worker(True)

    def do_start():
        start_worker(False)

    def do_stop():
        stop_event.set()
        emit("WARN", "已发送停止请求，正在安全中止…")
        var_status.set("正在停止…")

    btn_check.configure(command=do_check)
    btn_start.configure(command=do_start)
    btn_stop.configure(command=do_stop)

    emit("INFO", "就绪。RAR 分卷推荐使用 .part1.rar 命名；ZIP 分卷推荐 .zip.001 命名。")

    # 右键快捷模式：启动时自动填好路径/输出目录并开始解压（默认只解压不删，安全）
    if quick_mode:
        def apply_quick():
            p = initial_path.strip()
            if not p:
                return
            var_path.set(p)
            is_dir = os.path.isdir(p)
            parent = p if is_dir else os.path.dirname(p)
            if quick_mode == "here":
                out = parent
                smart = False
            elif quick_mode == "smart":
                out = parent
                smart = True
            elif quick_mode == "subfolder":
                base = os.path.basename(p.rstrip("\\/"))
                if not is_dir:
                    # 去掉分卷后缀，取归档基名
                    for pat in (r"\.part\d+\.[^.]+$", r"\.(zip|rar)\.\d+$", r"\.\d+$", r"\.r\d{2}$", r"\.z\d{2}$"):
                        base = re.sub(pat, "", base, flags=re.I)
                        if base != os.path.basename(p.rstrip("\\/")):
                            break
                    if base.lower().endswith(".zip") or base.lower().endswith(".rar"):
                        base = os.path.splitext(base)[0]
                out = os.path.join(parent, base)
                smart = False
            elif quick_mode == "browse":
                chosen = filedialog.askdirectory(parent=root, title="选择解压到的目标文件夹")
                if not chosen:
                    emit("WARN", "已取消选择目标文件夹，未开始解压")
                    return
                out = chosen
                smart = False
            else:
                out = os.path.join(parent, "解压结果")
                smart = False
            var_out.set(out)
            emit("STEP", f"右键快捷解压（{quick_mode}）：{p} → {out}")
            start_worker(False, smart=smart)
        root.after(300, apply_quick)

    return root


# ---------------------------------------------------------------------------
# 命令行模式
# ---------------------------------------------------------------------------

def cli_main(argv: List[str]) -> int:
    import argparse
    ap = argparse.ArgumentParser(description="分卷压缩包分批解压工具（命令行模式）")
    ap.add_argument("volume", help="第一个分卷文件，或分卷所在文件夹")
    ap.add_argument("--output", default="", help="输出目录（预校验时可省略）")
    ap.add_argument("--precheck", action="store_true", help="只预校验全部分卷，不解压")
    ap.add_argument("--auto-delete", action="store_true", help="启用自动删除已处理分卷（默认关闭）")
    ap.add_argument("--password", default="", help="解压密码")
    ap.add_argument("--log", default="", help="把日志同时写入文件（便于自动化测试和留档）")
    ap.add_argument("--smart", action="store_true", help="智能解压：剥掉唯一顶层文件夹")
    args = ap.parse_args(argv)

    log_fp = None
    if args.log:
        log_path = Path(args.log).expanduser().resolve()
        log_path.parent.mkdir(parents=True, exist_ok=True)
        log_fp = open(log_path, "w", encoding="utf-8")

    def emit(level, msg):
        if level.startswith("__"):
            return
        line = f"[{now_str()}] [{level}] {msg}"
        print(line, flush=True)
        if log_fp:
            log_fp.write(line + "\n")
            log_fp.flush()

    try:
        try:
            ctx = TaskContext(emit, threading.Event())
            ok = run_task(args.volume, args.output or str(Path(args.volume).parent / "解压结果"),
                          args.auto_delete, args.password, args.precheck, ctx,
                          smart_flatten=args.smart)
            return 0 if ok else 2
        except Exception as exc:
            emit("ERR", f"致命错误：{exc}")
            return 2
    finally:
        if log_fp:
            log_fp.close()


def main() -> int:
    prepare_console()
    try:
        argv = sys.argv[1:]
        if "--install-menu" in argv:
            return 0 if install_context_menu() else 1
        if "--uninstall-menu" in argv:
            return 0 if uninstall_context_menu() else 1
        if "--cli" in argv:
            return cli_main([a for a in argv if a != "--cli"])
        initial_path = ""
        quick_mode = ""
        if "--path" in argv:
            i = argv.index("--path")
            if i + 1 < len(argv):
                initial_path = argv[i + 1]
        if "--quick" in argv:
            i = argv.index("--quick")
            if i + 2 < len(argv):
                quick_mode = argv[i + 1]
                initial_path = argv[i + 2]
        build_gui(initial_path, quick_mode).mainloop()
        return 0
    except Exception as exc:
        # 防止无控制台 exe 弹出 PyInstaller 错误框并卡死
        try:
            print(f"[FATAL] {exc}", flush=True)
        except Exception:
            pass
        return 2


if __name__ == "__main__":
    sys.exit(main())
