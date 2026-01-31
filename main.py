from __future__ import annotations

import sys
if sys.platform != 'win32':
    raise RuntimeError("This script requires Windows.")

"""
Stateless vision-driven Windows UI agent with a single visual truth.
The only graphics emitted by the agent are drawn into a transparent topmost layered window (overlay).
Every screenshot sent to the model is composed in software as: desktop pixels + OS cursor + overlay pixels (alpha blend).
The exact PNG bytes saved to disk are the same bytes sent to the model.
No post-processing draws into screenshots after capture; if it is visible in a saved frame, it was visible on screen.
Persistent UI labels are stored in annotations.json and re-rendered on every frame.
Model output is one JSON action object per step; tools are executed via Win32 input injection.
Test mode uses the same executor and capture pipeline and provides per-field defaults on ENTER.
"""

import base64
import ctypes
import ctypes.wintypes as w
import json
import re
import struct
import time
import urllib.request
import zlib
from dataclasses import dataclass, field
from datetime import datetime
from enum import IntFlag
from functools import cache
from pathlib import Path
from typing import Any, Callable, Literal, cast


MODEL_NAME = "qwen3-vl-4b-instruct-1m"
API_URL = "http://localhost:1234/v1/chat/completions"
REQUEST_TIMEOUT_S = 60

MAX_API_RETRIES = 3
BACKOFF_FACTOR = 2.0

MAX_STEPS = 60
MAX_RETRIES = 3
RETRY_DELAY_S = 1.5
MAX_CONSECUTIVE_FAILURES = 4

INPUT_DELAY_S = 0.05

SCREENSHOT_QUALITY = 2
SCREEN_W, SCREEN_H = {1: (1536, 864), 2: (1024, 576), 3: (512, 288)}[SCREENSHOT_QUALITY]

DELAY_AFTER_ACTION_S = 0.10
DELAY_MOVE_HOVER_S = 1.50
DELAY_SCROLL_S = 0.15

CROSS_SIZE = 30
LINE_THICKNESS = 3
ARROW_SIZE = 18

HUD_ENABLED = True
HUD_MAX_CHARS = 180
HUD_MARGIN = 6
HUD_HEIGHT = 24
HUD_MAX_WIDTH = 960
OVERLAY_REASSERT_PULSES = 2
OVERLAY_REASSERT_PAUSE_S = 0.03

DEFAULT_TASK = (
    "Open Microsoft Paint from the Start menu then use the mouse to draw a simple cat face "
    "with two circles for eyes one triangle for nose and curved line for smile then save the "
    "file as cat in the Pictures folder and close Paint when done"
)

ActionTool = Literal["click", "move", "drag", "type", "scroll", "annotate", "recall", "done"]


@cache
def _dll(name: str) -> ctypes.WinDLL:
    return ctypes.WinDLL(name, use_last_error=True)


user32 = _dll("user32")
gdi32 = _dll("gdi32")
kernel32 = _dll("kernel32")


def _set_dpi_awareness() -> None:
    try:
        ctypes.WinDLL("Shcore", use_last_error=True).SetProcessDpiAwareness(2)
    except Exception:
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass


_set_dpi_awareness()


class MouseEvent(IntFlag):
    MOVE = 0x0001
    ABSOLUTE = 0x8000
    LEFT_DOWN = 0x0002
    LEFT_UP = 0x0004
    WHEEL = 0x0800
    HWHEEL = 0x1000


class KeyEvent(IntFlag):
    KEYUP = 0x0002
    UNICODE = 0x0004


class WinStyle(IntFlag):
    EX_TOPMOST = 0x00000008
    EX_LAYERED = 0x00080000
    EX_TRANSPARENT = 0x00000020
    EX_NOACTIVATE = 0x08000000
    EX_TOOLWINDOW = 0x00000080
    POPUP = 0x80000000


INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
WHEEL_DELTA = 120
SRCCOPY = 0x00CC0020

SW_SHOWNOACTIVATE = 4

ULW_ALPHA = 2
AC_SRC_ALPHA = 1

SWP_NOSIZE = 1
SWP_NOMOVE = 2
SWP_NOACTIVATE = 16
SWP_SHOWWINDOW = 64
HWND_TOPMOST = -1

COLOR_RED = 0xFFFF0000
COLOR_GREEN = 0xFF00FF00
COLOR_YELLOW = 0xFFFFFF00
COLOR_MAGENTA = 0xFFFF00FF
COLOR_BLACK = 0xFF000000
ANNOTATE_COLOR = COLOR_YELLOW

DI_NORMAL = 0x0003
CURSOR_SHOWING = 0x00000001

TRANSPARENT = 1

DEFAULT_CHARSET = 1
OUT_DEFAULT_PRECIS = 0
CLIP_DEFAULT_PRECIS = 0
DEFAULT_QUALITY = 0
DEFAULT_PITCH = 0


LRESULT = ctypes.c_ssize_t
WPARAM = ctypes.c_size_t
LPARAM = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, w.HWND, w.UINT, WPARAM, LPARAM)
ULONG_PTR = ctypes.c_size_t


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", w.LONG),
        ("dy", w.LONG),
        ("mouseData", w.DWORD),
        ("dwFlags", w.DWORD),
        ("time", w.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", w.WORD),
        ("wScan", w.WORD),
        ("dwFlags", w.DWORD),
        ("time", w.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [("uMsg", w.DWORD), ("wParamL", w.WORD), ("wParamH", w.WORD)]


class _INPUTunion(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]


class INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [("type", w.DWORD), ("u", _INPUTunion)]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", w.DWORD),
        ("biWidth", w.LONG),
        ("biHeight", w.LONG),
        ("biPlanes", w.WORD),
        ("biBitCount", w.WORD),
        ("biCompression", w.DWORD),
        ("biSizeImage", w.DWORD),
        ("biXPelsPerMeter", w.LONG),
        ("biYPelsPerMeter", w.LONG),
        ("biClrUsed", w.DWORD),
        ("biClrImportant", w.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", ctypes.c_uint * 1)]


class CURSORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", w.DWORD),
        ("flags", w.DWORD),
        ("hCursor", w.HANDLE),
        ("ptScreenPos", w.POINT),
    ]


class ICONINFO(ctypes.Structure):
    _fields_ = [
        ("fIcon", w.BOOL),
        ("xHotspot", w.DWORD),
        ("yHotspot", w.DWORD),
        ("hbmMask", w.HBITMAP),
        ("hbmColor", w.HBITMAP),
    ]


class BLENDFUNCTION(ctypes.Structure):
    _fields_ = [
        ("BlendOp", ctypes.c_ubyte),
        ("BlendFlags", ctypes.c_ubyte),
        ("SourceConstantAlpha", ctypes.c_ubyte),
        ("AlphaFormat", ctypes.c_ubyte),
    ]


class WNDCLASS(ctypes.Structure):
    _fields_ = [
        ("style", ctypes.c_uint),
        ("lpfnWndProc", ctypes.c_void_p),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", w.HINSTANCE),
        ("hIcon", w.HANDLE),
        ("hCursor", w.HANDLE),
        ("hbrBackground", w.HANDLE),
        ("lpszMenuName", w.LPCWSTR),
        ("lpszClassName", w.LPCWSTR),
    ]


user32.DefWindowProcW.argtypes = [w.HWND, w.UINT, WPARAM, LPARAM]
user32.DefWindowProcW.restype = LRESULT

_SendInput = user32.SendInput
_SendInput.argtypes = (w.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
_SendInput.restype = w.UINT

gdi32.TextOutW.argtypes = [w.HDC, ctypes.c_int, ctypes.c_int, w.LPCWSTR, ctypes.c_int]
gdi32.TextOutW.restype = w.BOOL


# Win32 signatures (64-bit safe)
gdi32.CreateCompatibleDC.argtypes = [w.HDC]
gdi32.CreateCompatibleDC.restype = w.HDC
gdi32.CreateDIBSection.argtypes = [w.HDC, ctypes.POINTER(BITMAPINFO), w.UINT, ctypes.POINTER(ctypes.c_void_p), w.HANDLE, w.DWORD]
gdi32.CreateDIBSection.restype = w.HBITMAP
gdi32.SelectObject.argtypes = [w.HDC, w.HGDIOBJ]
gdi32.SelectObject.restype = w.HGDIOBJ
gdi32.BitBlt.argtypes = [w.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, w.HDC, ctypes.c_int, ctypes.c_int, w.DWORD]
gdi32.BitBlt.restype = w.BOOL
gdi32.DeleteObject.argtypes = [w.HGDIOBJ]
gdi32.DeleteObject.restype = w.BOOL
gdi32.DeleteDC.argtypes = [w.HDC]
gdi32.DeleteDC.restype = w.BOOL
gdi32.SetBkMode.argtypes = [w.HDC, ctypes.c_int]
gdi32.SetBkMode.restype = ctypes.c_int
gdi32.SetTextColor.argtypes = [w.HDC, w.DWORD]
gdi32.SetTextColor.restype = w.DWORD
gdi32.CreateFontW.restype = w.HFONT

user32.ReleaseDC.argtypes = [w.HWND, w.HDC]
user32.ReleaseDC.restype = ctypes.c_int
user32.GetCursorInfo.argtypes = [ctypes.POINTER(CURSORINFO)]
user32.GetCursorInfo.restype = w.BOOL
user32.GetIconInfo.argtypes = [w.HICON, ctypes.POINTER(ICONINFO)]
user32.GetIconInfo.restype = w.BOOL
user32.DrawIconEx.argtypes = [w.HDC, ctypes.c_int, ctypes.c_int, w.HICON, ctypes.c_int, ctypes.c_int, w.UINT, w.HBRUSH, w.UINT]
user32.DrawIconEx.restype = w.BOOL

user32.UpdateLayeredWindow.argtypes = [
    w.HWND,
    w.HDC,
    ctypes.POINTER(w.POINT),
    ctypes.POINTER(w.SIZE),
    w.HDC,
    ctypes.POINTER(w.POINT),
    w.DWORD,
    ctypes.POINTER(BLENDFUNCTION),
    w.DWORD
]
user32.UpdateLayeredWindow.restype = w.BOOL

user32.SetWindowPos.argtypes = [
    w.HWND,
    w.HWND,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    w.UINT
]
user32.SetWindowPos.restype = w.BOOL



def get_screen_size() -> tuple[int, int]:
    return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)


@dataclass(slots=True)
class CoordConverter:
    sw: int
    sh: int
    mw: int
    mh: int

    def norm_to_screen(self, xn: float, yn: float) -> tuple[int, int]:
        return int(xn * self.sw / 1000), int(yn * self.sh / 1000)

    def to_win32_normalized(self, x: int, y: int) -> tuple[int, int]:
        return (
            max(0, min(65535, int(x * 65535 / max(1, self.sw - 1)))),
            max(0, min(65535, int(y * 65535 / max(1, self.sh - 1)))),
        )


def _send_input(inputs: list[INPUT], *, delay_s: float = INPUT_DELAY_S) -> None:
    arr = (INPUT * len(inputs))(*inputs)
    if _SendInput(len(inputs), arr, ctypes.sizeof(INPUT)) != len(inputs):
        raise ctypes.WinError(ctypes.get_last_error())
    if delay_s > 0:
        time.sleep(delay_s)


def _get_cursor_pos() -> tuple[int, int]:
    pt = w.POINT()
    if user32.GetCursorPos(ctypes.byref(pt)):
        return int(pt.x), int(pt.y)
    return (0, 0)


def mouse_move(x: int, y: int, conv: CoordConverter) -> None:
    ax, ay = conv.to_win32_normalized(x, y)
    i = INPUT(type=INPUT_MOUSE)
    i.mi = MOUSEINPUT(ax, ay, 0, MouseEvent.MOVE | MouseEvent.ABSOLUTE, 0, 0)
    _send_input([i])


def mouse_click(x: int, y: int, conv: CoordConverter) -> None:
    ax, ay = conv.to_win32_normalized(x, y)
    seq: list[tuple[int, int]] = [(MouseEvent.MOVE, 0), (MouseEvent.LEFT_DOWN, 0), (MouseEvent.LEFT_UP, 0)]
    inputs: list[INPUT] = []
    for flag, data in seq:
        i = INPUT(type=INPUT_MOUSE)
        i.mi = MOUSEINPUT(ax, ay, data, int(flag) | int(MouseEvent.ABSOLUTE), 0, 0)
        inputs.append(i)
    _send_input(inputs)


def mouse_drag(x1: int, y1: int, x2: int, y2: int, conv: CoordConverter, *, steps: int = 14, step_pause_s: float = 0.01) -> None:
    ax1, ay1 = conv.to_win32_normalized(x1, y1)
    ax2, ay2 = conv.to_win32_normalized(x2, y2)

    def send(flags: int, dx: int, dy: int, *, delay: float = INPUT_DELAY_S) -> None:
        inp = INPUT(type=INPUT_MOUSE)
        inp.mi = MOUSEINPUT(dx, dy, 0, flags, 0, 0)
        _send_input([inp], delay_s=delay)

    send(int(MouseEvent.MOVE | MouseEvent.ABSOLUTE), ax1, ay1)
    send(int(MouseEvent.LEFT_DOWN | MouseEvent.ABSOLUTE), ax1, ay1)

    steps = max(1, int(steps))
    pause = max(0.0, float(step_pause_s))
    for k in range(1, steps + 1):
        t = k / steps
        dx = int(ax1 + (ax2 - ax1) * t)
        dy = int(ay1 + (ay2 - ay1) * t)
        send(int(MouseEvent.MOVE | MouseEvent.ABSOLUTE), dx, dy, delay=0.0)
        if pause:
            time.sleep(pause)

    send(int(MouseEvent.LEFT_UP | MouseEvent.ABSOLUTE), ax2, ay2)



def type_text(text: str) -> None:
    if not text:
        return
    inputs: list[INPUT] = []
    for ch in text:
        b = ch.encode("utf-16le")
        for i in range(0, len(b), 2):
            cu = b[i] | (b[i + 1] << 8)
            for flags in (KeyEvent.UNICODE, KeyEvent.UNICODE | KeyEvent.KEYUP):
                inp = INPUT(type=INPUT_KEYBOARD)
                inp.ki = KEYBDINPUT(0, cu, int(flags), 0, 0)
                inputs.append(inp)
    _send_input(inputs)


def scroll(dx: float = 0.0, dy: float = 0.0) -> None:
    inputs: list[INPUT] = []
    for delta, flag in ((dy, MouseEvent.WHEEL), (dx, MouseEvent.HWHEEL)):
        if not delta:
            continue
        ticks = max(1, abs(int(delta)) // 100)
        direction = 1 if delta > 0 else -1
        for _ in range(ticks):
            inp = INPUT(type=INPUT_MOUSE)
            inp.mi = MOUSEINPUT(0, 0, WHEEL_DELTA * direction, int(flag), 0, 0)
            inputs.append(inp)
    if inputs:
        _send_input(inputs)


def _capture_desktop_bgra(sw: int, sh: int, *, include_cursor: bool) -> bytes:
    sdc = user32.GetDC(0)
    mdc = gdi32.CreateCompatibleDC(sdc)

    bmi = BITMAPINFO()
    ctypes.memset(ctypes.byref(bmi), 0, ctypes.sizeof(bmi))
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = sw
    bmi.bmiHeader.biHeight = -sh
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    bmi.bmiHeader.biCompression = 0

    bits = ctypes.c_void_p()
    hbm = gdi32.CreateDIBSection(sdc, ctypes.byref(bmi), 0, ctypes.byref(bits), None, 0)
    if not hbm:
        raise ctypes.WinError(ctypes.get_last_error())
    gdi32.SelectObject(mdc, hbm)

    if not gdi32.BitBlt(mdc, 0, 0, sw, sh, sdc, 0, 0, SRCCOPY):
        raise ctypes.WinError(ctypes.get_last_error())

    if include_cursor:
        ci = CURSORINFO(cbSize=ctypes.sizeof(CURSORINFO))
        if user32.GetCursorInfo(ctypes.byref(ci)) and (ci.flags & CURSOR_SHOWING):
            ii = ICONINFO()
            if user32.GetIconInfo(ci.hCursor, ctypes.byref(ii)):
                x = int(ci.ptScreenPos.x) - int(ii.xHotspot)
                y = int(ci.ptScreenPos.y) - int(ii.yHotspot)
                user32.DrawIconEx(mdc, x, y, ci.hCursor, 0, 0, 0, 0, DI_NORMAL)
                if ii.hbmMask:
                    gdi32.DeleteObject(ii.hbmMask)
                if ii.hbmColor:
                    gdi32.DeleteObject(ii.hbmColor)

    assert bits.value is not None
    out = ctypes.string_at(cast(int, bits.value), sw * sh * 4)

    user32.ReleaseDC(0, sdc)
    gdi32.DeleteDC(mdc)
    gdi32.DeleteObject(hbm)
    return out


@cache
def _nn_maps(sw: int, sh: int, dw: int, dh: int) -> tuple[list[int], list[int]]:
    xm = [((x * sw) // dw) * 4 for x in range(dw)]
    ym = [(y * sh) // dh for y in range(dh)]
    return xm, ym


def _downsample_nn_bgra(src: bytes, sw: int, sh: int, dw: int, dh: int) -> bytes:
    if (sw, sh) == (dw, dh):
        return src
    xm, ym = _nn_maps(sw, sh, dw, dh)
    src_mv = memoryview(src)
    dst = bytearray(dw * dh * 4)
    row_bytes = sw * 4
    for y, sy in enumerate(ym):
        srow = src_mv[sy * row_bytes : (sy + 1) * row_bytes]
        base = y * dw * 4
        for x, sx4 in enumerate(xm):
            di = base + x * 4
            dst[di : di + 4] = srow[sx4 : sx4 + 4]
    return bytes(dst)


def _alpha_blend_bgra(base: bytes, overlay: bytes) -> bytes:
    out = bytearray(base)
    ov = memoryview(overlay)
    for i in range(0, len(out), 4):
        oa = ov[i + 3]
        if not oa:
            continue
        inv = 255 - oa
        out[i + 0] = (ov[i + 0] * oa + out[i + 0] * inv + 127) // 255
        out[i + 1] = (ov[i + 1] * oa + out[i + 1] * inv + 127) // 255
        out[i + 2] = (ov[i + 2] * oa + out[i + 2] * inv + 127) // 255
    return bytes(out)


def _encode_png_rgb(bgra: bytes, width: int, height: int) -> bytes:
    raw = bytearray((width * 3 + 1) * height)
    stride_src = width * 4
    stride_dst = width * 3 + 1
    for y in range(height):
        raw[y * stride_dst] = 0
        row = bgra[y * stride_src : (y + 1) * stride_src]
        di = y * stride_dst + 1
        raw[di : di + width * 3 : 3] = row[2::4]
        raw[di + 1 : di + width * 3 : 3] = row[1::4]
        raw[di + 2 : di + width * 3 : 3] = row[0::4]
    comp = zlib.compress(bytes(raw))
    ihdr = struct.pack(">2I5B", width, height, 8, 2, 0, 0, 0)

    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", comp) + chunk(b"IEND", b"")


def _wndproc(hwnd: int, msg: int, wparam: int, lparam: int) -> int:
    return user32.DefWindowProcW(hwnd, msg, wparam, lparam)


_wndproc_cb: WNDPROC = WNDPROC(_wndproc)


class OverlayManager:
    def __init__(self, sw: int, sh: int) -> None:
        self.w, self.h = sw, sh
        self.hwnd: w.HWND | None = None
        self.hdc: w.HDC | None = None
        self.hbitmap: w.HBITMAP | None = None
        self.bits: ctypes.c_void_p | None = None
        self._font: w.HFONT | None = None
        self._action: dict[str, Any] | None = None
        self._annotations: dict[str, dict[str, Any]] = {}
        self._highlight_label: str | None = None
        self._highlight_color: int = COLOR_MAGENTA

    def __enter__(self) -> "OverlayManager":
        self._init()
        return self

    def __exit__(self, *_: Any) -> None:
        self.close()

    def _init(self) -> None:
        if self.hwnd:
            return
        hinst = kernel32.GetModuleHandleW(None)
        cls_name = "AIAgentOverlayWindow"

        wc = WNDCLASS()
        wc.style = 0
        wc.lpfnWndProc = ctypes.cast(_wndproc_cb, ctypes.c_void_p)
        wc.cbClsExtra = 0
        wc.cbWndExtra = 0
        wc.hInstance = hinst
        wc.hIcon = None
        wc.hCursor = user32.LoadCursorW(None, 32512)
        wc.hbrBackground = None
        wc.lpszMenuName = None
        wc.lpszClassName = cls_name

        atom = user32.RegisterClassW(ctypes.byref(wc))
        if atom == 0 and ctypes.get_last_error() != 1410:
            raise ctypes.WinError(ctypes.get_last_error())

        ex = (
            WinStyle.EX_LAYERED
            | WinStyle.EX_TRANSPARENT
            | WinStyle.EX_TOPMOST
            | WinStyle.EX_NOACTIVATE
            | WinStyle.EX_TOOLWINDOW
        )
        self.hwnd = user32.CreateWindowExW(
            int(ex),
            cls_name,
            "AI Overlay",
            int(WinStyle.POPUP),
            0,
            0,
            self.w,
            self.h,
            0,
            0,
            hinst,
            None,
        )
        if not self.hwnd:
            raise ctypes.WinError(ctypes.get_last_error())

        self.hdc = gdi32.CreateCompatibleDC(0)

        bmi = BITMAPINFO()
        ctypes.memset(ctypes.byref(bmi), 0, ctypes.sizeof(bmi))
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = self.w
        bmi.bmiHeader.biHeight = -self.h
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0

        bits = ctypes.c_void_p()
        self.hbitmap = gdi32.CreateDIBSection(0, ctypes.byref(bmi), 0, ctypes.byref(bits), None, 0)
        if not self.hbitmap or not bits:
            self.close()
            raise ctypes.WinError(ctypes.get_last_error())

        self.bits = bits
        gdi32.SelectObject(self.hdc, self.hbitmap)
        ctypes.memset(self.bits, 0, self.w * self.h * 4)

        self._font = self._create_font()
        if self._font:
            gdi32.SelectObject(self.hdc, self._font)
        gdi32.SetBkMode(self.hdc, TRANSPARENT)
        gdi32.SetTextColor(self.hdc, 0x00FFFFFF)

        p = self._p()
        self._draw_hud(p)
        self._refresh()
        user32.ShowWindow(self.hwnd, SW_SHOWNOACTIVATE)
        self.reassert_topmost()

    def _create_font(self) -> w.HFONT | None:
        try:
            return cast(
                w.HFONT,
                gdi32.CreateFontW(
                    -16,
                    0,
                    0,
                    0,
                    700,
                    0,
                    0,
                    0,
                    DEFAULT_CHARSET,
                    OUT_DEFAULT_PRECIS,
                    CLIP_DEFAULT_PRECIS,
                    DEFAULT_QUALITY,
                    DEFAULT_PITCH,
                    "Segoe UI",
                ),
            )
        except Exception:
            return None

    def bring_to_front(self) -> None:
        if not self.hwnd:
            return
        if not user32.SetWindowPos(
            self.hwnd,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW,
        ):
            raise ctypes.WinError(ctypes.get_last_error())

    def reassert_topmost(self, pulses: int = OVERLAY_REASSERT_PULSES, pause_s: float = OVERLAY_REASSERT_PAUSE_S) -> None:
        if not self.hwnd:
            return
        pulses = max(1, int(pulses))
        pause = max(0.0, float(pause_s))
        for i in range(pulses):
            self.bring_to_front()
            if i + 1 < pulses and pause:
                time.sleep(pause)

    def close(self) -> None:
        if self.hwnd:
            user32.DestroyWindow(self.hwnd)
        if self._font:
            gdi32.DeleteObject(self._font)
        if self.hbitmap:
            gdi32.DeleteObject(self.hbitmap)
        if self.hdc:
            gdi32.DeleteDC(self.hdc)
        try:
            user32.UnregisterClassW("AIAgentOverlayWindow", kernel32.GetModuleHandleW(None))
        except Exception:
            pass
        self.hwnd = self.hdc = self.hbitmap = self.bits = self._font = None

    def _refresh(self) -> None:
        if not self.hwnd or not self.hdc:
            return
        bf = BLENDFUNCTION(0, 0, 255, AC_SRC_ALPHA)
        sz = w.SIZE(self.w, self.h)
        ps = w.POINT(0, 0)
        pd = w.POINT(0, 0)
        if not user32.UpdateLayeredWindow(
            self.hwnd,
            0,
            ctypes.byref(pd),
            ctypes.byref(sz),
            self.hdc,
            ctypes.byref(ps),
            0,
            ctypes.byref(bf),
            ULW_ALPHA,
        ):
            raise ctypes.WinError(ctypes.get_last_error())
        self.bring_to_front()

    def get_bgra_bytes(self) -> bytes:
        if not self.bits:
            return b""
        assert self.bits.value is not None
        return ctypes.string_at(cast(int, self.bits.value), self.w * self.h * 4)

    def set_action(self, action: dict[str, Any] | None) -> None:
        self._action = action

    def set_annotations(self, annotations: dict[str, dict[str, Any]]) -> None:
        self._annotations = annotations

    def set_highlight(self, label: str | None, *, color: int = COLOR_MAGENTA) -> None:
        self._highlight_label = label
        self._highlight_color = color

    def _clear(self) -> None:
        if self.bits:
            ctypes.memset(self.bits, 0, self.w * self.h * 4)

    def _p(self) -> "ctypes.POINTER[ctypes.c_uint32]":
        assert self.bits is not None
        return cast("ctypes.POINTER[ctypes.c_uint32]", ctypes.cast(self.bits, ctypes.POINTER(ctypes.c_uint32)))

    def _thick_h(self, p: "ctypes.POINTER[ctypes.c_uint32]", y: int, x1: int, x2: int, c: int, t: int) -> None:
        for dy in range(-(t // 2), t // 2 + 1):
            yy = y + dy
            if 0 <= yy < self.h:
                row = yy * self.w
                for xx in range(max(0, x1), min(self.w, x2 + 1)):
                    p[row + xx] = c

    def _thick_v(self, p: "ctypes.POINTER[ctypes.c_uint32]", x: int, y1: int, y2: int, c: int, t: int) -> None:
        for dx in range(-(t // 2), t // 2 + 1):
            xx = x + dx
            if 0 <= xx < self.w:
                for yy in range(max(0, y1), min(self.h, y2 + 1)):
                    p[yy * self.w + xx] = c

    def _rect(self, p: "ctypes.POINTER[ctypes.c_uint32]", x: int, y: int, w0: int, h0: int, c: int, t: int) -> None:
        x2 = x + max(1, w0)
        y2 = y + max(1, h0)
        self._thick_h(p, y, x, x2, c, t)
        self._thick_h(p, y2, x, x2, c, t)
        self._thick_v(p, x, y, y2, c, t)
        self._thick_v(p, x2, y, y2, c, t)

    def _fill(self, p: "ctypes.POINTER[ctypes.c_uint32]", x: int, y: int, w0: int, h0: int, c: int) -> None:
        x2 = min(self.w, x + max(1, w0))
        y2 = min(self.h, y + max(1, h0))
        x1 = max(0, x)
        y1 = max(0, y)
        for yy in range(y1, y2):
            row = yy * self.w
            for xx in range(x1, x2):
                p[row + xx] = c

    def _line(self, p: "ctypes.POINTER[ctypes.c_uint32]", x1: int, y1: int, x2: int, y2: int, c: int, t: int) -> None:
        dx, dy = x2 - x1, y2 - y1
        steps = max(1, int((dx * dx + dy * dy) ** 0.5))
        for i in range(steps + 1):
            xx = int(x1 + dx * i / steps)
            yy = int(y1 + dy * i / steps)
            self._thick_h(p, yy, xx - t, xx + t, c, t)

    def _arrow(self, p: "ctypes.POINTER[ctypes.c_uint32]", x: int, y: int, c: int) -> None:
        for i in range(-ARROW_SIZE, ARROW_SIZE + 1):
            for j in range(-ARROW_SIZE, ARROW_SIZE + 1):
                if abs(i) + abs(j) <= ARROW_SIZE:
                    xx, yy = x + i, y + j
                    if 0 <= xx < self.w and 0 <= yy < self.h:
                        p[yy * self.w + xx] = c

    def _cross(self, p: "ctypes.POINTER[ctypes.c_uint32]", x: int, y: int, c: int) -> None:
        self._thick_h(p, y, x - CROSS_SIZE, x + CROSS_SIZE, c, LINE_THICKNESS)
        self._thick_v(p, x, y - CROSS_SIZE, y + CROSS_SIZE, c, LINE_THICKNESS)

    def _force_opaque(self, x: int, y: int, w0: int, h0: int) -> None:
        if not self.bits:
            return
        p = self._p()
        x2 = min(self.w, x + max(1, w0))
        y2 = min(self.h, y + max(1, h0))
        x1 = max(0, x)
        y1 = max(0, y)
        for yy in range(y1, y2):
            row = yy * self.w
            for xx in range(x1, x2):
                p[row + xx] |= 0xFF000000

    def _text(self, x: int, y: int, s: str) -> None:
        if not self.hdc or not s:
            return
        gdi32.TextOutW(self.hdc, x, y, s, len(s))


    def _draw_hud(self, p: "ctypes.POINTER[ctypes.c_uint32]") -> None:
        if not HUD_ENABLED or not self._action:
            return
        tool = str(self._action.get("tool", "")).strip()
        just = str(self._action.get("justification", "")).strip()
        if not tool and not just:
            return
        j = re.sub(r"\s+", " ", just)
        s = f"{tool.upper()}: {j}".strip() if j else tool.upper()
        if len(s) > HUD_MAX_CHARS:
            s = s[:HUD_MAX_CHARS]
        x = HUD_MARGIN
        y = HUD_MARGIN
        bw = min(self.w - x - HUD_MARGIN, min(HUD_MAX_WIDTH, max(220, len(s) * 8 + 16)))
        bh = HUD_HEIGHT
        self._fill(p, x, y, bw, bh, COLOR_BLACK)
        self._text(x + 8, y + 4, s)
        self._force_opaque(x, y, bw, bh)


    def render(self) -> None:
        if not self.bits:
            return
        self._clear()
        p = self._p()

        for lbl, a in self._annotations.items():
            sx = int(a.get("x", 0))
            sy = int(a.get("y", 0))
            w0 = int(a.get("width", 120))
            h0 = int(a.get("height", 120))
            self._rect(p, sx, sy, w0, h0, ANNOTATE_COLOR, 3)

            desc = str(a.get("description", "")).strip()
            t = (f"{lbl}: {desc}" if desc else lbl)[:120]

            bx = sx
            by = max(0, sy - 22)
            bw = min(self.w - bx, max(140, min(520, len(t) * 8 + 12)))
            bh = 20
            self._fill(p, bx, by, bw, bh, COLOR_BLACK)
            self._text(bx + 6, by + 3, t)
            self._force_opaque(bx, by, bw, bh)

        if self._highlight_label:
            a = self._annotations.get(self._highlight_label)
            if a:
                sx = int(a.get("x", 0))
                sy = int(a.get("y", 0))
                w0 = int(a.get("width", 120))
                h0 = int(a.get("height", 120))
                self._rect(p, sx, sy, w0, h0, self._highlight_color, 5)

        if self._action:
            tool = str(self._action.get("tool", ""))
            if tool in ("click", "move", "drag"):
                x1 = int(self._action.get("from_px", self._action.get("px", 0)))
                y1 = int(self._action.get("from_py", self._action.get("py", 0)))
                x2 = int(self._action.get("px", 0))
                y2 = int(self._action.get("py", 0))
                self._line(p, x1, y1, x2, y2, COLOR_GREEN, LINE_THICKNESS)
                self._arrow(p, x2, y2, COLOR_GREEN)
                if tool == "click":
                    self._cross(p, x2, y2, COLOR_RED)
            elif tool == "scroll":
                self._rect(p, 2, 2, self.w - 6, self.h - 6, COLOR_RED, 3)

        self._draw_hud(p)
        self._refresh()


class AnnotationManager:
    def __init__(self, path: Path, conv: CoordConverter):
        self.path, self.conv = path, conv
        self._data: dict[str, dict[str, Any]] = {}
        if path.exists():
            try:
                self._data = cast(dict[str, dict[str, Any]], json.loads(path.read_text(encoding="utf-8")))
            except Exception:
                self._data = {}

    def _save(self) -> None:
        try:
            self.path.write_text(json.dumps(self._data, indent=2), encoding="utf-8")
        except Exception:
            pass

    def add(self, label: str, xn: int, yn: int, width: int, height: int, desc: str, conf: float = 1.0) -> None:
        px, py = self.conv.norm_to_screen(float(xn), float(yn))
        self._data[label] = {
            "x": px,
            "y": py,
            "x_norm": xn,
            "y_norm": yn,
            "width": int(width),
            "height": int(height),
            "description": desc,
            "timestamp": datetime.now().isoformat(),
            "confidence": float(conf),
        }
        self._save()

    def all(self) -> dict[str, dict[str, Any]]:
        return self._data

    def norm(self, label: str) -> tuple[int, int] | None:
        a = self._data.get(label)
        if not a:
            return None
        if "x_norm" in a and "y_norm" in a:
            return (int(a["x_norm"]), int(a["y_norm"]))
        return (int(a.get("x", 0) * 1000 / self.conv.sw), int(a.get("y", 0) * 1000 / self.conv.sh))


def capture_truth_model_bgra(conv: CoordConverter, ov: OverlayManager) -> bytes:
    base_full = _capture_desktop_bgra(conv.sw, conv.sh, include_cursor=True)
    base = _downsample_nn_bgra(base_full, conv.sw, conv.sh, SCREEN_W, SCREEN_H)
    overlay_full = ov.get_bgra_bytes()
    if not overlay_full:
        return base
    overlay = _downsample_nn_bgra(overlay_full, conv.sw, conv.sh, SCREEN_W, SCREEN_H)
    return _alpha_blend_bgra(base, overlay)


def save_truth_screenshot(path: Path, model_bgra: bytes) -> bytes:
    png = _encode_png_rgb(model_bgra, SCREEN_W, SCREEN_H)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(png)
    return png


@dataclass(slots=True)
class ActionResult:
    log_msg: str
    model_msg: str
    action_viz: dict[str, Any] | None
    success: bool
    post_delay_s: float


@dataclass(slots=True)
class ActionCommand:
    tool: ActionTool
    justification: str = ""
    x: float | None = None
    y: float | None = None
    width: int = 120
    height: int = 120
    text: str = ""
    label: str = ""
    dx: float = 0.0
    dy: float = 0.0
    description: str = ""
    confidence: float = 1.0
    x1: float | None = None
    y1: float | None = None
    x2: float | None = None
    y2: float | None = None

    @staticmethod
    def _num(v: Any) -> float | None:
        if v is None:
            return None
        if isinstance(v, list) and v:
            v = v[0]
        try:
            return float(v)
        except (ValueError, TypeError):
            return None

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "ActionCommand":
        return cls(
            tool=cast(ActionTool, d.get("tool", "")),
            justification=str(d.get("justification", "")).strip(),
            x=cls._num(d.get("x")) if "x" in d else None,
            y=cls._num(d.get("y")) if "y" in d else None,
            width=int(d.get("width", 120)),
            height=int(d.get("height", 120)),
            text=str(d.get("text", "")),
            label=str(d.get("label", "")).strip(),
            dx=cls._num(d.get("dx")) or 0.0,
            dy=cls._num(d.get("dy")) or 0.0,
            description=str(d.get("description", "")).strip(),
            confidence=float(d.get("confidence", 1.0)),
            x1=cls._num(d.get("x1")) if "x1" in d else None,
            y1=cls._num(d.get("y1")) if "y1" in d else None,
            x2=cls._num(d.get("x2")) if "x2" in d else None,
            y2=cls._num(d.get("y2")) if "y2" in d else None,
        )

    def validate(self) -> tuple[bool, str]:
        match self.tool:
            case "click" | "move":
                if self.x is None or self.y is None:
                    return False, "missing x,y"
                if not (0 <= self.x <= 1000 and 0 <= self.y <= 1000):
                    return False, "x,y out of range"
            case "drag":
                if self.x1 is None or self.y1 is None or self.x2 is None or self.y2 is None:
                    return False, "missing x1,y1,x2,y2"
                if not (0 <= self.x1 <= 1000 and 0 <= self.y1 <= 1000 and 0 <= self.x2 <= 1000 and 0 <= self.y2 <= 1000):
                    return False, "drag coords out of range"
            case "scroll":
                if abs(self.dx) > 10000 or abs(self.dy) > 10000:
                    return False, "scroll too large"
            case "type":
                if len(self.text) > 2000:
                    return False, "text too long"
            case "annotate":
                if self.x is None or self.y is None or not self.label or not self.description:
                    return False, "missing annotate fields"
                if not (0 <= self.x <= 1000 and 0 <= self.y <= 1000):
                    return False, "x,y out of range"
            case "recall":
                if not self.label:
                    return False, "missing label"
            case "done":
                pass
            case _:
                return False, "unknown tool"
        return True, "ok"

    def signature(self) -> str:
        if self.tool in ("click", "move", "annotate"):
            return f"{self.tool}:{int(self.x or 0)}:{int(self.y or 0)}:{self.label}"
        if self.tool == "drag":
            return f"drag:{int(self.x1 or 0)}:{int(self.y1 or 0)}:{int(self.x2 or 0)}:{int(self.y2 or 0)}"
        if self.tool == "scroll":
            return f"scroll:{int(self.dx)}:{int(self.dy)}"
        if self.tool == "type":
            return f"type:{len(self.text)}"
        if self.tool == "recall":
            return f"recall:{self.label}"
        return self.tool


class ActionExecutor:
    def __init__(self, conv: CoordConverter, ann: AnnotationManager, ov: OverlayManager):
        self.conv, self.ann, self.ov = conv, ann, ov

    def _sync(self, action_viz: dict[str, Any] | None, highlight: str | None = None) -> None:
        self.ov.set_action(action_viz)
        self.ov.set_annotations(self.ann.all())
        self.ov.set_highlight(highlight)
        self.ov.render()

    def _exec_xy(
        self,
        cmd: ActionCommand,
        fn: Callable[[int, int, CoordConverter], None],
        tool: str,
        post_delay: float,
    ) -> ActionResult:
        assert cmd.x is not None and cmd.y is not None
        from_x, from_y = _get_cursor_pos()
        sx, sy = self.conv.norm_to_screen(cmd.x, cmd.y)
        fn(sx, sy, self.conv)
        viz: dict[str, Any] = {"tool": tool, "from_px": from_x, "from_py": from_y, "px": sx, "py": sy, "justification": cmd.justification}
        self._sync(viz)
        return ActionResult(f"{tool} norm({int(cmd.x)},{int(cmd.y)})", f"{tool} at ({int(cmd.x)},{int(cmd.y)})", viz, True, post_delay)

    def _exec_drag(self, cmd: ActionCommand, post_delay: float) -> ActionResult:
        assert cmd.x1 is not None and cmd.y1 is not None and cmd.x2 is not None and cmd.y2 is not None
        cursor_x, cursor_y = _get_cursor_pos()
        sx1, sy1 = self.conv.norm_to_screen(cmd.x1, cmd.y1)
        sx2, sy2 = self.conv.norm_to_screen(cmd.x2, cmd.y2)
        mouse_drag(sx1, sy1, sx2, sy2, self.conv)
        viz: dict[str, Any] = {
            "tool": "drag",
            "from_px": sx1,
            "from_py": sy1,
            "px": sx2,
            "py": sy2,
            "cursor_from_px": cursor_x,
            "cursor_from_py": cursor_y,
            "justification": cmd.justification,
        }
        self._sync(viz)
        return ActionResult(
            f"drag norm({int(cmd.x1)},{int(cmd.y1)})->({int(cmd.x2)},{int(cmd.y2)})",
            f"Dragged ({int(cmd.x1)},{int(cmd.y1)}) to ({int(cmd.x2)},{int(cmd.y2)})",
            viz,
            True,
            post_delay,
        )

    def execute(self, cmd: ActionCommand) -> ActionResult:
        match cmd.tool:
            case "click":
                return self._exec_xy(cmd, mouse_click, "click", DELAY_AFTER_ACTION_S)
            case "move":
                return self._exec_xy(cmd, mouse_move, "move", DELAY_MOVE_HOVER_S)
            case "drag":
                return self._exec_drag(cmd, DELAY_AFTER_ACTION_S)
            case "type":
                type_text(cmd.text)
                viz: dict[str, Any] = {"tool": "type", "text": cmd.text, "justification": cmd.justification}
                self._sync(viz)
                return ActionResult(f"typed {len(cmd.text)} chars", f"Typed {len(cmd.text)} chars", viz, True, DELAY_AFTER_ACTION_S)
            case "scroll":
                scroll(cmd.dx, cmd.dy)
                viz: dict[str, Any] = {"tool": "scroll", "dx": cmd.dx, "dy": cmd.dy, "justification": cmd.justification}
                self._sync(viz)
                return ActionResult(f"scrolled dx={cmd.dx} dy={cmd.dy}", f"Scrolled dx={int(cmd.dx)} dy={int(cmd.dy)}", viz, True, DELAY_SCROLL_S)
            case "annotate":
                assert cmd.x is not None and cmd.y is not None
                lbl = cmd.label
                xn, yn = int(cmd.x), int(cmd.y)
                self.ann.add(lbl, xn, yn, cmd.width, cmd.height, cmd.description, cmd.confidence)
                viz: dict[str, Any] = {"tool": "annotate", "label": lbl, "justification": cmd.description}
                self._sync(viz, highlight=lbl)
                return ActionResult(f"annotated {lbl} norm({xn},{yn})", f"Annotated {lbl}", viz, True, DELAY_AFTER_ACTION_S)
            case "recall":
                pos = self.ann.norm(cmd.label)
                if not pos:
                    self._sync({"tool": "recall", "label": cmd.label, "justification": "missing"})
                    return ActionResult(f"recall {cmd.label} failed", f"Recall {cmd.label} failed", None, False, DELAY_AFTER_ACTION_S)
                xn, yn = pos
                sx, sy = self.conv.norm_to_screen(float(xn), float(yn))
                viz: dict[str, Any] = {"tool": "recall", "px": sx, "py": sy, "justification": cmd.label}
                self._sync(viz, highlight=cmd.label)
                return ActionResult(f"recalled {cmd.label} norm({xn},{yn})", f"Recalled {cmd.label}", viz, True, DELAY_AFTER_ACTION_S)
            case "done":
                self._sync(None)
                return ActionResult("done", "Done", None, True, 0.0)
            case _:
                return ActionResult("unknown", "Unknown", None, False, DELAY_AFTER_ACTION_S)


SYSTEM_PROMPT = """
You control a Windows computer via tools. Output one JSON object only - no extra text.

Tools:
click    -> {"tool":"click","x":500,"y":300,"justification":"..."}
move     -> {"tool":"move","x":500,"y":300,"justification":"..."}
drag     -> {"tool":"drag","x1":500,"y1":300,"x2":700,"y2":300,"justification":"..."}
type     -> {"tool":"type","text":"...","justification":"..."}
scroll   -> {"tool":"scroll","dx":0,"dy":-300,"justification":"..."}
annotate -> {"tool":"annotate","label":"...","x":500,"y":300,"description":"...","width":120,"height":120}
recall   -> {"tool":"recall","label":"...","justification":"..."}
done     -> {"tool":"done","justification":"..."}

Rules:
- x,y are normalized 0..1000 (0,0 top-left; 1000,1000 bottom-right).
- Use CURRENT screenshot primarily. Use PREVIOUS screenshot only as context.
- Use drag for drawing/painting (press-hold-move-release).
- Prefer annotate for stable UI targets, then recall by label.
- Every command includes a justification describing what you see and why.
"""


@dataclass(slots=True)
class HistoryEntry:
    tool: str
    justification: str
    result: str
    success: bool
    timestamp: datetime = field(default_factory=datetime.now)


def _encode_image_data_url(png: bytes) -> str:
    return "data:image/png;base64," + base64.b64encode(png).decode("ascii")


def _format_annotations(anns: dict[str, Any]) -> str:
    if not anns:
        return ""
    items: list[str] = []
    for lbl in list(anns)[:8]:
        d = anns.get(lbl, {})
        desc = str(d.get("description", "")).strip()
        items.append(f"{lbl}: {desc}" if desc else lbl)
    return "\n".join(items)


def _build_messages(goal: str, recent: list[HistoryEntry], anns: dict[str, Any], prev_png: bytes | None, curr_png: bytes) -> list[dict[str, Any]]:
    parts: list[str] = [f"Goal: {goal}\n\n"]
    if recent:
        parts.append("Recent actions (last 4):\n")
        for e in recent[-4:]:
            parts.append(f"- {e.tool} ({'OK' if e.success else 'FAIL'}): {e.justification}\n")
        parts.append("\n")
    labels = _format_annotations(anns)
    if labels:
        parts.append("Labels:\n")
        parts.append(labels)
        parts.append("\n\n")
    parts.append("Decide the next action using the screenshots.\n")

    content: list[dict[str, Any]] = [{"type": "text", "text": "".join(parts)}]
    if prev_png is not None:
        content.append({"type": "text", "text": "Previous screenshot (t-1):"})
        content.append({"type": "image_url", "image_url": {"url": _encode_image_data_url(prev_png), "detail": "auto"}})
    content.append({"type": "text", "text": "Current screenshot (t):"})
    content.append({"type": "image_url", "image_url": {"url": _encode_image_data_url(curr_png), "detail": "auto"}})

    return [{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": content}]


def call_vlm(goal: str, recent: list[HistoryEntry], anns: dict[str, Any], prev_png: bytes | None, curr_png: bytes) -> str:
    payload: dict[str, Any] = {
        "model": MODEL_NAME,
        "messages": _build_messages(goal, recent, anns, prev_png, curr_png),
        "temperature": 0.2,
        "max_tokens": 250,
    }

    last_err: Exception | None = None
    for attempt in range(MAX_API_RETRIES):
        try:
            req = urllib.request.Request(
                API_URL,
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json"},
            )
            with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT_S * (1 + attempt)) as resp:
                data = json.load(resp)
            choices: list[Any] = data.get("choices") or []
            if not choices:
                raise RuntimeError("no choices")
            content: Any = choices[0].get("message", {}).get("content", "")
            if isinstance(content, list):
                return "".join(p.get("text", "") if isinstance(p, dict) else str(p) for p in content)
            return str(content)
        except Exception as e:
            last_err = e
            time.sleep(BACKOFF_FACTOR**attempt)

    raise RuntimeError(f"vlm call failed: {last_err}")


def parse_response(resp: str) -> dict[str, Any] | None:
    s = resp.strip()
    if s.startswith("```"):
        s = "\n".join(line for line in s.splitlines() if not line.strip().startswith("```")).strip()
    start = s.find("{")
    if start == -1:
        return None
    depth = 0
    end = None
    for i in range(start, len(s)):
        if s[i] == "{":
            depth += 1
        elif s[i] == "}":
            depth -= 1
            if depth == 0:
                end = i
                break
    if end is None:
        return None
    candidate = s[start : end + 1]
    try:
        return cast(dict[str, Any], json.loads(candidate))
    except Exception:
        return None


class AgentRunner:
    def __init__(self, goal: str) -> None:
        self.goal = goal
        sw, sh = get_screen_size()
        self.conv = CoordConverter(sw, sh, SCREEN_W, SCREEN_H)

        self.run_dir = Path("dump") / f"run_{time.strftime('%Y%m%d_%H%M%S')}"
        self.run_dir.mkdir(parents=True, exist_ok=True)

        self.ann = AnnotationManager(self.run_dir / "annotations.json", self.conv)
        self.recent: list[HistoryEntry] = []
        self.sigs: list[str] = []
        self.failures = 0
        self.prev_png: bytes | None = None
        self.curr_png: bytes | None = None

    def run(self) -> None:
        log_path = self.run_dir / "log.txt"
        with OverlayManager(self.conv.sw, self.conv.sh) as ov, open(log_path, "w", encoding="utf-8") as logf:
            logf.write(f"Goal: {self.goal}\n")
            logf.write(f"Screen: {self.conv.sw}x{self.conv.sh}\n")
            logf.write(f"Model: {MODEL_NAME}\n")
            logf.write(f"Start: {datetime.now().isoformat()}\n\n")
            logf.flush()

            ov.set_annotations(self.ann.all())
            ov.set_action(None)
            ov.set_highlight(None)
            ov.render()
            ov.reassert_topmost()

            bgra0 = capture_truth_model_bgra(self.conv, ov)
            self.curr_png = save_truth_screenshot(self.run_dir / "step000_screen.png", bgra0)
            self.prev_png = None

            ex = ActionExecutor(self.conv, self.ann, ov)

            for step in range(1, MAX_STEPS + 1):
                assert self.curr_png is not None

                cmd: ActionCommand | None = None
                for attempt in range(MAX_RETRIES):
                    try:
                        resp = call_vlm(self.goal, self.recent, self.ann.all(), self.prev_png, self.curr_png)
                        d = parse_response(resp)
                        if not d:
                            raise ValueError("parse_failed")
                        cmd = ActionCommand.from_dict(d)
                        ok, why = cmd.validate()
                        if not ok:
                            raise ValueError(why)
                        break
                    except Exception as e:
                        if attempt == MAX_RETRIES - 1:
                            logf.write(f"ERROR step {step:03d}: {e}\n")
                            logf.flush()
                            self.failures += 1
                            cmd = None
                        else:
                            time.sleep(RETRY_DELAY_S)

                if cmd is None:
                    if self.failures >= MAX_CONSECUTIVE_FAILURES:
                        logf.write("ABORT: too many failures\n")
                        logf.flush()
                        return
                    continue

                if cmd.tool == "done":
                    logf.write(f"Step {step:03d}: DONE\n{cmd.justification}\n")
                    logf.flush()
                    return

                res = ex.execute(cmd)
                time.sleep(max(0.0, res.post_delay_s))

                ov.reassert_topmost()
                ov.render()

                bgra = capture_truth_model_bgra(self.conv, ov)
                png = save_truth_screenshot(self.run_dir / f"step{step:03d}_screen.png", bgra)

                step_ok = bool(res.success)
                self.failures = 0 if step_ok else self.failures + 1

                self.sigs.append(cmd.signature())
                self.sigs = self.sigs[-10:]
                if len(self.sigs) >= 4 and len(set(self.sigs[-4:])) == 1:
                    self.recent.clear()
                    self.sigs.clear()

                logf.write(f"Step {step:03d}: {cmd.tool} {'OK' if step_ok else 'FAIL'} sig={cmd.signature()}\n")
                if cmd.justification:
                    logf.write(f"  just: {cmd.justification}\n")
                logf.flush()

                self.recent.append(HistoryEntry(cmd.tool, cmd.justification, res.model_msg, step_ok))
                self.recent = self.recent[-8:]

                self.prev_png = self.curr_png
                self.curr_png = png

                if self.failures >= MAX_CONSECUTIVE_FAILURES:
                    logf.write("ABORT: too many failures\n")
                    logf.flush()
                    return


def test_mode() -> None:
    print("=== TEST MODE ===")
    print("Commands: click, move, drag, type, scroll, annotate, recall, done, quit")
    print("Press ENTER to accept defaults; type a value to override.")
    sw, sh = get_screen_size()
    conv = CoordConverter(sw, sh, SCREEN_W, SCREEN_H)
    run_dir = Path("dump") / f"test_{time.strftime('%Y%m%d_%H%M%S')}"
    run_dir.mkdir(parents=True, exist_ok=True)
    ann = AnnotationManager(run_dir / "annotations.json", conv)

    default_just = (
        "TEST JUSTIFICATION: long text to validate overlay label truncation and readability. "
        "Expected to clip when exceeding max label length."
    )

    default_desc = (
        "TEST DESCRIPTION: long annotation description to validate persistence, clipping, and truth capture."
    )
    default_text = (
        "TEST TYPE: The quick brown fox jumps over the lazy dog. 0123456789 "
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz."
    )

    def pstr(label: str, default: str) -> str:
        v = input(f"{label} [{default}]: ").strip()
        return v or default

    def pfloat(label: str, default: float) -> float:
        v = input(f"{label} [{default}]: ").strip()
        return default if not v else float(v)

    def pint(label: str, default: int) -> int:
        v = input(f"{label} [{default}]: ").strip()
        return default if not v else int(v)

    step = 0
    with OverlayManager(sw, sh) as ov:
        ov.set_annotations(ann.all())
        ov.set_action(None)
        ov.set_highlight(None)
        ov.render()
        ov.reassert_topmost()

        png0 = save_truth_screenshot(run_dir / f"test{step:03d}.png", capture_truth_model_bgra(conv, ov))
        _ = png0
        print(f"Screenshot: test{step:03d}.png")

        ex = ActionExecutor(conv, ann, ov)

        while True:
            cmd_str = input("Tool> ").strip().lower()
            if not cmd_str or cmd_str == "quit":
                break
            if cmd_str == "done":
                print("Done.")
                break

            step += 1
            cmd: ActionCommand | None = None
            try:
                match cmd_str:
                    case "click":
                        cmd = ActionCommand("click", pstr("  just", default_just), pfloat("  x", 500.0), pfloat("  y", 500.0))
                    case "move":
                        cmd = ActionCommand("move", pstr("  just", default_just), pfloat("  x", 500.0), pfloat("  y", 500.0))
                    case "drag":
                        cmd = ActionCommand(
                            "drag",
                            pstr("  just", default_just),
                            x1=pfloat("  x1", 420.0),
                            y1=pfloat("  y1", 420.0),
                            x2=pfloat("  x2", 650.0),
                            y2=pfloat("  y2", 650.0),
                        )
                    case "type":
                        cmd = ActionCommand("type", pstr("  just", default_just), text=pstr("  text", default_text))
                    case "scroll":
                        cmd = ActionCommand("scroll", pstr("  just", default_just), dx=pfloat("  dx", 0.0), dy=pfloat("  dy", -600.0))
                    case "annotate":
                        cmd = ActionCommand(
                            "annotate",
                            "",
                            pfloat("  x", 500.0),
                            pfloat("  y", 500.0),
                            width=pint("  width", 180),
                            height=pint("  height", 90),
                            label=pstr("  label", "test_target"),
                            description=pstr("  desc", default_desc),
                            confidence=1.0,
                        )
                    case "recall":
                        cmd = ActionCommand("recall", pstr("  just", default_just), label=pstr("  label", "test_target"))
                    case _:
                        print("Unknown tool.")
                        continue

                ok, why = cmd.validate() if cmd else (False, "no cmd")
                if not ok:
                    print(f"Invalid: {why}")
                    continue

                res = ex.execute(cmd)
                time.sleep(max(0.0, res.post_delay_s))

                ov.reassert_topmost()
                ov.render()

                save_truth_screenshot(run_dir / f"test{step:03d}.png", capture_truth_model_bgra(conv, ov))
                print(f"Result: {res.log_msg}")
                print(f"Screenshot: test{step:03d}.png")
            except Exception as e:
                print(f"Error: {e}")


def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        test_mode()
        return

    print(f"Default task: {DEFAULT_TASK}")
    choice = input("Press ENTER to accept, or type 'n' for custom: ").strip().lower()
    goal = input("Enter task: ").strip() if choice == "n" else DEFAULT_TASK
    if not goal:
        print("No task provided")
        return
    AgentRunner(goal).run()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit(0)