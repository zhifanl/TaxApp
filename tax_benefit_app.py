# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
import re
import sys
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import datetime

# ====================== Design System ======================
# Palette: Official Chinese Government Authority
# Light base, deep navy structure, red accent for urgency
C = {
    # Surfaces
    "bg":           "#f0f2f5",   # page background - off-white
    "surface":      "#ffffff",   # card/panel white
    "surface2":     "#f7f8fa",   # subtle inset
    "surface3":     "#eef0f5",   # table row alt

    # Structure
    "navy":         "#0a2463",   # deep authority navy (primary)
    "navy_light":   "#1a3a8f",   # hover navy
    "navy_dark":    "#061440",   # topbar navy
    "red":          "#c0392b",   # alert red
    "red_light":    "#e74c3c",   # lighter red
    "gold":         "#b8860b",   # seal gold accent

    # Borders & dividers
    "border":       "#d1d9e6",   # standard border
    "border_dark":  "#aab4c8",   # stronger border
    "border_navy":  "#0a2463",   # navy rule line

    # Text
    "text":         "#1a1a2e",   # primary text
    "text_2":       "#3d4b6b",   # secondary text
    "text_3":       "#6b7a99",   # tertiary / placeholder
    "text_inv":     "#ffffff",   # inverse (on dark bg)

    # Semantic
    "ok":           "#1a7a4a",   # success green
    "ok_bg":        "#edfaf3",
    "warn":         "#b8860b",   # warning gold
    "warn_bg":      "#fdf8e1",
    "danger":       "#c0392b",   # danger red
    "danger_bg":    "#fdecea",

    # Inputs
    "input_bg":     "#ffffff",
    "input_border": "#b0bdd0",
    "input_focus":  "#0a2463",

    # Buttons
    "btn_primary":  "#0a2463",
    "btn_success":  "#1a7a4a",
    "btn_warn":     "#b8860b",
    "btn_danger":   "#c0392b",
    "btn_neutral":  "#4a5568",
}

# Font stack — system fonts that render Chinese beautifully
_CH = "Microsoft YaHei" if sys.platform == "win32" else (
      "PingFang SC" if sys.platform == "darwin" else "Noto Sans CJK SC")

F = {
    "display":  (_CH, 18, "bold"),
    "h1":       (_CH, 14, "bold"),
    "h2":       (_CH, 12, "bold"),
    "body":     (_CH, 11),
    "body_b":   (_CH, 11, "bold"),
    "small":    (_CH, 9),
    "small_b":  (_CH, 9, "bold"),
    "mono":     ("Courier New", 11),
    "tag":      (_CH, 9, "bold"),
}

# ====================== Data ======================
ISSUE_GUIDE_MAP = {
    "\u540c\u671f\u7533\u62a5\u7684\u589e\u503c\u7a0e\u6536\u5165\u4e0e\u4f01\u4e1a\u6240\u5f97\u7a0e\u6536\u5165\u6709\u5dee\u5f02": [
        "1. \u662f\u5426\u7531\u4e8e\u4e24\u7a0e\u786e\u8ba4\u6536\u5165\u65f6\u95f4\u4e0d\u4e00\u81f4\u3002",
        "2. \u662f\u5426\u7531\u4e8e\u4e24\u7a0e\u786e\u8ba4\u6536\u5165\u53e3\u5f84\u4e0d\u4e00\u81f4\u5bfc\u81f4\u3002",
        "3. \u7ed3\u5408\u5408\u540c\u7b7e\u8ba2\u60c5\u51b5\u3001\u9879\u76ee\u8fdb\u5ea6\u3001\u9280\u884c\u6d41\u6c34\u65f6\u95f4\u8fdb\u884c\u7efc\u5408\u5224\u65ad\u3002",
    ],
    "\u5de5\u5546\u4e1a\u7eb3\u7a0e\u4eba\u6210\u672c\u8d39\u7528\u5360\u6bd4\u5f02\u5e38": [
        "1. \u68c0\u67e5\u4ee3\u53d1\u4eba\u5458\u5de5\u8d44\u662f\u5426\u771f\u5b9e\u5230\u8d26\uff08\u9280\u884c\u6d41\u6c34\uff09\u3002",
        "2. \u5b9e\u5730\u6838\u67e5\u5e93\u5b58\uff0c\u56fa\u5b9a\u8d44\u4ea7\u8fdb\u884c\u81ea\u67e5\u6e05\u70b9\u3002",
        "3. \u6838\u67e5\u8fd0\u8f93\u53d1\u7968\u3001\u51fa\u8d27\u6e05\u5355\u7b49\u8f85\u52a9\u6750\u6599\u3002",
    ],
    "\u8d39\u7528\u504f\u9ad8": [
        "1. \u6838\u5b9e\u8d39\u7528\u662f\u5426\u771f\u5b9e\u53d1\u751f\u3002",
        "2. \u68c0\u67e5\u8d39\u7528\u5bf9\u5e94\u7684\u5408\u540c\u3001\u53d1\u7968\u53ca\u9280\u884c\u6d41\u6c34\u3002",
        "3. \u5224\u65ad\u662f\u5426\u5b58\u5728\u4e2a\u4eba\u6d88\u8d39\u6216\u865a\u5217\u8d39\u7528\u60c5\u5f62\u3002",
    ],
    "\u6210\u672c\u504f\u9ad8": [
        "1. \u6838\u5b9e\u6210\u672c\u662f\u5426\u771f\u5b9e\u53d1\u751f\u3002",
        "2. \u6838\u67e5\u5e93\u5b58\u3001\u51fa\u5165\u5e93\u8bb0\u5f55\u53ca\u5bf9\u5e94\u53d1\u7968\u3002",
        "3. \u5224\u65ad\u662f\u5426\u5b58\u5728\u865a\u5217\u6210\u672c\u60c5\u5f62\u3002",
    ],
    "\u670d\u52a1\u4e1a\u7eb3\u7a0e\u4eba\u6210\u672c\u8d39\u7528\u5360\u6bd4\u5f02\u5e38": [
        "1. \u68c0\u67e5\u4eba\u5de5\u6210\u672c\u662f\u5426\u771f\u5b9e\uff08\u9280\u884c\u6d41\u6c34\uff09\u3002",
        "2. \u6838\u67e5\u5916\u5305\u5408\u540c\u53ca\u5bf9\u5e94\u53d1\u7968\u3002",
        "3. \u4e0e\u540c\u884c\u4e1a\u5e73\u5747\u6c34\u5e73\u5bf9\u6bd4\u5206\u6790\u3002",
    ],
    "\u7591\u4f3c\u672a\u53d6\u5f97\u5408\u6cd5\u6709\u6548\u51ed\u8bc1\u5217\u652f": [
        "1. \u662f\u5426\u5c5e\u4e8e\u4ee5\u4e0b\u516b\u7c7b\u53ef\u7a0e\u524d\u6263\u9664\u60c5\u5f62\uff1a",
        "   \u2022 \u5883\u5916\u8d39\u7528\u5f62\u5f0f\u53d1\u7968",
        "   \u2022 \u5dee\u65c5\u5305\u5e72\u5185\u90e8\u51ed\u8bc1",
        "   \u2022 \u653f\u5e9c\u6536\u8d39\u8d22\u653f\u7968\u636e",
        "   \u2022 \u653f\u5e9c\u5e94\u7a0e\u6536\u6b3e\u51ed\u8bc1",
        "   \u2022 \u6c34\u7535\u8d39\u5206\u5272\u6263\u9664",
        "   \u2022 \u4e0a\u5e74\u5e93\u5b58\u672c\u5e74\u7ed3\u8f6c",
        "   \u2022 \u6682\u4f30\u5165\u5e93\u8de8\u671f\u7968\u636e",
        "2. \u9664\u4e0a\u8ff0\u60c5\u5f62\u5916\uff0c\u76f8\u5173\u652f\u51fa\u4e0d\u5f97\u7a0e\u524d\u6263\u9664\u3002",
        "3. \u901a\u8fc7\u9280\u884c\u6d41\u6c34\u3001\u5408\u540c\u3001\u7269\u6d41\u8d44\u6599\u8fdb\u884c\u4ea4\u53c9\u6838\u5b9e\u3002",
    ],
    "\u7591\u4f3c\u591a\u5217\u5de5\u8d44\u85aa\u91d1\u652f\u51fa\u6216\u5c11\u6263\u7f34\u4e2a\u4eba\u6240\u5f97\u7a0e": [
        "1. \u662f\u5426\u5b58\u5728\u591a\u5217\u5de5\u8d44\u652f\u51fa\u3001\u5c11\u7f34\u4e2a\u4eba\u6240\u5f97\u7a0e\u60c5\u5f62\u3002",
        "2. \u8981\u6c42\u63d0\u4f9b\u52b3\u52a8\u5408\u540c\u3001\u5de5\u8d44\u652f\u4ed8\u8bb0\u5f55\u3002",
        "3. \u6838\u5bf9\u5458\u5de5\u82b1\u540d\u518c\uff0c\u5e76\u8fdb\u884c\u4eba\u5458\u8bbf\u8c08\u3002",
    ],
    "\u8fdb\u4e00\u6b65\u6838\u5b9e\u7eb3\u7a0e\u4eba\u662f\u5426\u5c11\u8ba1\u5370\u82b1\u7a0e\u8ba1\u7a0e\u4f9d\u636e": [
        "1. \u5411\u4e0a\u6e38\u5355\u4f4d\u53d1\u51fd\uff0c\u5f00\u5c55\u5f02\u5730\u534f\u67e5\u3002",
        "2. \u67e5\u9605\u5408\u540c\u3001\u53d1\u7968\u53ca\u9280\u884c\u6d41\u6c34\uff0c\u7ea6\u8c08\u4f01\u4e1a\u8fdb\u884c\u6838\u5b9e\u3002",
    ],
    "\u7591\u4f3c\u5c11\u8f6c\u51fa\u7528\u4e8e\u7b80\u6613\u8ba1\u7a0e\u7684\u8fdb\u9879\u7a0e\u989d": [
        "1. \u67e5\u627e\u6297\u6263\u8fdb\u9879\u53d1\u7968\uff0c\u662f\u5426\u5b58\u5728\u201c\u623f\u5c4b\u51fa\u51fa\u79df\u201d\u201c\u7269\u4e1a\u670d\u52a1\u201d\u7b49\u540c\u65f6\u7528\u4e8e\u5e94\u7a0e/\u514d\u7a0e\u9879\u76ee\u3002",
        "2. \u6838\u67e5\u662f\u5426\u5b58\u5728\u514d\u7a0e\u8d27\u7269\u9500\u552e\u4f46\u672a\u8f6c\u51fa\u8fdb\u9879\u7a0e\u989d\u60c5\u5f62\u3002",
    ],
}

FIELD_SOURCE_MAP = {
    "A_\u4e3b\u8425\u884c\u4e1a": {
        "title": "\u4e3b\u8425\u884c\u4e1a",
        "source": "\u91d1\u7a0e\u4e09\u671f\u7cfb\u7edf \u2014 \u7a0e\u52a1\u767b\u8bb0\u4fe1\u606f\u67e5\u8be2\u6a21\u5757\u4e0b\u300c\u884c\u4e1a\u300d\u67e5\u8be2\u7ed3\u679c\u9009\u62e9\u3002",
    },
    "B_\u671f\u95f4": {
        "title": "\u671f\u95f4",
        "source": "\u4e0d\u5c11\u4e8e\u4e00\u4e2a\u7eb3\u7a0e\u671f\uff0c\u5efa\u8bae\u5f55\u5165\u5b8c\u6574\u5e74\u5ea6\u3002",
    },
    "C_\u8425\u4e1a\u6536\u5165": {
        "title": "\u8425\u4e1a\u6536\u5165",
        "source": "\u300a\u4e2d\u534e\u4eba\u6c11\u5171\u548c\u56fd\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u8868\u300b\u300c\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u4e3b\u8868\u300d\u7b2c 1 \u680f\u300c\u4e00\u3001\u8425\u4e1a\u6536\u5165\u300d\u3002",
    },
    "D_\u9500\u552e\u6536\u5165": {
        "title": "\u9500\u552e\u6536\u5165",
        "source": "\u5404\u6708\u300a\u589e\u503c\u7a0e\u53ca\u9644\u52a0\u7a0e\u8d39\u7533\u62a5\u8868\uff08\u4e00\u822c\u7eb3\u7a0e\u4eba\u9002\u7528\uff09\u300b\u7b2c 1\u30015\u30017\u30018 \u680f\u4e4b\u548c\u3002\u540c\u4e00\u5e74\u5ea6\u8fde\u7eed\u5404\u6708\u53ef\u53c2\u8003\u6700\u540e\u4e00\u671f\u7d2f\u8ba1\u6570\u636e\u3002",
    },
    "E_\u6210\u672c": {
        "title": "\u6210\u672c",
        "source": "\u300a\u4e2d\u534e\u4eba\u6c11\u5171\u548c\u56fd\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u8868\u300b\u300c\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u4e3b\u8868\u300d\u7b2c 2 \u680f\u300c\u51cf\uff1a\u8425\u4e1a\u6210\u672c\u300d\u3002",
    },
    "F_\u9500\u552e\u8d39\u7528": {
        "title": "\u9500\u552e\u8d39\u7528",
        "source": "\u300a\u4e2d\u534e\u4eba\u6c11\u5171\u548c\u56fd\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u8868\u300b\u300c\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u4e3b\u8868\u300d\u7b2c 4 \u680f\u300c\u51cf\uff1a\u9500\u552e\u8d39\u7528\u300d\u3002",
    },
    "G_\u7ba1\u7406\u8d39\u7528": {
        "title": "\u7ba1\u7406\u8d39\u7528",
        "source": "\u300a\u4e2d\u534e\u4eba\u6c11\u5171\u548c\u56fd\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u8868\u300b\u300c\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u4e3b\u8868\u300d\u7b2c 5 \u680f\u300c\u51cf\uff1a\u7ba1\u7406\u8d39\u7528\u300d\u3002",
    },
    "H_\u8d22\u52a1\u8d39\u7528": {
        "title": "\u8d22\u52a1\u8d39\u7528",
        "source": "\u300a\u4e2d\u534e\u4eba\u6c11\u5171\u548c\u56fd\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u8868\u300b\u300c\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u4e3b\u8868\u300d\u7b2c 5 \u680f\u300c\u51cf\uff1a\u8d22\u52a1\u8d39\u7528\u300d\u3002",
    },
    "I_\u5de5\u8d44\u85aa\u91d1": {
        "title": "\u5de5\u8d44\u85aa\u91d1",
        "source": "\u300a\u4e2d\u534e\u4eba\u6c11\u5171\u548c\u56fd\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u8868\u300b A105050 \u804c\u5de5\u85aa\u916c\u652f\u51fa\u53ca\u7eb3\u7a0e\u8c03\u6574\u660e\u7ec6\u8868\u7b2c 1 \u680f\u300c\u4e00\u3001\u5de5\u8d44\u85aa\u91d1\u652f\u51fa\u2014\u5b9e\u9645\u53d1\u751f\u989d\u300d\u3002",
    },
    "J_\u4e2a\u7a0e\u6263\u7f34\u5de5\u8d44\u603b\u989d": {
        "title": "\u4e2a\u7a0e\u6263\u7f34\u5de5\u8d44\u603b\u989d",
        "source": "\u81ea\u7136\u4eba\u7535\u5b50\u7a0e\u52a1\u5c40 \u2014 \u6263\u7f34\u7533\u62a5\u5206\u6237\u6e05\u518c\u67e5\u8be2\uff08ITS\uff09\uff0c\u8f93\u5165\u7a0e\u53f7\uff0c\u67e5\u8be2\u4e2a\u4eba\u6240\u5f97\u7a0e\u4ee3\u6263\u4ee3\u7f34\u5de5\u8d44\u603b\u989d\u3002",
    },
    "K_\u8ba1\u63d0\u6298\u65f7": {
        "title": "\u8ba1\u63d0\u6298\u65f7",
        "source": "\u300a\u4e2d\u534e\u4eba\u6c11\u5171\u548c\u56fd\u4f01\u4e1a\u6240\u5f97\u7a0e\u5e74\u5ea6\u7eb3\u7a0e\u7533\u62a5\u8868\u300b A105080 \u8d44\u4ea7\u6298\u65f7\u3001\u6478\u9500\u53ca\u7eb3\u7a0e\u8c03\u6574\u660e\u7ec6\u8868\u7b2c 30 \u680f\u300c\u672c\u5e74\u6298\u65f7\u3001\u6478\u9500\u989d\u2014\u5408\u8ba1\u300d\u3002",
    },
    "L_\u5f53\u671f\u5f00\u7968\u989d\u5ea6": {
        "title": "\u5f53\u671f\u5f00\u7968\u989d\u5ea6",
        "source": "\u91d1\u7a0e\u4e09\u671f\u7cfb\u7edf \u2014 \u98ce\u9669\u7ba1\u7406\u7cfb\u7edf\u9996\u9875 \u2014 \u7eb3\u7a0e\u4eba\u53d1\u7968\u7968\u79cd\u5206\u6790\u3002",
    },
    "M_\u5f53\u671f\u53d7\u7968\u989d\u5ea6": {
        "title": "\u5f53\u671f\u53d7\u7968\u989d\u5ea6",
        "source": "\u91d1\u7a0e\u4e09\u671f\u7cfb\u7edf \u2014 \u98ce\u9669\u7ba1\u7406\u7cfb\u7edf\u9996\u9875 \u2014 \u7eb3\u7a0e\u4eba\u53d1\u7968\u7968\u79cd\u5206\u6790\u3002",
    },
    "N_\u5370\u82b1\u7a0e\u8ba1\u7a0e\u4f9d\u636e": {
        "title": "\u5370\u82b1\u7a0e\u8ba1\u7a0e\u4f9d\u636e",
        "source": "\u91d1\u7a0e\u4e09\u671f\u7cfb\u7edf \u2014 \u7533\u62a5\u660e\u7ec6\u67e5\u8be2\uff0c\u9009\u62e9\u5bf9\u5e94\u5c5e\u671f\u3001\u5f55\u5165\u7eb3\u7a0e\u4eba\u8bc6\u522b\u53f7\uff0c\u300c\u5f81\u6536\u9879\u76ee\u300d\u9009\u62e9\u300c\u5370\u82b1\u7a0e\u300d\uff0c\u67e5\u8be2\u7ed3\u679c\u4e2d\u7684\u300c\u8ba1\u7a0e\u4f9d\u636e\u300d\u5408\u8ba1\u3002",
    },
    "O_\u7b80\u6613\u8ba1\u7a0e\u9500\u552e\u989d": {
        "title": "\u7b80\u6613\u8ba1\u7a0e\u9500\u552e\u989d",
        "source": "\u5404\u6708\u300a\u589e\u503c\u7a0e\u53ca\u9644\u52a0\u7a0e\u8d39\u7533\u62a5\u8868\uff08\u4e00\u822c\u7eb3\u7a0e\u4eba\u9002\u7528\uff09\u300b\u7b2c 5 \u680f\u300c\u6309\u7b80\u6613\u5f81\u6536\u529e\u6cd5\u8ba1\u7a0e\u9500\u552e\u989d\u300d\u4e4b\u548c\u3002",
    },
    "P_\u514d\u7a0e\u9500\u552e\u989d": {
        "title": "\u514d\u7a0e\u9500\u552e\u989d",
        "source": "\u5404\u6708\u300a\u589e\u503c\u7a0e\u53ca\u9644\u52a0\u7a0e\u8d39\u7533\u62a5\u8868\uff08\u4e00\u822c\u7eb3\u7a0e\u4eba\u9002\u7528\uff09\u300b\u7b2c 8 \u680f\u300c\u514d\u7a0e\u9500\u552e\u989d\u300d\u4e4b\u548c\u3002",
    },
    "Q_\u8fdb\u9879\u7a0e\u989d": {
        "title": "\u8fdb\u9879\u7a0e\u989d",
        "source": "\u5404\u6708\u300a\u589e\u503c\u7a0e\u53ca\u9644\u52a0\u7a0e\u8d39\u7533\u62a5\u8868\uff08\u4e00\u822c\u7eb3\u7a0e\u4eba\u9002\u7528\uff09\u300b\u7b2c 12 \u680f\u300c\u8fdb\u9879\u7a0e\u989d\u300d\u4e4b\u548c\u3002",
    },
    "R_\u8fdb\u9879\u7a0e\u989d\u8f6c\u51fa": {
        "title": "\u8fdb\u9879\u7a0e\u989d\u8f6c\u51fa\u989d",
        "source": "\u5404\u6708\u300a\u589e\u503c\u7a0e\u53ca\u9644\u52a0\u7a0e\u8d39\u7533\u62a5\u8868\uff08\u4e00\u822c\u7eb3\u7a0e\u4eba\u9002\u7528\uff09\u300b\u7b2c 13 \u680f\u300c\u8fdb\u9879\u7a0e\u989d\u8f6c\u51fa\u300d\u4e4b\u548c\u3002",
    },
}

# ====================== UI Helpers ======================
def _hex_to_rgb(h):
    h = h.lstrip("#")
    return int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)

def _blend(c1, c2, t=0.15):
    r1,g1,b1 = _hex_to_rgb(c1)
    r2,g2,b2 = _hex_to_rgb(c2)
    return f"#{int(r1+(r2-r1)*t):02x}{int(g1+(g2-g1)*t):02x}{int(b1+(b2-b1)*t):02x}"

def mk_entry(parent, width=20):
    e = tk.Entry(
        parent,
        font=F["body"],
        width=width,
        relief=tk.FLAT,
        bd=0,
        fg=C["text"],
        bg=C["input_bg"],
        insertbackground=C["navy"],
        selectbackground=C["navy"],
        selectforeground="#fff",
        highlightthickness=1,
        highlightbackground=C["input_border"],
        highlightcolor=C["input_focus"],
    )
    return e

def mk_flat_btn(parent, text, bg, fg="#fff", command=None, width=10, padx=18, pady=8):
    b = tk.Button(
        parent, text=text, font=F["body_b"],
        bg=bg, fg=fg,
        activebackground=_blend(bg, "#000000", 0.15),
        activeforeground=fg,
        relief=tk.FLAT, bd=0,
        padx=padx, pady=pady,
        cursor="hand2", width=width,
        command=command,
        highlightthickness=0,
    )
    b.bind("<Enter>", lambda e: b.config(bg=_blend(bg, "#000000", 0.12)))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b

def hdivider(parent, color=None, pady=0):
    f = tk.Frame(parent, bg=color or C["border"], height=1)
    f.pack(fill=tk.X, pady=pady)
    return f

def section_label(parent, text):
    """Red left-bar section header — standard government form style"""
    row = tk.Frame(parent, bg=C["surface"])
    row.pack(fill=tk.X, pady=(14, 8))
    tk.Frame(row, bg=C["navy"], width=4).pack(side=tk.LEFT, fill=tk.Y)
    tk.Label(row, text=f"  {text}", font=F["h1"],
             bg=C["surface"], fg=C["navy"],
             anchor="w", padx=6, pady=6).pack(side=tk.LEFT, fill=tk.X)
    return row

def popup_base(root, title, w=680, h=460):
    win = tk.Toplevel(root)
    win.title(title)
    win.configure(bg=C["surface"])
    win.geometry(f"{w}x{h}+{root.winfo_x()+60}+{root.winfo_y()+60}")
    win.resizable(False, False)
    win.grab_set()
    # Header
    hdr = tk.Frame(win, bg=C["navy_dark"], pady=14)
    hdr.pack(fill=tk.X)
    tk.Frame(hdr, bg=C["gold"], width=3).pack(side=tk.LEFT, fill=tk.Y)
    return win, hdr

# ====================== Main App ======================
class TaxBenefitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("\u751f\u4ea7\u7ecf\u8425\u5408\u89c4\u68c0\u6d4b\u5668")
        self.root.configure(bg=C["bg"])
        self.root.geometry("1100x820+60+30")
        self.root.minsize(1000, 700)

        self._setup_styles()

        self.tax_items = [
            "\u5236\u9020\u4e1a\u7f13\u7a0e",
            "\u4ea4\u901a\u8fd0\u8f93\u51cf\u514d",
            "\u5c0f\u5fae\u4f01\u4e1a\u51cf\u514d",
            "\u9ad8\u65b0\u6280\u672f\u4f01\u4e1a\u51cf\u514d",
            "\u73af\u4fdd\u8bbe\u5907\u51cf\u514d",
            "\u7814\u53d1\u8d39\u7528\u52a0\u8ba1\u6263\u9664",
        ]
        self.entries = {}
        self.rule_inputs = {}
        self._status_var = tk.StringVar(value="\u5c31\u7eea  \u8bf7\u586b\u5199\u4f01\u4e1a\u4fe1\u606f\u5e76\u5f55\u5165\u68c0\u6d4b\u6570\u636e")

        self._build()

        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after_idle(self.root.attributes, "-topmost", False)

    # --------------------------------------------------
    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("default")
        s.configure("Gov.TCombobox",
            foreground=C["text"],
            fieldbackground=C["input_bg"],
            background=C["input_bg"],
            selectbackground=C["navy"],
            selectforeground="#fff",
            bordercolor=C["input_border"],
            arrowcolor=C["navy"],
            relief="flat",
            padding=4,
        )
        s.map("Gov.TCombobox",
            fieldbackground=[("readonly", C["input_bg"])],
            foreground=[("readonly", C["text"])])
        s.configure("Gov.Vertical.TScrollbar",
            background=C["border"],
            troughcolor=C["surface2"],
            arrowcolor=C["text_3"],
            bordercolor=C["border"],
            relief="flat",
            width=10,
        )
        s.map("Gov.Vertical.TScrollbar",
            background=[("active", C["navy_light"])])

    # --------------------------------------------------
    def _build(self):
        # ── Top navigation bar ──────────────────────────
        self._build_topbar()

        # ── Main body: sidebar + content ────────────────
        body = tk.Frame(self.root, bg=C["bg"])
        body.pack(fill=tk.BOTH, expand=True)

        self._build_sidebar(body)
        self._build_content(body)

        # ── Status bar ──────────────────────────────────
        self._build_statusbar()

    # --------------------------------------------------
    def _build_topbar(self):
        bar = tk.Frame(self.root, bg=C["navy_dark"], height=56)
        bar.pack(fill=tk.X)
        bar.pack_propagate(False)

        # Red-gold left accent block
        accent = tk.Frame(bar, bg=C["gold"], width=5)
        accent.pack(side=tk.LEFT, fill=tk.Y)

        # Emblem placeholder
        emblem = tk.Label(bar, text="\u263c", font=(_CH, 22),
                          bg=C["navy_dark"], fg=C["gold"])
        emblem.pack(side=tk.LEFT, padx=(14, 4))

        # Title block
        title_blk = tk.Frame(bar, bg=C["navy_dark"])
        title_blk.pack(side=tk.LEFT, padx=(0, 30))
        tk.Label(title_blk, text="\u751f\u4ea7\u7ecf\u8425\u5408\u89c4\u68c0\u6d4b\u5668",
                 font=F["display"], bg=C["navy_dark"], fg="#ffffff").pack(anchor="w")
        tk.Label(title_blk,
                 text="Tax Compliance Risk Monitor  |  \u91d1\u7a0e\u4e09\u671f\u540c\u6b3e",
                 font=F["small"], bg=C["navy_dark"], fg="#99aac8").pack(anchor="w")

        # Right: datetime
        now = datetime.datetime.now().strftime("%Y\u5e74%m\u6708%d\u65e5  %H:%M")
        tk.Label(bar, text=now, font=F["small"],
                 bg=C["navy_dark"], fg="#99aac8").pack(side=tk.RIGHT, padx=20)

        # Version tag
        ver = tk.Label(bar, text=" v2.0 ", font=F["tag"],
                       bg=C["gold"], fg=C["navy_dark"])
        ver.pack(side=tk.RIGHT, padx=(0, 12), pady=16)

    # --------------------------------------------------
    def _build_sidebar(self, parent):
        sb = tk.Frame(parent, bg=C["navy"], width=180)
        sb.pack(side=tk.LEFT, fill=tk.Y)
        sb.pack_propagate(False)

        tk.Label(sb, text="\u529f\u80fd\u5bfc\u822a", font=F["small_b"],
                 bg=C["navy"], fg="#6a85b3",
                 padx=16, pady=12).pack(anchor="w")

        hdivider(sb, color="#1e3a70", pady=0)

        menu_items = [
            ("\u2022  \u4f01\u4e1a\u4fe1\u606f\u767b\u8bb0",  True),
            ("\u2022  \u7591\u70b9\u89c4\u5219\u68c0\u67e5",   True),
            ("\u2022  \u7a0e\u6536\u4f18\u60e0\u6838\u67e5",   True),
            ("\u2022  \u64cd\u4f5c\u63a7\u5236\u53f0",         True),
        ]
        for label, active in menu_items:
            bg = "#0d2d6b" if active else C["navy"]
            fg = "#e8ecf4" if active else "#4a6898"
            f = tk.Frame(sb, bg=bg, cursor="hand2")
            f.pack(fill=tk.X)
            tk.Label(f, text=label, font=F["body"],
                     bg=bg, fg=fg,
                     padx=16, pady=10, anchor="w").pack(fill=tk.X)
            hdivider(f, color="#1e3a70")

        # Sidebar footer: legal note
        tk.Label(sb,
                 text="\u4e2d\u534e\u4eba\u6c11\u5171\u548c\u56fd\n\u56fd\u5bb6\u7a0e\u52a1\u5c40\n\u5185\u90e8\u7cfb\u7edf",
                 font=F["small"], bg=C["navy"], fg="#3d5a8a",
                 justify="center").pack(side=tk.BOTTOM, pady=20)

    # --------------------------------------------------
    def _build_content(self, parent):
        # Scrollable right content area
        right = tk.Frame(parent, bg=C["bg"])
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(right, bg=C["bg"], highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(right, orient="vertical",
                             command=canvas.yview, style="Gov.Vertical.TScrollbar")
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(fill=tk.BOTH, expand=True)

        self.content_frame = tk.Frame(canvas, bg=C["bg"])
        self.content_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        cw = canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cw, width=e.width))
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self._build_company_card()
        self._build_rules_card()
        self._build_tax_table_card()
        self._build_action_bar()

    # --------------------------------------------------
    def _card(self, title, subtitle=""):
        """Standard government form card"""
        outer = tk.Frame(self.content_frame, bg=C["bg"], padx=20, pady=8)
        outer.pack(fill=tk.X)

        card = tk.Frame(outer, bg=C["surface"],
                        highlightbackground=C["border"],
                        highlightthickness=1)
        card.pack(fill=tk.X)

        # Card header: thin navy top stripe
        stripe = tk.Frame(card, bg=C["navy"], height=3)
        stripe.pack(fill=tk.X, side=tk.TOP)

        # Header row
        hdr_row = tk.Frame(card, bg=C["surface2"],
                           highlightbackground=C["border"], highlightthickness=0)
        hdr_row.pack(fill=tk.X)

        tk.Label(hdr_row, text=title, font=F["h1"],
                 bg=C["surface2"], fg=C["navy"],
                 padx=16, pady=10, anchor="w").pack(side=tk.LEFT)

        if subtitle:
            tk.Label(hdr_row, text=subtitle, font=F["small"],
                     bg=C["surface2"], fg=C["text_3"],
                     padx=0, pady=10).pack(side=tk.LEFT)

        # Thin rule between header and body
        tk.Frame(card, bg=C["border"], height=1).pack(fill=tk.X)

        body = tk.Frame(card, bg=C["surface"], padx=20, pady=16)
        body.pack(fill=tk.X)
        return body

    # --------------------------------------------------
    def _build_company_card(self):
        body = self._card("\u4e00\u3001\u4f01\u4e1a\u57fa\u672c\u4fe1\u606f",
                          "  \u8bf7\u6309\u5b9e\u586b\u5199\uff0c\u7b26\u5408\u5de5\u5546\u767b\u8bb0\u4fe1\u606f")

        fields = [
            ("\u793e\u4f1a\u4fe1\u7528\u4ee3\u7801\uff1a", "credit_code_entry", 38),
            ("\u4f01\u4e1a\u540d\u79f0\uff1a",             "company_name_entry", 38),
        ]
        for lbl, attr, w in fields:
            row = tk.Frame(body, bg=C["surface"])
            row.pack(fill=tk.X, pady=5)
            tk.Label(row, text=lbl, font=F["body_b"],
                     bg=C["surface"], fg=C["text_2"],
                     width=16, anchor="e").pack(side=tk.LEFT)
            e = mk_entry(row, width=w)
            e.pack(side=tk.LEFT, padx=(8, 0), ipady=5)
            setattr(self, attr, e)

        # Required note
        tk.Label(body, text="\u26a0  \u5e26 \u2731 \u9879\u4e3a\u5fc5\u586b\u9879",
                 font=F["small"], bg=C["surface"], fg=C["text_3"]).pack(anchor="w", pady=(6,0))

    # --------------------------------------------------
    def _build_rules_card(self):
        body = self._card("\u4e8c\u3001\u7591\u70b9\u89c4\u5219\u68c0\u67e5\u6570\u636e\u5f55\u5165",
                          "  Excel \u540c\u6b3e\u8ba1\u7b97\u53e3\u5f84")

        fields = [
            {"name":"\u4e3b\u8425\u884c\u4e1a (A)", "key":"A_\u4e3b\u8425\u884c\u4e1a", "type":"select",
             "choices":["\u6279\u53d1\u96f6\u552e","\u5236\u9020","\u5efa\u7b51\u5b89\u88c5",
                        "\u4ea4\u901a\u8fd0\u8f93","\u751f\u6d3b\u670d\u52a1","\u5176\u4ed6"]},
            {"name":"\u671f\u95f4 (B)",                 "key":"B_\u671f\u95f4",              "type":"text"},
            {"name":"\u8425\u4e1a\u6536\u5165 (C)",     "key":"C_\u8425\u4e1a\u6536\u5165", "type":"number"},
            {"name":"\u9500\u552e\u6536\u5165 (D)",     "key":"D_\u9500\u552e\u6536\u5165", "type":"number"},
            {"name":"\u6210\u672c (E)",                  "key":"E_\u6210\u672c",             "type":"number"},
            {"name":"\u9500\u552e\u8d39\u7528 (F)",     "key":"F_\u9500\u552e\u8d39\u7528", "type":"number"},
            {"name":"\u7ba1\u7406\u8d39\u7528 (G)",     "key":"G_\u7ba1\u7406\u8d39\u7528", "type":"number"},
            {"name":"\u8d22\u52a1\u8d39\u7528 (H)",     "key":"H_\u8d22\u52a1\u8d39\u7528", "type":"number"},
            {"name":"\u5de5\u8d44\u85aa\u91d1 (I)",     "key":"I_\u5de5\u8d44\u85aa\u91d1", "type":"number"},
            {"name":"\u4e2a\u7a0e\u6263\u7f34\u5de5\u8d44 (J)", "key":"J_\u4e2a\u7a0e\u6263\u7f34\u5de5\u8d44\u603b\u989d", "type":"number"},
            {"name":"\u8ba1\u63d0\u6298\u65f7 (K)",     "key":"K_\u8ba1\u63d0\u6298\u65f7", "type":"number"},
            {"name":"\u5f53\u671f\u5f00\u7968\u989d\u5ea6 (L)", "key":"L_\u5f53\u671f\u5f00\u7968\u989d\u5ea6", "type":"number"},
            {"name":"\u5f53\u671f\u53d7\u7968\u989d\u5ea6 (M)", "key":"M_\u5f53\u671f\u53d7\u7968\u989d\u5ea6", "type":"number"},
            {"name":"\u5370\u82b1\u7a0e\u8ba1\u7a0e\u4f9d\u636e (N)", "key":"N_\u5370\u82b1\u7a0e\u8ba1\u7a0e\u4f9d\u636e", "type":"number"},
            {"name":"\u7b80\u6613\u8ba1\u7a0e\u9500\u552e\u989d (O)", "key":"O_\u7b80\u6613\u8ba1\u7a0e\u9500\u552e\u989d", "type":"number"},
            {"name":"\u514d\u7a0e\u9500\u552e\u989d (P)", "key":"P_\u514d\u7a0e\u9500\u552e\u989d", "type":"number"},
            {"name":"\u8fdb\u9879\u7a0e\u989d (Q)",     "key":"Q_\u8fdb\u9879\u7a0e\u989d",  "type":"number"},
            {"name":"\u8fdb\u9879\u7a0e\u989d\u8f6c\u51fa (R)", "key":"R_\u8fdb\u9879\u7a0e\u989d\u8f6c\u51fa", "type":"number"},
        ]

        half = (len(fields) + 1) // 2
        grid = tk.Frame(body, bg=C["surface"])
        grid.pack(fill=tk.X)
        left_col  = tk.Frame(grid, bg=C["surface"])
        right_col = tk.Frame(grid, bg=C["surface"])
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 16))
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Column headers
        for col in (left_col, right_col):
            hdr = tk.Frame(col, bg=C["surface3"])
            hdr.pack(fill=tk.X, pady=(0, 4))
            tk.Frame(hdr, bg=C["navy"], width=3).pack(side=tk.LEFT, fill=tk.Y)
            tk.Label(hdr, text="  \u5b57\u6bb5\u540d\u79f0",
                     font=F["small_b"], bg=C["surface3"],
                     fg=C["text_3"], pady=4).pack(side=tk.LEFT)
            tk.Label(hdr, text="\u6570\u636e\u5185\u5bb9",
                     font=F["small_b"], bg=C["surface3"],
                     fg=C["text_3"]).pack(side=tk.RIGHT, padx=60)

        self.rule_inputs.clear()
        for i, f in enumerate(fields):
            parent = left_col if i < half else right_col
            fkey   = f["key"]

            row = tk.Frame(parent, bg=C["surface"],
                           highlightbackground=C["border"], highlightthickness=1)
            row.pack(fill=tk.X, pady=2)

            # Label cell — fixed width, right-aligned, light bg
            lbl_cell = tk.Frame(row, bg=C["surface2"], width=148)
            lbl_cell.pack(side=tk.LEFT, fill=tk.Y)
            lbl_cell.pack_propagate(False)
            tk.Label(lbl_cell, text=f["name"],
                     font=F["body"], fg=C["text_2"],
                     bg=C["surface2"], anchor="e",
                     padx=8, pady=6).pack(fill=tk.BOTH, expand=True)

            # Input cell
            inp_cell = tk.Frame(row, bg=C["surface"])
            inp_cell.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=6, pady=4)

            if f["type"] == "select":
                cb = ttk.Combobox(inp_cell, values=f["choices"],
                                  state="readonly", width=16,
                                  font=F["body"], style="Gov.TCombobox")
                cb.pack(side=tk.LEFT, ipady=3)
                cb.set(f["choices"][0])
                self.rule_inputs[fkey] = cb
            else:
                e = mk_entry(inp_cell, width=17)
                e.pack(side=tk.LEFT, ipady=4)
                if f["type"] == "number":
                    e.bind("<KeyRelease>", self._num_hint)
                self.rule_inputs[fkey] = e

            # Source button
            if fkey in FIELD_SOURCE_MAP:
                src_btn = tk.Label(
                    inp_cell,
                    text=" \u6765\u6e90 ",
                    font=F["tag"],
                    bg=C["surface2"],
                    fg=C["navy"],
                    relief=tk.FLAT,
                    cursor="hand2",
                    padx=4, pady=2,
                    highlightthickness=1,
                    highlightbackground=C["border_dark"],
                )
                src_btn.pack(side=tk.LEFT, padx=(6, 0))
                src_btn.bind("<Button-1>", lambda e, k=fkey: self._show_source(k))
                src_btn.bind("<Enter>", lambda e, w=src_btn: w.config(bg=C["navy"], fg="#fff"))
                src_btn.bind("<Leave>", lambda e, w=src_btn: w.config(bg=C["surface2"], fg=C["navy"]))

        # Hint below grid
        tk.Label(body,
                 text="\u25b3  \u6570\u91cf\u5b57\u6bb5\u5355\u4f4d\u4e3a\u4eba\u6c11\u5e01\u5143\uff0c\u652f\u6301\u6700\u591a\u4e24\u4f4d\u5c0f\u6570\u3002\u5c0f\u6570\u5019\u70b9\u524d\u5fc5\u987b\u4e3a\u6574\u6570\u3002",
                 font=F["small"], bg=C["surface"], fg=C["text_3"]).pack(anchor="w", pady=(10, 0))

    # --------------------------------------------------
    def _build_tax_table_card(self):
        body = self._card("\u4e09\u3001\u7a0e\u6536\u4f18\u60e0\u6838\u67e5",
                          "  \u5982\u4e0d\u6d89\u53ca\u53ef\u8df3\u8fc7\u6b64\u90e8\u5206")

        # Table header
        hdr = tk.Frame(body, bg=C["navy"])
        hdr.pack(fill=tk.X)
        col_weights = [("  \u7a0e\u6536\u9879\u76ee", 2),
                       ("\u5e94\u4eab\u4f18\u60e0\uff08\u5143\uff09", 1),
                       ("\u5df2\u4eab\u4f18\u60e0\uff08\u5143\uff09", 1),
                       ("\u672a\u4eab\u4f18\u60e0\uff08\u5143\uff09", 1)]
        for txt, wt in col_weights:
            tk.Label(hdr, text=txt, font=F["body_b"],
                     bg=C["navy"], fg="#ffffff",
                     anchor="center", padx=12, pady=8).pack(
                side=tk.LEFT, expand=(wt==1), fill=tk.X)

        for idx, item in enumerate(self.tax_items):
            rb = C["surface"] if idx % 2 == 0 else C["surface3"]
            row = tk.Frame(body, bg=rb,
                           highlightbackground=C["border"], highlightthickness=1)
            row.pack(fill=tk.X, pady=1)

            # Item name
            tk.Label(row, text=f"  {item}", font=F["body"],
                     bg=rb, fg=C["text"], anchor="w",
                     padx=8, pady=7, width=18).pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Thin separator
            tk.Frame(row, bg=C["border"], width=1).pack(side=tk.LEFT, fill=tk.Y, pady=4)

            should_e = mk_entry(row, width=16)
            should_e.pack(side=tk.LEFT, padx=10, pady=5, ipady=4)
            should_e.bind("<KeyRelease>", self._validate_num)

            tk.Frame(row, bg=C["border"], width=1).pack(side=tk.LEFT, fill=tk.Y, pady=4)

            enjoyed_e = mk_entry(row, width=16)
            enjoyed_e.pack(side=tk.LEFT, padx=10, pady=5, ipady=4)
            enjoyed_e.bind("<KeyRelease>", self._validate_num)

            tk.Frame(row, bg=C["border"], width=1).pack(side=tk.LEFT, fill=tk.Y, pady=4)

            not_e = tk.Entry(row, font=F["mono"], width=16,
                             state="readonly", relief=tk.FLAT, bd=0,
                             readonlybackground=C["surface2"],
                             fg=C["text_3"], justify="center",
                             highlightthickness=1,
                             highlightbackground=C["border"])
            not_e.pack(side=tk.LEFT, padx=10, pady=5, ipady=4)

            self.entries[item] = {
                "should": should_e, "enjoyed": enjoyed_e, "not_enjoyed": not_e
            }

        # Table footer note
        note_row = tk.Frame(body, bg=C["surface2"], pady=6)
        note_row.pack(fill=tk.X, pady=(6, 0))
        tk.Label(note_row,
                 text="  \u672a\u4eab\u4f18\u60e0 = \u5e94\u4eab\u4f18\u60e0 \u2212 \u5df2\u4eab\u4f18\u60e0\uff0c\u8d1f\u6570\u6309\u96f6\u8ba1\u7b97",
                 font=F["small"], bg=C["surface2"], fg=C["text_3"]).pack(anchor="w")

    # --------------------------------------------------
    def _build_action_bar(self):
        outer = tk.Frame(self.content_frame, bg=C["bg"], padx=20, pady=10)
        outer.pack(fill=tk.X)
        card = tk.Frame(outer, bg=C["surface"],
                        highlightbackground=C["border"], highlightthickness=1)
        card.pack(fill=tk.X)

        tk.Frame(card, bg=C["navy"], height=3).pack(fill=tk.X)

        btn_row = tk.Frame(card, bg=C["surface"], padx=20, pady=14)
        btn_row.pack(fill=tk.X)

        btns = [
            ("\u67e5\u8be2\u4f18\u60e0",         C["btn_primary"],  self.calculate_benefits,  "\u8ba1\u7b97\u7a0e\u6536\u4f18\u60e0\u672a\u4eab\u91d1\u989d"),
            ("\u89c4\u5219\u68c0\u67e5",          C["btn_success"],  self.run_rule_checks,     "\u8fd0\u884c\u5168\u90e8\u7591\u70b9\u68c0\u6d4b\u89c4\u5219"),
            ("\u6e05\u7a7a\u8868\u5355",          C["btn_warn"],     self.reset_form,          "\u6e05\u7a7a\u6240\u6709\u8f93\u5165\u5185\u5bb9"),
            ("\u9000\u51fa\u7cfb\u7edf",          C["btn_danger"],   self.exit_app,            "\u9000\u51fa\u5e94\u7528\u7a0b\u5e8f"),
        ]

        for label, color, cmd, tip in btns:
            col = tk.Frame(btn_row, bg=C["surface"])
            col.pack(side=tk.LEFT, padx=(0, 12))
            b = mk_flat_btn(col, label, color, command=cmd, width=10, padx=20, pady=10)
            b.pack()
            tk.Label(col, text=tip, font=F["small"],
                     bg=C["surface"], fg=C["text_3"]).pack(pady=(4, 0))

        # Right: system info
        info = tk.Frame(btn_row, bg=C["surface"])
        info.pack(side=tk.RIGHT)
        tk.Label(info, text="\u56fd\u5bb6\u7a0e\u52a1\u5c40  \u91d1\u7a0e\u4e09\u671f\u7cfb\u7edf",
                 font=F["small_b"], bg=C["surface"], fg=C["navy"]).pack(anchor="e")
        tk.Label(info, text="\u5185\u90e8\u5408\u89c4\u68c0\u6d4b\u5de5\u5177  \u4ec5\u4f9b\u5de5\u4f5c\u4eba\u5458\u4f7f\u7528",
                 font=F["small"], bg=C["surface"], fg=C["text_3"]).pack(anchor="e")

    # --------------------------------------------------
    def _build_statusbar(self):
        bar = tk.Frame(self.root, bg=C["navy_dark"], height=28)
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        bar.pack_propagate(False)

        tk.Frame(bar, bg=C["gold"], width=3).pack(side=tk.LEFT, fill=tk.Y)

        tk.Label(bar, textvariable=self._status_var,
                 font=F["small"], bg=C["navy_dark"],
                 fg="#99aac8", anchor="w",
                 padx=12).pack(side=tk.LEFT, fill=tk.Y)

        # Right system status
        tk.Label(bar, text="\u7cfb\u7edf\u6b63\u5e38  \u2714",
                 font=F["small_b"], bg=C["navy_dark"],
                 fg=C["ok"]).pack(side=tk.RIGHT, padx=16)

    def _set_status(self, msg):
        self._status_var.set(f"\u25b6  {msg}")

    # ====================== Popups ======================
    def _show_source(self, field_key):
        info = FIELD_SOURCE_MAP.get(field_key)
        if not info:
            return
        win, hdr = popup_base(self.root, "\u6570\u636e\u6765\u6e90\u8bf4\u660e", 660, 280)

        tk.Label(hdr, text=f"   \u6570\u636e\u6765\u6e90\u8bf4\u660e  \u2014  {info['title']}",
                 font=F["h1"], bg=C["navy_dark"], fg="#ffffff",
                 padx=12).pack(side=tk.LEFT)

        body = tk.Frame(win, bg=C["surface"], padx=28, pady=20)
        body.pack(fill=tk.BOTH, expand=True)

        # Field label row
        lbl_row = tk.Frame(body, bg=C["surface2"],
                           highlightbackground=C["border"], highlightthickness=1)
        lbl_row.pack(fill=tk.X, pady=(0, 12))
        tk.Label(lbl_row, text="  \u5b57\u6bb5", font=F["small_b"],
                 bg=C["surface2"], fg=C["text_3"], pady=5).pack(side=tk.LEFT)
        tk.Label(lbl_row, text=info["title"], font=F["body_b"],
                 bg=C["surface2"], fg=C["navy"], padx=8).pack(side=tk.LEFT)

        # Source path
        tk.Label(body, text="\u67e5\u8be2\u8def\u5f84 / \u6570\u636e\u6765\u6e90",
                 font=F["small_b"], bg=C["surface"],
                 fg=C["text_3"]).pack(anchor="w")
        tk.Frame(body, bg=C["navy"], height=1).pack(fill=tk.X, pady=(2, 8))

        tk.Label(body, text=info["source"],
                 font=F["body"], bg=C["surface"], fg=C["text"],
                 wraplength=580, justify="left", anchor="w").pack(anchor="w")

        mk_flat_btn(win, "\u786e\u8ba4", C["btn_primary"],
                    command=win.destroy, width=8).pack(pady=14)

    def _show_results(self, issues):
        win, hdr = popup_base(self.root, "\u89c4\u5219\u68c0\u67e5\u7ed3\u679c", 860, 560)

        count = len(issues)
        red_c = sum(1 for _, s in issues if s == "red")
        yel_c = sum(1 for _, s in issues if s == "yellow")

        tk.Label(hdr, text=f"   \u89c4\u5219\u68c0\u67e5\u7ed3\u679c\u62a5\u544a",
                 font=F["h1"], bg=C["navy_dark"], fg="#ffffff", padx=12).pack(side=tk.LEFT)

        # Summary badges on header right
        if count == 0:
            tk.Label(hdr, text=" \u65e0\u7591\u70b9 ", font=F["small_b"],
                     bg=C["ok"], fg="#fff", padx=6, pady=3).pack(side=tk.RIGHT, padx=16)
        else:
            if red_c:
                tk.Label(hdr, text=f" \u9ad8\u98ce\u9669 {red_c} ", font=F["small_b"],
                         bg=C["danger"], fg="#fff", padx=6, pady=3).pack(side=tk.RIGHT, padx=4)
            if yel_c:
                tk.Label(hdr, text=f" \u9700\u6838\u5b9e {yel_c} ", font=F["small_b"],
                         bg=C["warn"], fg="#fff", padx=6, pady=3).pack(side=tk.RIGHT, padx=4)

        # Body
        body = tk.Frame(win, bg=C["surface"])
        body.pack(fill=tk.BOTH, expand=True)

        # Summary strip
        summ = tk.Frame(body, bg=C["surface2"],
                        highlightbackground=C["border"], highlightthickness=1)
        summ.pack(fill=tk.X, padx=20, pady=(14, 6))
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        tk.Label(summ,
                 text=f"  \u68c0\u67e5\u65f6\u95f4\uff1a{now}    \u53d1\u73b0\u7591\u70b9\u5171 {count} \u6761",
                 font=F["small"], bg=C["surface2"], fg=C["text_2"],
                 pady=6, anchor="w").pack(side=tk.LEFT)

        if count == 0:
            ok_frame = tk.Frame(body, bg=C["ok_bg"],
                                highlightbackground=C["ok"], highlightthickness=1)
            ok_frame.pack(fill=tk.X, padx=20, pady=10)
            tk.Label(ok_frame,
                     text="  \u2714  \u672a\u53d1\u73b0\u4efb\u4f55\u7591\u70b9\uff0c\u6240\u6709\u6307\u6807\u5747\u5728\u6b63\u5e38\u8303\u56f4\u5185\u3002",
                     font=F["body_b"], bg=C["ok_bg"], fg=C["ok"],
                     pady=14, anchor="w", padx=16).pack(fill=tk.X)
        else:
            # Scrollable list
            cont = tk.Frame(body, bg=C["surface"])
            cont.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 6))

            cv = tk.Canvas(cont, bg=C["surface"], highlightthickness=0)
            sb = ttk.Scrollbar(cont, orient="vertical",
                               command=cv.yview, style="Gov.Vertical.TScrollbar")
            cv.configure(yscrollcommand=sb.set)
            sb.pack(side=tk.RIGHT, fill=tk.Y)
            cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            lf = tk.Frame(cv, bg=C["surface"])
            lf.bind("<Configure>",
                lambda e: cv.configure(scrollregion=cv.bbox("all")))
            cwin = cv.create_window((0, 0), window=lf, anchor="nw")
            cv.bind("<Configure>", lambda e: cv.itemconfig(cwin, width=e.width))

            for idx, (msg, sev) in enumerate(issues):
                if sev == "red":
                    bar_c, bg, fg, tag = C["danger"], C["danger_bg"], C["danger"], "\u9ad8\u98ce\u9669"
                else:
                    bar_c, bg, fg, tag = C["warn"], C["warn_bg"], C["warn"], "\u9700\u6838\u5b9e"

                item_frame = tk.Frame(lf, bg=bg,
                                      highlightbackground=C["border_dark"],
                                      highlightthickness=1)
                item_frame.pack(fill=tk.X, pady=3)

                # Left severity bar
                tk.Frame(item_frame, bg=bar_c, width=5).pack(side=tk.LEFT, fill=tk.Y)

                # Index number
                tk.Label(item_frame, text=f" {idx+1:02d} ",
                         font=F["small_b"], bg=bar_c, fg="#fff",
                         pady=10, padx=4).pack(side=tk.LEFT)

                # Severity tag
                tk.Label(item_frame, text=f" {tag} ",
                         font=F["tag"], bg=fg, fg="#fff",
                         padx=6, pady=3).pack(side=tk.LEFT, padx=(8, 4), pady=10)

                # Message
                tk.Label(item_frame, text=msg,
                         font=F["body_b"], bg=bg, fg=C["text"],
                         wraplength=440, justify="left",
                         anchor="w").pack(side=tk.LEFT, padx=8, pady=10, fill=tk.X, expand=True)

                # Guide button
                if msg in ISSUE_GUIDE_MAP:
                    gb = mk_flat_btn(item_frame, "\u5904\u7f6e\u6307\u5f15",
                                     C["btn_primary"], command=lambda m=msg: self._show_guide(m),
                                     width=8, padx=10, pady=6)
                    gb.pack(side=tk.RIGHT, padx=10, pady=8)

        # Close button
        foot = tk.Frame(win, bg=C["surface2"],
                        highlightbackground=C["border"], highlightthickness=1)
        foot.pack(fill=tk.X, side=tk.BOTTOM)
        mk_flat_btn(foot, "\u5173\u95ed", C["btn_neutral"],
                    command=win.destroy, width=8).pack(pady=10)

    def _show_guide(self, issue_msg):
        lines = ISSUE_GUIDE_MAP.get(issue_msg, [])
        win, hdr = popup_base(self.root, "\u7591\u70b9\u5904\u7f6e\u6307\u5f15",
                              680, min(520, 180 + len(lines)*34))

        tk.Label(hdr, text="   \u7591\u70b9\u5904\u7f6e\u6307\u5f15",
                 font=F["h1"], bg=C["navy_dark"], fg="#ffffff", padx=12).pack(side=tk.LEFT)

        body = tk.Frame(win, bg=C["surface"], padx=24, pady=18)
        body.pack(fill=tk.BOTH, expand=True)

        # Issue label
        iss_row = tk.Frame(body, bg=C["danger_bg"],
                           highlightbackground=C["danger"], highlightthickness=1)
        iss_row.pack(fill=tk.X, pady=(0, 14))
        tk.Frame(iss_row, bg=C["danger"], width=4).pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(iss_row, text=f"  {issue_msg}",
                 font=F["body_b"], bg=C["danger_bg"], fg=C["danger"],
                 wraplength=580, justify="left",
                 padx=8, pady=10).pack(side=tk.LEFT, fill=tk.X)

        # Guide steps
        tk.Label(body, text="\u5904\u7f6e\u6b65\u9aa4\u4e0e\u6838\u67e5\u8981\u70b9\uff1a",
                 font=F["h2"], bg=C["surface"],
                 fg=C["navy"]).pack(anchor="w", pady=(0, 6))
        tk.Frame(body, bg=C["navy"], height=1).pack(fill=tk.X, pady=(0, 8))

        for line in lines:
            row = tk.Frame(body, bg=C["surface"])
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=line, font=F["body"],
                     bg=C["surface"], fg=C["text"],
                     anchor="w", justify="left",
                     wraplength=580).pack(anchor="w", padx=4)

        mk_flat_btn(win, "\u5173\u95ed", C["btn_neutral"],
                    command=win.destroy, width=8).pack(pady=14)

    # ====================== Logic ======================
    def _num_hint(self, event):
        w = event.widget
        val = w.get().strip()
        if not val:
            w.config(highlightbackground=C["input_border"])
            return
        ok = bool(re.match(r'^\d+(\.\d{1,2})?$', val))
        w.config(highlightbackground=C["input_focus"] if ok else C["danger"])

    def _validate_num(self, event):
        e = event.widget
        val = e.get()
        if val and not re.match(r'^\d*\.?\d*$', val):
            e.delete(len(val)-1, tk.END)

    def _get_dec(self, key) -> Decimal:
        widget = self.rule_inputs[key]
        if isinstance(widget, ttk.Combobox):
            raise ValueError(f"{key} \u662f\u4e0b\u62c9\u6846")
        txt = widget.get().strip()
        if not txt:
            return Decimal("0")
        if not re.match(r'^\d+(\.\d{1,2})?$', txt):
            raise ValueError(f"\u300c{key}\u300d \u683c\u5f0f\u4e0d\u6b63\u786e\uff0c\u8bf7\u8f93\u5165\u6574\u6570\u6216\u4e24\u4f4d\u5c0f\u6570")
        return Decimal(txt).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    def _set_ro(self, entry, value, flag=False):
        entry.config(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, value)
        entry.config(
            fg=C["danger"] if flag else C["ok"],
            readonlybackground=C["danger_bg"] if flag else C["ok_bg"],
            highlightbackground=C["danger"] if flag else C["border"],
        )
        entry.config(state="readonly")

    # ---------- Query ----------
    def calculate_benefits(self):
        if not self.credit_code_entry.get().strip() or not self.company_name_entry.get().strip():
            messagebox.showwarning("\u63d0\u793a", "\u8bf7\u5148\u5b8c\u6210\u4e00\u3001\u4f01\u4e1a\u57fa\u672c\u4fe1\u606f\u586b\u5199")
            return
        try:
            total = 0.0
            any_diff = False
            for item in self.tax_items:
                refs = self.entries[item]
                s = float(refs["should"].get().strip() or 0)
                j = float(refs["enjoyed"].get().strip() or 0)
                diff = max(0.0, s - j)
                total += diff
                if diff > 0:
                    any_diff = True
                self._set_ro(refs["not_enjoyed"], f"{diff:,.2f}", diff > 0)

            if any_diff:
                self._set_status(f"\u67e5\u8be2\u5b8c\u6210 \u2014 \u603b\u672a\u4eab\u4f18\u60e0 \uffe5{total:,.2f} \u5143\uff0c\u5efa\u8bae\u5462\u5411\u4f01\u4e1a\u544a\u77e5")
                messagebox.showinfo("\u67e5\u8be2\u5b8c\u6210",
                    f"\u8ba1\u7b97\u5b8c\u6210\n\n\u603b\u672a\u4eab\u4f18\u60e0\u91d1\u989d\uff1a\uffe5 {total:,.2f} \u5143\n\n\u5efa\u8bae\u5c31\u672a\u4eab\u4f18\u60e0\u9879\u76ee\u5411\u4f01\u4e1a\u8fdb\u884c\u5462\u793a")
            else:
                self._set_status("\u67e5\u8be2\u5b8c\u6210 \u2014 \u6682\u65e0\u672a\u4eab\u4f18\u60e0\u9879\u76ee")
                messagebox.showinfo("\u67e5\u8be2\u5b8c\u6210", "\u8ba1\u7b97\u5b8c\u6210\uff0c\u6682\u65e0\u672a\u4eab\u4f18\u60e0\u9879\u76ee\u3002")
        except Exception as ex:
            messagebox.showerror("\u9519\u8bef", f"\u8ba1\u7b97\u5f02\u5e38\uff1a{ex}")

    # ---------- Rules ----------
    def run_rule_checks(self):
        try:
            A   = self.rule_inputs["A_\u4e3b\u8425\u884c\u4e1a"].get().strip()
            C_v = self._get_dec("C_\u8425\u4e1a\u6536\u5165")
            D   = self._get_dec("D_\u9500\u552e\u6536\u5165")
            E   = self._get_dec("E_\u6210\u672c")
            F_v = self._get_dec("F_\u9500\u552e\u8d39\u7528")
            G   = self._get_dec("G_\u7ba1\u7406\u8d39\u7528")
            H   = self._get_dec("H_\u8d22\u52a1\u8d39\u7528")
            I   = self._get_dec("I_\u5de5\u8d44\u85aa\u91d1")
            J   = self._get_dec("J_\u4e2a\u7a0e\u6263\u7f34\u5de5\u8d44\u603b\u989d")
            K   = self._get_dec("K_\u8ba1\u63d0\u6298\u65f7")
            L   = self._get_dec("L_\u5f53\u671f\u5f00\u7968\u989d\u5ea6")
            M   = self._get_dec("M_\u5f53\u671f\u53d7\u7968\u989d\u5ea6")
            N   = self._get_dec("N_\u5370\u82b1\u7a0e\u8ba1\u7a0e\u4f9d\u636e")
            O   = self._get_dec("O_\u7b80\u6613\u8ba1\u7a0e\u9500\u552e\u989d")
            P   = self._get_dec("P_\u514d\u7a0e\u9500\u552e\u989d")
            Q   = self._get_dec("Q_\u8fdb\u9879\u7a0e\u989d")
            R   = self._get_dec("R_\u8fdb\u9879\u7a0e\u989d\u8f6c\u51fa")

            yellow_keys = {
                "\u8fdb\u4e00\u6b65\u6838\u5b9e\u7eb3\u7a0e\u4eba\u662f\u5426\u5c11\u8ba1\u5370\u82b1\u7a0e\u8ba1\u7a0e\u4f9d\u636e",
                "\u5de5\u5546\u4e1a\u7eb3\u7a0e\u4eba\u6210\u672c\u8d39\u7528\u5360\u6bd4\u5f02\u5e38",
                "\u670d\u52a1\u4e1a\u7eb3\u7a0e\u4eba\u6210\u672c\u8d39\u7528\u5360\u6bd4\u5f02\u5e38",
                "\u7591\u4f3c\u591a\u5217\u5de5\u8d44\u85aa\u91d1\u652f\u51fa\u6216\u5c11\u6263\u7f34\u4e2a\u4eba\u6240\u5f97\u7a0e",
                "\u7591\u4f3c\u5c11\u8f6c\u51fa\u7528\u4e8e\u7b80\u6613\u8ba1\u7a0e\u7684\u8fdb\u9879\u7a0e\u989d",
            }
            def ws(msg): return (msg, "yellow" if msg in yellow_keys else "red")
            issues = []

            if abs(D - C_v) > Decimal("100"):
                issues.append(ws("\u540c\u671f\u7533\u62a5\u7684\u589e\u503c\u7a0e\u6536\u5165\u4e0e\u4f01\u4e1a\u6240\u5f97\u7a0e\u6536\u5165\u6709\u5dee\u5f02"))

            if C_v > 0:
                total_r = (E + F_v + G + H) / C_v
                fee_r   = (F_v + G + H) / C_v
                cost_r  = E / C_v
                if A in {"\u6279\u53d1\u96f6\u552e", "\u5236\u9020"} and total_r >= Decimal("0.70"):
                    issues.append(ws("\u5de5\u5546\u4e1a\u7eb3\u7a0e\u4eba\u6210\u672c\u8d39\u7528\u5360\u6bd4\u5f02\u5e38"))
                    if cost_r > Decimal("0.50"): issues.append(ws("\u6210\u672c\u504f\u9ad8"))
                    if fee_r  >= Decimal("0.50"): issues.append(ws("\u8d39\u7528\u504f\u9ad8"))
                if A in {"\u751f\u6d3b\u670d\u52a1", "\u4ea4\u901a\u8fd0\u8f93"} and total_r >= Decimal("0.60"):
                    issues.append(ws("\u670d\u52a1\u4e1a\u7eb3\u7a0e\u4eba\u6210\u672c\u8d39\u7528\u5360\u6bd4\u5f02\u5e38"))
                    if cost_r > Decimal("0.50"): issues.append(ws("\u6210\u672c\u504f\u9ad8"))
                    if fee_r  >= Decimal("0.50"): issues.append(ws("\u8d39\u7528\u504f\u9ad8"))

            if (E + F_v + G + H - I - K - M) > Decimal("20000"):
                issues.append(ws("\u7591\u4f3c\u672a\u53d6\u5f97\u5408\u6cd5\u6709\u6548\u51ed\u8bc1\u5217\u652f"))

            if I >= Decimal("500000") and (I - J) >= Decimal("100000"):
                issues.append(ws("\u7591\u4f3c\u591a\u5217\u5de5\u8d44\u85aa\u91d1\u652f\u51fa\u6216\u5c11\u6263\u7f34\u4e2a\u4eba\u6240\u5f97\u7a0e"))

            base = C_v + E + F_v + G - I
            if base >= Decimal("1000000") and N < base:
                issues.append(ws("\u8fdb\u4e00\u6b65\u6838\u5b9e\u7eb3\u7a0e\u4eba\u662f\u5426\u5c11\u8ba1\u5370\u82b1\u7a0e\u8ba1\u7a0e\u4f9d\u636e"))

            cands = [D, C_v, L]
            ts = max(cands) if any(x > 0 for x in cands) else Decimal("0")
            if Q > 0 and ts > 0:
                exp = (Q * (O / ts)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                if R + Decimal("0.01") < exp:
                    issues.append(ws("\u7591\u4f3c\u5c11\u8f6c\u51fa\u7528\u4e8e\u7b80\u6613\u8ba1\u7a0e\u7684\u8fdb\u9879\u7a0e\u989d"))

            count = len(issues)
            if count == 0:
                self._set_status("\u89c4\u5219\u68c0\u67e5\u5b8c\u6210 \u2014 \u672a\u53d1\u73b0\u7591\u70b9")
            else:
                red_c = sum(1 for _, s in issues if s == "red")
                self._set_status(f"\u89c4\u5219\u68c0\u67e5\u5b8c\u6210 \u2014 \u53d1\u73b0 {count} \u6761\u7591\u70b9\uff0c\u5176\u4e2d\u9ad8\u98ce\u9669 {red_c} \u6761\uff0c\u8bf7\u5c3d\u5feb\u6838\u67e5")
            self._show_results(issues)

        except ValueError as ex:
            messagebox.showerror("\u8f93\u5165\u9519\u8bef", str(ex))
        except Exception as ex:
            messagebox.showerror("\u7cfb\u7edf\u9519\u8bef", f"\u89c4\u5219\u68c0\u67e5\u5f02\u5e38\uff1a{ex}")

    # ---------- Reset ----------
    def reset_form(self):
        if not messagebox.askyesno("\u786e\u8ba4\u64cd\u4f5c",
                "\u786e\u5b9a\u8981\u6e05\u7a7a\u6240\u6709\u8f93\u5165\u5185\u5bb9\uff1f\n\u6b64\u64cd\u4f5c\u4e0d\u53ef\u64a4\u9500\u3002"):
            return
        self.credit_code_entry.delete(0, tk.END)
        self.company_name_entry.delete(0, tk.END)
        for key, widget in self.rule_inputs.items():
            if isinstance(widget, ttk.Combobox):
                widget.set(widget["values"][0])
            else:
                widget.delete(0, tk.END)
                widget.config(highlightbackground=C["input_border"])
        for item in self.tax_items:
            refs = self.entries[item]
            refs["should"].delete(0, tk.END)
            refs["enjoyed"].delete(0, tk.END)
            refs["not_enjoyed"].config(state="normal")
            refs["not_enjoyed"].delete(0, tk.END)
            refs["not_enjoyed"].config(state="readonly",
                                       fg=C["text_3"],
                                       readonlybackground=C["surface2"],
                                       highlightbackground=C["border"])
        self._set_status("\u8868\u5355\u5df2\u91cd\u7f6e \u2014 \u8bf7\u91cd\u65b0\u5f55\u5165\u4f01\u4e1a\u4fe1\u606f")
        messagebox.showinfo("\u64cd\u4f5c\u5b8c\u6210", "\u8868\u5355\u5185\u5bb9\u5df2\u5168\u90e8\u6e05\u7a7a\u3002")

    # ---------- Exit ----------
    def exit_app(self):
        if messagebox.askyesno("\u9000\u51fa\u786e\u8ba4", "\u786e\u5b9a\u8981\u9000\u51fa\u7cfb\u7edf\uff1f"):
            self.root.quit()
            self.root.destroy()


def main():
    root = tk.Tk()
    app = TaxBenefitApp(root)
    root.protocol("WM_DELETE_WINDOW", app.exit_app)
    root.mainloop()


if __name__ == "__main__":
    main()