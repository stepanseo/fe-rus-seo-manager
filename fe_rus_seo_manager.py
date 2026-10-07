# -*- coding: utf-8 -*-
"""
FE-RUS SEO Manager v1.21.4

Единое Windows-приложение:
1. Категория - получает category_id, OCFilter options и значения.
2. WordKeeper - авторизация, получение TOP-30 с продолжением после перезапуска.
3. Анализ - классификация CREATE / SKIP / REVIEW / EXISTS и конкуренты.
4. Проверка URL - проверка ожидаемых OCFilter URL до создания.
5. OCFilter - реальный direct POST в штатный addPage, без Playwright.
6. Экспорт - XLSX/CSV для внешней программы генерации контента (CSV: ; + UTF-8 BOM).

Важно:
- "Показывать в ТОП меню" всегда отправляется как menu_status=0.
- Пароли и логины сохраняются в JSON в %APPDATA%\\FE-RUS SEO Manager\\config.json.
- Ctrl+A/C/V/X работают во всех Entry.
- OCFilter CREATE запускается только после явного нажатия.
"""

import os
import re
import json
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue
import webbrowser
import types
import traceback
from pathlib import Path
from urllib.parse import urlparse, urljoin, unquote_plus
from collections import Counter

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import requests
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font

APP = "FE-RUS SEO Manager"
BASE = "https://fe-rus.ru"
ADMIN = BASE + "/admin/"
LOGIN_WK = "https://word-keeper.ru/login"
TOP30_WK = "https://word-keeper.ru/core/ajax_top30"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/152 Safari/537.36"
CFG = Path(os.environ.get("APPDATA", Path.home())) / APP / "config.json"

CREATE_ROUTE = "extension/module/ocfilter/addPage"
PAGE_ROUTE = "extension/module/ocfilter/page"
MENU_STATUS = "0"  # ВСЕГДА отключено

# Ускорение TOP-30 WordKeeper.
# 0.5 сек между запросами вместо прежних 2 сек.
WK_DELAY = 0.5
# CSV/checkpoint сохраняются пакетно, чтобы не переписывать весь накопительный
# файл после каждого запроса. При остановке/ошибке пакет принудительно сохраняется.
WK_SAVE_EVERY = 10


def norm(x):
    x = str(x or "").lower().replace("ё", "е")
    x = re.sub(r"\s+", " ", x).strip()
    return x


def canonical_value(x):
    """Нормализует значение характеристики для сопоставления строк.

    В CSV/Excel число может прийти как 1.0, а результат OCFilter как "1".
    Для сопоставления это должно считаться одним и тем же значением.
    Остальные значения сравниваются через norm().
    """
    s = norm(x)
    if not s:
        return ""
    sn = s.replace(",", ".")
    if re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", sn):
        try:
            from decimal import Decimal
            d = Decimal(sn)
            return format(d, "f").rstrip("0").rstrip(".") or "0"
        except Exception:
            pass
    return s


def result_row_key(x):
    """Стабильный ключ строки для объединения результатов OCFilter с анализом."""
    return (
        norm(x.get("keyword", "")),
        norm(x.get("filter", "")),
        canonical_value(x.get("value", "")),
    )


def slugify(text):
    tr = {
        "а":"a","б":"b","в":"v","г":"g","д":"d","е":"e","ё":"e","ж":"zh","з":"z",
        "и":"i","й":"j","к":"k","л":"l","м":"m","н":"n","о":"o","п":"p","р":"r",
        "с":"s","т":"t","у":"u","ф":"f","х":"h","ц":"ts","ч":"ch","ш":"sh","щ":"sch",
        "ъ":"","ы":"y","ь":"","э":"e","ю":"yu","я":"ya",
    }
    s = str(text or "").lower()
    s = "".join(tr.get(c, c) for c in s)
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")


def save_cfg(data):
    CFG.parent.mkdir(parents=True, exist_ok=True)
    CFG.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_cfg():
    try:
        return json.loads(CFG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def win_clipboard_get():
    """Надёжно читает Unicode-текст из системного буфера Windows.

    Tkinter clipboard_get() в некоторых сборках/EXE может не отрабатывать
    для ttk.Entry. Поэтому для Windows используем нативный WinAPI, а Tk
    оставляем как fallback.
    """
    if os.name != "nt":
        return ""
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    CF_UNICODETEXT = 13

    # ВАЖНО (баг на 64-битной Windows): без явных argtypes/restype ctypes
    # по умолчанию считает возвращаемое значение 32-битным (c_int) и ОБРЕЗАЕТ
    # реальный 64-битный хэндл, который возвращает GetClipboardData. Дальше
    # GlobalLock получает битый хэндл, тихо возвращает NULL, и функция без
    # какой-либо ошибки отдаёт пустую строку - снаружи выглядит как "вставка
    # просто не работает". Прописываем указателеразмерные типы явно.
    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype = wintypes.HANDLE
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    user32.CloseClipboard.restype = wintypes.BOOL

    if not user32.OpenClipboard(None):
        return ""
    try:
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ""
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return ""
        try:
            return ctypes.wstring_at(ptr)
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


def win_clipboard_set(text):
    """Записывает Unicode-текст в системный буфер Windows."""
    if os.name != "nt":
        return False
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    CF_UNICODETEXT = 13
    GMEM_MOVEABLE = 0x0002
    GMEM_ZEROINIT = 0x0040

    # Та же поправка на 64-битные хэндлы/указатели, что и в win_clipboard_get.
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalFree.restype = wintypes.HGLOBAL
    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.EmptyClipboard.restype = wintypes.BOOL
    user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    user32.SetClipboardData.restype = wintypes.HANDLE
    user32.CloseClipboard.restype = wintypes.BOOL

    value = str(text or "")
    data = (value + "\x00").encode("utf-16-le")
    hmem = kernel32.GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, len(data))
    if not hmem:
        return False
    ptr = kernel32.GlobalLock(hmem)
    if not ptr:
        kernel32.GlobalFree(hmem)
        return False
    try:
        ctypes.memmove(ptr, data, len(data))
    finally:
        kernel32.GlobalUnlock(hmem)
    if not user32.OpenClipboard(None):
        kernel32.GlobalFree(hmem)
        return False
    try:
        user32.EmptyClipboard()
        if not user32.SetClipboardData(CF_UNICODETEXT, hmem):
            kernel32.GlobalFree(hmem)
            return False
        hmem = None
        return True
    finally:
        user32.CloseClipboard()


def http_session():
    s = requests.Session()
    s.headers.update({"User-Agent": UA, "Accept-Language": "ru-RU,ru;q=0.9"})
    return s


def http_status_code(response):
    """Безопасно получает HTTP-код из requests.Response или словаря-обертки."""
    if hasattr(response, "status_code"):
        try:
            return int(response.status_code)
        except (TypeError, ValueError):
            return response.status_code

    if isinstance(response, dict):
        for key in ("status_code", "http_status", "status", "code"):
            value = response.get(key)
            if value not in (None, ""):
                try:
                    return int(value)
                except (TypeError, ValueError):
                    return str(value)
        nested = response.get("response")
        if nested is not None and nested is not response:
            return http_status_code(nested)

    return "UNKNOWN"


def http_response_text(response):
    """Безопасно получает текст HTTP-ответа из Response или словаря."""
    if hasattr(response, "text"):
        return str(response.text or "")

    if isinstance(response, dict):
        for key in ("text", "html", "body", "content", "response_text"):
            value = response.get(key)
            if isinstance(value, bytes):
                return value.decode("utf-8", errors="replace")
            if isinstance(value, str):
                return value
        nested = response.get("response")
        if nested is not None and nested is not response:
            return http_response_text(nested)

    return ""


def find_col(df, names):
    cols = {str(c).lower().strip(): c for c in df.columns}
    for n in names:
        if n in cols:
            return cols[n]
    for c in df.columns:
        s = str(c).lower()
        if any(n in s for n in names):
            return c
    return None


# -----------------------------------------------------------------------------
# CATEGORY / OCFilter DATA
# -----------------------------------------------------------------------------

def category_name(url):
    p = urlparse(url).path.rstrip("/").split("/")
    return norm(p[-1].replace("-", " ")) if p and p[-1] else ""


def category_page_label(html):
    """Получает название категории для поиска в OpenCart.

    H1 на региональных страницах FE-RUS часто содержит суффикс
    «в Москве». Это не часть названия категории OpenCart, поэтому
    убираем географический хвост и возвращаем несколько вариантов:
    точный H1 и очищенный вариант.
    """
    soup = BeautifulSoup(html, "html.parser")
    candidates = []

    for tag in soup.find_all(["h1", "title"]):
        txt = re.sub(r"\s+", " ", tag.get_text(" ", strip=True)).strip()
        if txt and txt not in candidates:
            candidates.append(txt)

    h1 = soup.find("h1")
    raw = ""
    if h1:
        raw = re.sub(r"\s+", " ", h1.get_text(" ", strip=True)).strip()
    elif candidates:
        raw = candidates[0]

    if not raw:
        return ""

    # Убираем типовые региональные окончания.
    cleaned = re.sub(
        r"\s+(?:в|во)\s+(?:городе\s+)?"
        r"(?:москве|санкт[- ]петербурге|спб|новосибирске|екатеринбурге|"
        r"казани|нижнем\s+новгороде|самаре|омске|ростове[- ]на[- ]дону|"
        r"челябинске|перми|уфе|краснодаре|воронеже|калуге|туле|"
        r"тюмени|барнауле|томске|кемерово|иркутске|хабаровске|"
        r"владивостоке|сургуте|эстонии)\s*$",
        "",
        raw,
        flags=re.I,
    ).strip(" -")

    return cleaned or raw


def find_category_id(html, category_url="", oc_obj=None):
    """Определяет ID именно текущей категории, а не первый category_id в HTML.

    На страницах OpenCart category_id встречается много раз: в блоках меню,
    родительских категориях, товарах и служебном JS. Поэтому брать первый
    найденный ID нельзя. Сначала ищем ID внутри данных OCFilter, затем
    оцениваем все вхождения в HTML по близости к названию/slug текущей
    категории.
    """
    category_slug = norm(urlparse(category_url).path.rstrip("/").split("/")[-1])
    category_text = norm(category_slug.replace("-", " "))
    tokens = [t for t in re.split(r"[^a-z0-9а-яё]+", category_text) if len(t) >= 3]

    # 1. Ищем category_id внутри объекта OCFilter. Это наиболее надежный источник,
    # если тема/модуль передает текущую категорию в ocFilterData.
    obj_candidates = []
    def walk_obj(x, context=""):
        if isinstance(x, dict):
            for k, v in x.items():
                kl = str(k).lower()
                if kl in {"category_id", "categoryid", "current_category_id", "currentcategoryid"}:
                    sv = str(v or "").strip()
                    if sv.isdigit() and sv != "0":
                        obj_candidates.append((sv, context + " " + str(k)))
                walk_obj(v, context + " " + str(k))
        elif isinstance(x, list):
            for v in x:
                walk_obj(v, context)

    if oc_obj is not None:
        walk_obj(oc_obj)
        if obj_candidates:
            # Если рядом есть название/slug категории, отдаем приоритет ему.
            scored = []
            for cid, ctx in obj_candidates:
                c = norm(ctx)
                score = sum(10 for t in tokens if t in c)
                scored.append((score, cid))
            scored.sort(reverse=True)
            return scored[0][1]

    # 2. Собираем ВСЕ category_id из HTML и выбираем наиболее вероятный
    # для текущей страницы, а не первое попавшееся значение.
    patterns = [
        r'"category_id"\s*:\s*"(\d+)"',
        r"'category_id'\s*:\s*'(\d+)'",
        r'data-category-id=["\'](\d+)["\']',
        r'category_id\s*=\s*["\']?(\d+)',
    ]
    matches = []
    for pat in patterns:
        for m in re.finditer(pat, html, re.I):
            cid = m.group(1)
            start = m.start()
            context = norm(html[max(0, start - 1800):min(len(html), m.end() + 1800)])
            score = 0
            if category_slug and category_slug in context:
                score += 120
            for token in tokens:
                if token in context:
                    score += 12
            # Текущая категория обычно находится рядом с breadcrumb/title/link.
            for marker, pts in (("breadcrumb", 10), ("хлеб", 10), ("category_id", 2),
                                ("current_category", 20), ("currentcategory", 20),
                                ("data-category-id", 15)):
                if marker in context:
                    score += pts
            # Служебные parent_id обычно не являются ID текущей категории.
            if "parent_category_id" in context or "parent-category-id" in context:
                score -= 15
            matches.append((score, start, cid))

    if not matches:
        return ""
    matches.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return matches[0][2]


def extract_balanced_object(html, marker):
    for m in re.finditer(marker, html, re.I):
        pos = html.find("{", m.end())
        if pos < 0 or pos - m.end() > 1000:
            continue
        depth = 0
        quote = None
        esc = False
        for i in range(pos, len(html)):
            c = html[i]
            if quote:
                if esc:
                    esc = False
                elif c == "\\":
                    esc = True
                elif c == quote:
                    quote = None
            else:
                if c in "\"'":
                    quote = c
                elif c == "{":
                    depth += 1
                elif c == "}":
                    depth -= 1
                    if depth == 0:
                        raw = html[pos:i + 1]
                        try:
                            return json.loads(raw)
                        except Exception:
                            raw2 = re.sub(r",\s*([}\]])", r"\1", raw)
                            try:
                                return json.loads(raw2)
                            except Exception:
                                pass
                        break
    return None


def parse_ocfilter_data(html):
    for marker in ["ocFilterData", "ocfilterData", "ocfData"]:
        obj = extract_balanced_object(html, marker)
        if obj is not None:
            return obj
    m = re.search(r'(\{"option_id".{0,250000})', html, re.I | re.S)
    if m:
        return extract_balanced_object(html[m.start():], "^")
    return None


def find_options(obj):
    found = []

    def walk(x):
        if isinstance(x, dict):
            if "option_id" in x and "name" in x and isinstance(x.get("values"), list):
                vals = []
                for v in x["values"]:
                    if isinstance(v, dict) and v.get("name") is not None:
                        vals.append({
                            "value_id": str(v.get("value_id", "")),
                            "name": str(v.get("name", "")).strip(),
                            "keyword": str(v.get("keyword", "")).strip(),
                            "params": str(v.get("params", "")).strip(),
                        })
                found.append({
                    "option_id": str(x.get("option_id", "")),
                    "name": str(x.get("name", "")).strip(),
                    "keyword": str(x.get("keyword", "")).strip(),
                    "type": str(x.get("type", "")).strip(),
                    "values": vals,
                })
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)

    walk(obj)
    seen = set()
    result = []
    for o in found:
        if o["option_id"] not in seen:
            seen.add(o["option_id"])
            result.append(o)
    return result


def value_alias(option, value):
    """Использует реальные keyword OCFilter, если он есть."""
    vk = str(value.get("keyword") or "").strip()
    if vk:
        return vk.strip("/")
    ok = str(option.get("keyword") or "").strip()
    prefix = ok or str(option.get("name") or "").strip()
    vv = str(value.get("name") or "").strip()
    if prefix and vv:
        return slugify(prefix) + "-" + slugify(vv)
    return slugify(vv)


def expected_url(category_url, alias):
    return category_url.rstrip("/") + "/" + alias.strip("/") + "/"


def filter_result_url(category_url, filter_keyword, value_keyword, value):
    """URL исходного результата OCFilter до создания SEO-страницы.

    Используем реальные keyword фильтра/значения, если они есть.
    Например: /tavr-alyuminievyj/diametr/40/ или /marka/ad31/.
    """
    fk = str(filter_keyword or "").strip().strip("/")
    vk = str(value_keyword or "").strip().strip("/")
    if not vk:
        vk = slugify(str(value or ""))
    if not fk or not vk:
        return ""
    return category_url.rstrip("/") + "/" + fk + "/" + vk + "/"


# -----------------------------------------------------------------------------
# WORDKEEPER
# -----------------------------------------------------------------------------

def get_csrf(html):
    soup = BeautifulSoup(html, "html.parser")
    token = soup.find("input", {"name": "csrf_token"})
    return token.get("value", "") if token else ""


def safe_sorted_strings(values):
    """Стабильная сортировка строковых значений из CSV/JSON-кэша.

    Старые версии программы могли сохранить в checkpoint/CSV смешанные
    типы (например, число и строку). Обычный sorted() в Python 3
    сравнивает их через < и падает с TypeError.
    """
    cleaned = []
    for value in values or []:
        if value is None:
            continue
        text = str(value).strip()
        if text:
            cleaned.append(text)
    return sorted(set(cleaned), key=lambda x: x.casefold())


def wordkeeper_ajax_post(ajax_session, query, region, cookie_header=""):
    """Выполняет AJAX TOP-30 через постоянную HTTP-сессию.

    Авторизационные cookies передаются как готовая строка Cookie, поэтому мы
    не используем проблемный cookie-jar авторизационной сессии WordKeeper.
    При этом отдельная ajax-сессия переиспользует TCP/TLS-соединение между
    запросами, что заметно ускоряет большой пакет TOP-30.
    """
    query = str(query or "").strip()
    region = str(region or "").strip()
    headers = {
        "User-Agent": UA,
        "Accept-Language": "ru-RU,ru;q=0.9",
        "Accept": "text/html, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://word-keeper.ru/core",
    }
    if cookie_header:
        headers["Cookie"] = cookie_header

    return ajax_session.post(
        TOP30_WK,
        data={"word": query, "region": region},
        headers=headers,
        timeout=120,
        allow_redirects=False,
    )


def wordkeeper_query_variants(category, value):
    """Формирует реальные поисковые запросы для одного значения OCFilter.

    Один value = одна каноническая SEO-кандидатная страница, но для WordKeeper
    допускается несколько вариантов поискового запроса. Это особенно важно для:
    - Ст0 / Ст 0 / сталь 0;
    - AISI 304 / AISI304 / аиси 304 / аиси304;
    - буквенно-цифровых марок с/без пробела.

    Результаты всех вариантов объединяются под одним canonical keyword.
    """
    category = re.sub(r"\s+", " ", str(category or "")).strip()
    value = re.sub(r"\s+", " ", str(value or "")).strip()
    if not category or not value:
        return []

    result = []
    seen = set()

    def add(v):
        v = re.sub(r"\s+", " ", str(v or "")).strip(" ,;\t\r\n")
        if not v:
            return
        key = norm(v)
        if key not in seen:
            seen.add(key)
            result.append(v)

    add(f"{category} {value}")
    compact = re.sub(r"\s+", "", value)

    # Сталь: Ст0 / Ст 0 / сталь 0 / ст 0.
    m = re.fullmatch(r"(?i)ст\s*(.+)", value)
    if m:
        tail = m.group(1).strip()
        if tail:
            add(f"{category} Ст {tail}")
            add(f"{category} ст {tail}")
            add(f"{category} сталь {tail}")

    # AISI может быть записано латиницей и кириллицей.
    # Это именно варианты запроса, а не разные значения OCFilter.
    m = re.fullmatch(r"(?i)(?:aisi|аиси)\s*([0-9]+[a-zа-яё]?)", value)
    if m:
        code = m.group(1)
        # Сохраняем буквенный суффикс как в исходном значении.
        add(f"{category} AISI {code}")
        add(f"{category} AISI{code}")
        add(f"{category} аиси {code}")
        add(f"{category} аиси{code}")

    # Если значение само содержит AISI/аиси внутри более сложной строки.
    m = re.search(r"(?i)(aisi|аиси)\s*([0-9]+[a-zа-яё]?)", value)
    if m:
        code = m.group(2)
        add(f"{category} AISI {code}")
        add(f"{category} AISI{code}")
        add(f"{category} аиси {code}")
        add(f"{category} аиси{code}")

    # Буквенный префикс + число: М1 -> М 1, АМг2 -> АМг 2, С355 -> С 355.
    m = re.fullmatch(r"([A-Za-zА-Яа-яЁё]+)([0-9][A-Za-zА-Яа-яЁё0-9.,/-]*)", compact)
    if m:
        add(f"{category} {m.group(1)} {m.group(2)}")

    # Начальные цифры + буквенная часть: 09Г2С -> 09 Г2С.
    m = re.fullmatch(r"([0-9]+)([A-Za-zА-Яа-яЁё].*)", compact)
    if m:
        add(f"{category} {m.group(1)} {m.group(2)}")

    # Х18Н10Т -> Х 18Н10Т.
    m = re.fullmatch(r"([A-Za-zА-Яа-яЁё]+)([0-9]+)([A-Za-zА-Яа-яЁё].*)", compact)
    if m:
        add(f"{category} {m.group(1)} {m.group(2)}{m.group(3)}")

    return result


# -----------------------------------------------------------------------------
# SERP PAGE CLASSIFICATION
# -----------------------------------------------------------------------------

SERP_CLASSIFIER_VERSION = "1.21.1"
# Параллельная классификация URL на этапе «Анализ».
# 8 потоков заметно ускоряют сетевые проверки, не создавая слишком
# агрессивную нагрузку на сайты из TOP-30.
SERP_CLASSIFY_WORKERS = 8
PRODUCT_PATH_MARKERS = (
    "/product/", "/products/", "/item/", "/goods/", "/tovar/", "/offer/",
    "/p/", "/detail/", "/produkt/", "/catalog/product/"
)
SEO_PATH_MARKERS = (
    "/category/", "/categories/", "/catalog/", "/catalogue/", "/filter/",
    "/filters/", "/marka/", "/brand/", "/tag/", "/tags/", "/razmer/",
    "/size/", "/diametr/", "/diameter/", "/collection/", "/collections/",
    "/list/", "/list-", "/shop/", "/catalogs/"
)
ARTICLE_PATH_MARKERS = (
    "/blog/", "/article/", "/articles/", "/news/", "/novosti/", "/stati/",
    "/statya/", "/journal/"
)


def _url_type_hint(url):
    path = urlparse(str(url or "")).path.lower()
    if any(x in path for x in ARTICLE_PATH_MARKERS):
        return "ARTICLE", 7, 0
    if any(x in path for x in PRODUCT_PATH_MARKERS):
        return "PRODUCT", 7, 0
    if any(x in path for x in SEO_PATH_MARKERS):
        return "SEO_PAGE", 0, 6
    # OpenCart-style product query URLs.
    q = urlparse(str(url or "")).query.lower()
    if "route=product/product" in q or "product_id=" in q:
        return "PRODUCT", 8, 0
    return "UNKNOWN", 0, 0


def _jsonld_types(soup):
    types = []
    for script in soup.find_all("script", attrs={"type": re.compile(r"application/ld\+json", re.I)}):
        raw = script.get_text(" ", strip=True)
        if not raw:
            continue
        # Нам достаточно определить тип; строгий JSON не обязателен.
        for m in re.finditer(r'"@type"\s*:\s*"?([^",}\]]+)', raw, re.I):
            types.append(m.group(1).strip().lower())
        for m in re.finditer(r'"@type"\s*:\s*\[([^\]]+)\]', raw, re.I):
            types.extend(re.findall(r'"([^"\\]+)"', m.group(1)))
    return set(types)


def _classify_serp_page(session, url, title="", snippet="", keyword="", value=""):
    """Определяет тип результата TOP-30.

    Возвращает PRODUCT / SEO_PAGE / ARTICLE / OTHER / UNKNOWN.
    Сначала используются URL-сигналы, затем HTML для неоднозначных страниц.
    """
    url = str(url or "").strip()
    title = str(title or "").strip()
    snippet = str(snippet or "").strip()
    path = urlparse(url).path.lower()
    hint, product_score, seo_score = _url_type_hint(url)
    article_score = 7 if hint == "ARTICLE" else 0
    signals = []

    if hint == "PRODUCT":
        signals.append("URL_PRODUCT")
    elif hint == "SEO_PAGE":
        signals.append("URL_CATALOG_FILTER")
    elif hint == "ARTICLE":
        signals.append("URL_ARTICLE")

    # Сильные URL-сигналы не требуют обязательной загрузки страницы.
    # Для неоднозначных URL HTML проверяем.
    need_fetch = hint == "UNKNOWN" or (hint == "SEO_PAGE" and not any(x in path for x in ("/filter/", "/marka/", "/razmer/", "/diametr/", "/list-", "/tag/")))
    fetched = False
    http_status = ""
    final_url = url
    h1 = ""
    breadcrumbs = ""
    product_links = 0
    filter_controls = 0
    pagination = False
    types = set()

    if need_fetch:
        try:
            r = session.get(url, timeout=(4, 10), allow_redirects=True)
            fetched = True
            http_status = str(r.status_code)
            final_url = r.url
            if r.status_code == 200 and r.text:
                html = r.text[:1800000]
                soup = BeautifulSoup(html, "html.parser")
                types = _jsonld_types(soup)
                h1 = soup.find("h1").get_text(" ", strip=True) if soup.find("h1") else ""
                crumb = soup.select('[class*="breadcrumb"], [class*="breadcrumbs"], nav[aria-label*="breadcrumb" i]')
                if crumb:
                    breadcrumbs = " | ".join(c.get_text(" ", strip=True) for c in crumb[:2])[:1000]

                # Product-разметка может находиться не только на карточке товара,
                # но и на категории/фильтре: например, карточки товаров внутри
                # ItemList могут содержать @type=Product, а товарные карточки
                # категории могут использовать Microdata Product. Поэтому наличие
                # Product-разметки само по себе НЕ является достаточным признаком
                # того, что вся страница является карточкой товара.
                has_jsonld_product = any(
                    t.lower() in {"product", "productgroup"}
                    for t in types
                )
                if has_jsonld_product:
                    signals.append("JSONLD_PRODUCT")

                if any(t.lower() == "article" or t.lower().endswith("article") for t in types):
                    article_score += 9; signals.append("JSONLD_ARTICLE")
                if any(t.lower() in {"itemlist", "collectionpage", "searchresults"} for t in types):
                    seo_score += 3; signals.append("JSONLD_LIST")

                # Product microdata. Как и JSON-LD Product, это только сигнал
                # наличия товарной разметки на странице, а не автоматический
                # признак типа страницы PRODUCT. На категориях такая разметка
                # также допустима для отдельных товаров из списка.
                has_microdata_product = bool(
                    soup.find(attrs={"itemtype": re.compile(r"schema\.org/(Product|ProductGroup)", re.I)})
                )
                if has_microdata_product:
                    signals.append("MICRODATA_PRODUCT")

                # Цена / SKU / артикул / корзина - сильные товарные признаки.
                visible = soup.get_text(" ", strip=True).lower()
                if re.search(r"(?:\bsku\b|\bmpn\b|артикул|код товара|product code)", visible, re.I):
                    product_score += 2; signals.append("SKU_OR_ARTICLE")
                if re.search(r"(?:в\s+корзин|добавить\s+в\s+корзин|купить|add\s+to\s+cart|buy now)", visible, re.I):
                    product_score += 3; signals.append("CART_OR_BUY")
                if re.search(r"(?:₽|руб\.?|цена|price)", visible, re.I):
                    product_score += 2; signals.append("PRICE")

                # Несколько товарных ссылок + фильтры/пагинация = каталог/SEO.
                seen_links = set()
                for a in soup.find_all("a", href=True):
                    ah = a.get("href") or ""
                    ap = urlparse(urljoin(final_url, ah)).path.lower()
                    if any(x in ap for x in PRODUCT_PATH_MARKERS):
                        seen_links.add(ap)
                    cls = " ".join(a.get("class") or []).lower()
                    if any(x in cls for x in ("product-card", "product-item", "catalog-item", "goods-item")):
                        seen_links.add(ap)
                    if len(seen_links) >= 25:
                        break
                product_links = len(seen_links)
                if product_links >= 3:
                    seo_score += 6; signals.append(f"MULTI_PRODUCT_LINKS:{product_links}")

                inputs = soup.find_all(["select", "input", "button"])
                for el in inputs:
                    blob = (str(el.get("name") or "") + " " + str(el.get("id") or "") + " " + " ".join(el.get("class") or [])).lower()
                    if any(x in blob for x in ("filter", "ocfilter", "manufacturer", "brand", "marka", "diametr", "razmer", "size")):
                        filter_controls += 1
                if filter_controls:
                    seo_score += min(5, 2 + filter_controls // 3); signals.append(f"FILTER_CONTROLS:{filter_controls}")

                pagination = bool(soup.select('[class*="pagination" i], a[rel="next"], link[rel="next"]')) or bool(re.search(r"(?:[?&]page=|/page/)", html, re.I))
                if pagination:
                    seo_score += 3; signals.append("PAGINATION")

                if re.search(r"(?:\bblog\b|стать[яи]|новост|полезн|журнал)", (h1 + " " + breadcrumbs + " " + title).lower()):
                    article_score += 3

                if h1 and any(x in h1.lower() for x in ("каталог", "лист", "труба", "сетка", "металлопрокат")) and product_links >= 2:
                    seo_score += 2; signals.append("CATALOG_H1")

                # Только после оценки структуры страницы учитываем Product-разметку
                # как слабый дополнительный товарный сигнал. Если страница уже
                # выглядит как категория/фильтр (список товаров, фильтры,
                # пагинация или ItemList), Product-разметка не должна перевешивать
                # эти признаки. Это защищает категории, где Product/Microdata
                # используется на карточках товаров.
                catalog_context = (
                    product_links >= 2
                    or filter_controls > 0
                    or pagination
                    or any(t.lower() in {"itemlist", "collectionpage", "searchresults"} for t in types)
                )
                if not catalog_context:
                    if has_jsonld_product:
                        product_score += 2
                    if has_microdata_product:
                        product_score += 1

        except Exception as e:
            signals.append("FETCH_ERROR:" + str(e)[:120])

    # Если URL уже дал сильный сигнал, но HTML не загружали, учитываем title.
    text_hint = " ".join([title, snippet, h1, breadcrumbs]).lower()
    if re.search(r"(?:артикул|sku|купить|в\s+корзин|цена|руб\.?)", text_hint, re.I):
        product_score += 1
    if re.search(r"(?:каталог|фильтр|марка|размер|диаметр|товары|лист[ыа]?|трубы)", text_hint, re.I):
        seo_score += 1

    if article_score >= max(product_score, seo_score) + 4 and article_score >= 7:
        page_type = "ARTICLE"
        confidence = min(0.99, 0.60 + article_score / 30)
    elif product_score >= 6 and product_score >= seo_score + 3:
        page_type = "PRODUCT"
        confidence = min(0.99, 0.58 + (product_score - seo_score) / 20)
    elif seo_score >= 5 and seo_score >= product_score + 3:
        page_type = "SEO_PAGE"
        confidence = min(0.99, 0.58 + (seo_score - product_score) / 20)
    elif hint in {"PRODUCT", "SEO_PAGE", "ARTICLE"}:
        page_type = hint
        confidence = 0.62
    elif product_score > seo_score and product_score >= 4:
        page_type = "PRODUCT"
        confidence = 0.55
    elif seo_score > product_score and seo_score >= 4:
        page_type = "SEO_PAGE"
        confidence = 0.55
    elif article_score >= 5:
        page_type = "ARTICLE"
        confidence = 0.52
    else:
        page_type = "UNKNOWN"
        confidence = 0.35

    if page_type not in {"PRODUCT", "SEO_PAGE", "ARTICLE"}:
        # Не удалось определить по HTML. Если путь очевидно товарный/SEO - не теряем сигнал.
        if hint == "PRODUCT":
            page_type = "PRODUCT"; confidence = max(confidence, 0.58)
        elif hint == "SEO_PAGE":
            page_type = "SEO_PAGE"; confidence = max(confidence, 0.58)
        elif hint == "ARTICLE":
            page_type = "ARTICLE"; confidence = max(confidence, 0.58)
        else:
            page_type = "OTHER"

    return {
        "url": url,
        "final_url": final_url,
        "page_type": page_type,
        "confidence": round(float(confidence), 2),
        "signals": "; ".join(dict.fromkeys(signals)),
        "http_status": http_status,
        "fetched": "1" if fetched else "0",
        "h1": h1[:500],
        "breadcrumbs": breadcrumbs[:1000],
        "product_links": product_links,
        "filter_controls": filter_controls,
        "pagination": "1" if pagination else "0",
        "classifier_version": SERP_CLASSIFIER_VERSION,
    }


def load_serp_classification_cache(path):
    path = Path(path)
    if not path.exists():
        return {}
    try:
        df = pd.read_csv(path, sep=";", encoding="utf-8-sig", dtype=str).fillna("")
        if "url" not in df.columns:
            return {}
        cache = {}
        for _, r in df.iterrows():
            if str(r.get("classifier_version", "")) != SERP_CLASSIFIER_VERSION:
                continue
            u = str(r.get("url", "")).strip()
            if u:
                cache[u.rstrip("/")] = r.to_dict()
        return cache
    except Exception:
        return {}


def save_serp_classification_cache(cache, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if not cache:
        return
    df = pd.DataFrame(list(cache.values()))
    preferred = ["url", "final_url", "page_type", "confidence", "signals", "http_status", "fetched", "h1", "breadcrumbs", "product_links", "filter_controls", "pagination", "classifier_version"]
    cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]
    df[cols].to_csv(path, sep=";", index=False, encoding="utf-8-sig", lineterminator="\n")


def parse_top30(html, keyword):
    soup = BeautifulSoup(html, "html.parser")
    excluded = {"word-keeper.ru", "yandex.ru", "yandexwebcache.net", "xtool.ru"}
    rows = []
    seen = set()
    pos = 0
    for a in soup.find_all("a"):
        href = (a.get("href") or "").strip()
        if not href.startswith("http"):
            continue
        try:
            domain = urlparse(href).netloc.lower().split(":")[0]
        except Exception:
            continue
        if domain.startswith("www."):
            domain = domain[4:]
        if domain in excluded or "/trust/" in href or "xtool.ru" in href.lower():
            continue
        u = href.rstrip("/")
        if u in seen:
            continue
        seen.add(u)
        pos += 1
        rows.append({
            "keyword": keyword,
            "position": pos,
            "domain": domain,
            "url": href,
            "title": a.get_text(" ", strip=True),
            "snippet": "",
        })
        if pos >= 30:
            break
    return rows


# -----------------------------------------------------------------------------
# OCFILTER DIRECT POST
# -----------------------------------------------------------------------------

def first_value(d, *keys):
    for k in keys:
        v = str(d.get(k) or "").strip()
        if v:
            return v
    return ""


def oc_category_terms(category):
    q = norm(category)
    result = []
    def add(v):
        v = norm(v)
        if v and v not in result:
            result.append(v)
    add(q)
    words = q.split()
    no_num = [w for w in words if not re.fullmatch(r"\d+(?:[.,]\d+)?", w)]
    add(" ".join(no_num))
    add(" ".join(w for w in no_num if w not in {"стальная", "стальной"}))
    if len(no_num) > 1:
        add(" ".join(no_num[1:]))
    important = [w for w in no_num if len(w) >= 6 and w not in {"труба", "трубы", "круглая", "круглый"}]
    add(" ".join(important))
    return result


def oc_category_score(name, requested):
    n = norm(name); q = norm(requested); score = 0
    if n == q: score += 2000
    if q and q in n: score += 800
    for word in set(q.split()):
        if word in set(n.split()): score += 100
        elif word in n: score += 25
    parts = re.split(r"\s*>\s*|›|»", str(name))
    last = norm(parts[-1])
    if last == q: score += 1000
    elif q and q in last: score += 400
    return score


def oc_category_safe(expected, found_name):
    en = norm(expected); fn = norm(found_name)
    if en == fn:
        return True
    words = {w for w in en.split() if len(w) >= 4 and w not in {"труба","трубы","круглая","круглый","стальная","стальной"}}
    leaf = norm(re.split(r"\s*>\s*|›|»", str(found_name))[-1])
    if words and all(w in leaf for w in words):
        return True
    if "жаропрочн" in en and "жаропрочн" in leaf and ("труба" in leaf or "трубы" in leaf):
        return True
    return False


def get_token_from_html(html):
    for p in [r'name=["\']user_token["\'][^>]*value=["\']([^"\']+)', r'[?&]user_token=([^&"\']+)']:
        m = re.search(p, html, re.I)
        if m: return m.group(1)
    return None


def oc_login(session, basic_user, basic_pass, oc_user, oc_pass):
    session.auth = (basic_user, basic_pass)
    r = session.get(ADMIN + "?route=common/login", timeout=30, allow_redirects=True)
    if r.status_code != 200:
        raise RuntimeError(f"OpenCart login HTTP {r.status_code}")
    soup = BeautifulSoup(r.text, "html.parser")
    form = soup.find("form")
    if not form:
        raise RuntimeError("Форма OpenCart не найдена")
    action = urljoin(r.url, form.get("action") or "?route=common/login")
    data = {}
    for inp in form.find_all("input"):
        if inp.get("name"):
            data[inp["name"]] = inp.get("value", "")
    data["username"] = oc_user
    data["password"] = oc_pass
    r = session.post(action, data=data, timeout=30, allow_redirects=True)
    token = get_token_from_html(r.text) or get_token_from_html(r.url)
    if not token:
        raise RuntimeError(f"Не удалось получить user_token. URL: {r.url}")
    return token


def oc_form_info(session, token):
    url = f"{ADMIN}?route={CREATE_ROUTE}&user_token={token}"
    r = session.get(url, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"OCFilter addPage HTTP {r.status_code}")
    langs = sorted(set(re.findall(r'name=["\']page_description\[(\d+)\]\[', r.text, re.I))) or ["1"]
    has_statusview = bool(re.search(r'name=["\']statusview["\']', r.text, re.I))
    return langs, has_statusview


def _category_seo_keywords_from_html(html):
    """Извлекает SEO keyword из admin-страницы категории/SEO URL."""
    soup = BeautifulSoup(html, "html.parser")
    out = []
    seen = set()

    def add(value):
        value = str(value or "").strip().strip("/")
        if not value:
            return
        value = re.sub(r"\s+", " ", value)
        key = norm(value)
        if key and key not in seen:
            seen.add(key)
            out.append(value)

    # Прямые input/select поля, если SEO URL хранится в форме категории.
    for el in soup.find_all(["input", "textarea"]):
        name = str(el.get("name") or "").lower()
        value = el.get("value")
        if value is None and el.name == "textarea":
            value = el.get_text(" ", strip=True)
        if value and ("keyword" in name or "seo_url" in name or "seo-keyword" in name):
            add(value)

    # В таблице Design -> SEO URL обычно есть Query + Keyword.
    for tr in soup.find_all("tr"):
        row_text = tr.get_text(" ", strip=True)
        if not row_text:
            continue
        for el in tr.find_all(["input", "textarea"]):
            name = str(el.get("name") or "").lower()
            if "keyword" in name:
                add(el.get("value") or el.get_text(" ", strip=True))
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        if cells:
            for cell in cells:
                if "category_id=" in cell.lower() or "category_id%3d" in cell.lower():
                    # Keyword чаще всего находится последней непустой ячейкой.
                    for candidate in reversed(cells):
                        if candidate and "category_id=" not in candidate.lower() and "category_id%3d" not in candidate.lower():
                            if 1 <= len(candidate) <= 200:
                                add(candidate)
                            break

    return out


def oc_category_id_by_seo_keyword(session, token, category_slug, log=lambda x: None):
    """Надежно получает category_id по ТОЧНОМУ публичному SEO slug категории.

    Это отдельный путь от autocomplete. Он нужен потому, что OpenCart
    autocomplete может вернуть похожую категорию с другим category_id.
    Для категории из URL вида /list-perforirovannyj-stal-konstruktsionnaja/
    ищем именно этот keyword в административном разделе SEO URL и извлекаем
    query=category_id=N.
    """
    slug = str(category_slug or "").strip().strip("/")
    if not slug:
        return None

    base = f"{ADMIN}?route=design/seo_url&user_token={token}"
    urls = [
        base + "&filter_keyword=" + requests.utils.quote(slug, safe=""),
        base + "&filter_keyword=" + requests.utils.quote(slug, safe="") + "&page=1",
    ]
    target_norm = norm(slug)
    target_slug = slugify(slug.replace("_", "-"))

    seen = set()
    for url in urls:
        try:
            r = session.get(url, timeout=30)
            if log:
                log(f"OpenCart SEO URL: keyword={slug!r} -> HTTP {r.status_code}")
            if r.status_code != 200:
                continue
            soup = BeautifulSoup(r.text, "html.parser")

            for tr in soup.find_all("tr"):
                row_html = str(tr)
                row_text = tr.get_text(" ", strip=True)
                if not row_text:
                    continue

                # Сначала определяем keyword строки.
                row_keywords = []
                for el in tr.find_all(["input", "textarea"]):
                    name = str(el.get("name") or "").lower()
                    if "keyword" in name:
                        value = el.get("value") or el.get_text(" ", strip=True)
                        if value:
                            row_keywords.append(str(value).strip().strip("/"))

                cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
                row_keywords.extend(cells)
                exact_keyword = False
                for kw in row_keywords:
                    kn = norm(kw.strip("/"))
                    ks = slugify(kw.strip("/").replace("_", "-"))
                    if kn == target_norm or (target_slug and ks == target_slug):
                        exact_keyword = True
                        break

                # Даже если разметка таблицы нестандартная, query обычно
                # присутствует в HTML как category_id=NNN.
                decoded = unquote_plus(unquote_plus(row_html))
                m = re.search(r"(?:category_id=|category_id%3D)(\d+)", decoded, re.I)
                if not m:
                    m = re.search(r"(?:^|[?&\s])category_id\s*[=:]\s*(\d+)", row_text, re.I)
                if not m:
                    continue
                cid = m.group(1)

                # Если keyword удалось извлечь, принимаем только точное
                # совпадение. Если фильтр OpenCart уже вернул нужный keyword,
                # допускаем строку с category_id даже при нестандартной верстке.
                if exact_keyword:
                    if cid not in seen:
                        seen.add(cid)
                        if log:
                            log(f"OpenCart SEO URL: EXACT keyword -> category_id={cid} | slug={slug}")
                        return cid

            # Дополнительный fallback: ищем category_id=... рядом с точным
            # keyword в исходном HTML, включая URL-encoded вариант.
            html_dec = unquote_plus(unquote_plus(r.text))
            pattern = re.compile(
                r"(?:keyword[^>]{0,500}?" + re.escape(slug) + r"[^>]{0,1000}?category_id=(\d+)"
                r"|category_id=(\d+)[^>]{0,1000}?keyword[^>]{0,500}?" + re.escape(slug) + r")",
                re.I | re.S,
            )
            m = pattern.search(html_dec)
            if m:
                cid = m.group(1) or m.group(2)
                if log:
                    log(f"OpenCart SEO URL: HTML fallback -> category_id={cid} | slug={slug}")
                return cid
        except Exception as e:
            if log:
                log(f"OpenCart SEO URL: ошибка keyword={slug!r}: {e}")
    return None


def oc_category_seo_keywords(session, token, category_id):
    """Читает SEO keyword конкретной категории из OpenCart admin.

    Проверяем как форму категории, так и стандартный список Design -> SEO URL.
    Возвращаем несколько найденных вариантов, потому что на сайте может быть
    несколько языковых SEO URL.
    """
    cid = str(category_id or "").strip()
    if not cid or not cid.isdigit():
        return []

    urls = [
        f"{ADMIN}?route=catalog/category/edit&user_token={token}&category_id={cid}",
        f"{ADMIN}?route=design/seo_url&user_token={token}&filter_query=category_id%3D{cid}",
        f"{ADMIN}?route=design/seo_url&user_token={token}&filter_query=category_id={cid}",
    ]
    result = []
    seen = set()
    for url in urls:
        try:
            r = session.get(url, timeout=25)
            if r.status_code != 200:
                continue
            for value in _category_seo_keywords_from_html(r.text):
                k = norm(value)
                if k not in seen:
                    seen.add(k)
                    result.append(value)
        except Exception:
            continue
    return result


def oc_resolve_category(session, token, category, category_slug="", log=lambda x: None):
    """Ищет категорию через OpenCart autocomplete с точной проверкой SEO URL.

    Критическое правило: если передан category_slug из публичного URL, сначала
    пытаемся найти autocomplete-кандидата, у которого SEO keyword РОВНО
    соответствует этому slug. Это защищает от ситуации, когда autocomplete
    возвращает похожую категорию с другим category_id (например 474 вместо 476).
    """
    url = f"{ADMIN}?route=catalog/category/autocomplete&user_token={token}"
    best = None
    candidates = {}
    variants = oc_category_terms(category)

    cleaned = re.sub(
        r"\s+(?:в|во)\s+(?:городе\s+)?"
        r"(?:москве|санкт[- ]петербурге|спб|новосибирске|екатеринбурге|"
        r"казани|нижнем\s+новгороде|самаре|омске|ростове[- ]на[- ]дону|"
        r"челябинске|перми|уфе|краснодаре|воронеже|калуге|туле|"
        r"тюмени|барнауле|томске|кемерово|иркутске|хабаровске|"
        r"владивостоке)\s*$",
        "",
        category,
        flags=re.I,
    ).strip(" -")
    if cleaned and norm(cleaned) != norm(category):
        variants = [cleaned] + variants

    seen_terms = set()
    for term in variants:
        term = str(term).strip()
        if not term or norm(term) in seen_terms:
            continue
        seen_terms.add(norm(term))
        try:
            r = session.get(url, params={"filter_name": term}, timeout=30)
            if r.status_code != 200:
                continue
            data = r.json()
            if not isinstance(data, list):
                continue

            for item in data:
                if not isinstance(item, dict):
                    continue
                cid = str(item.get("category_id") or "").strip()
                name = str(item.get("name") or "").strip()
                if not cid or cid == "0" or not name:
                    continue

                score = oc_category_score(name, cleaned or category)
                if norm(name) == norm(cleaned or category):
                    score += 5000
                leaf = norm(re.split(r"\s*>\s*|›|»", name)[-1])
                if leaf == norm(cleaned or category):
                    score += 3000

                candidate = {
                    "score": score,
                    "category_id": cid,
                    "name": name,
                    "term": term,
                }
                old = candidates.get(cid)
                if old is None or score > old["score"]:
                    candidates[cid] = candidate
                if best is None or score > best["score"]:
                    best = candidate
        except Exception:
            continue

    # 1. Самое важное: точное совпадение SEO keyword со slug публичной категории.
    if category_slug and candidates:
        target = str(category_slug).strip().strip("/").lower()
        target_decoded = re.sub(r"[^a-z0-9а-яё_-]+", "", target)
        target_slugified = slugify(target.replace("_", "-"))
        exact = []
        for cid, candidate in candidates.items():
            keywords = oc_category_seo_keywords(session, token, cid)
            if keywords:
                log(f"OpenCart SEO: category_id={cid} | {candidate['name']} | keyword={'; '.join(keywords[:5])}")
            for kw in keywords:
                k = str(kw).strip().strip("/").lower()
                k_norm = re.sub(r"[^a-z0-9а-яё_-]+", "", k)
                k_slugified = slugify(k.replace("_", "-"))
                if k == target or k_norm == target_decoded or (target_slugified and k_slugified == target_slugified):
                    exact.append(candidate)
                    break
        if exact:
            exact.sort(key=lambda x: x["score"], reverse=True)
            chosen = exact[0]
            log(f"OpenCart category resolver: EXACT SEO URL -> {chosen['category_id']} | {chosen['name']} | slug={category_slug}")
            return chosen
        log(f"OpenCart category resolver: exact SEO keyword не найден для slug={category_slug}; кандидаты={[(x['category_id'], x['name']) for x in candidates.values()]}")

        # Fallback: если парсер Design -> SEO URL не смог прочитать keyword,
        # но autocomplete вернул категорию с точным названием или точным leaf,
        # используем этот category_id. score >= 3000 означает именно такое
        # совпадение, поэтому случайную похожую категорию не выбираем.
        safe_candidates = [
            c for c in candidates.values()
            if c.get("score", 0) >= 3000
            and oc_category_safe(cleaned or category, c.get("name", ""))
        ]
        if safe_candidates:
            safe_candidates.sort(key=lambda x: x.get("score", 0), reverse=True)
            chosen = safe_candidates[0]
            log(f"OpenCart category resolver: FALLBACK EXACT NAME -> {chosen['category_id']} | {chosen['name']}")
            return chosen

        return None

    if best and best["score"] >= 3000:
        return best
    if best and oc_category_safe(cleaned or category, best["name"]):
        return best

    # Резервный путь: autocomplete в некоторых сборках OpenCart/OCFilter
    # может не вернуть категорию вообще. В этом случае ищем ее в штатном
    # административном списке категорий, где есть реальный category_id.
    # Важно: принимаем только безопасное совпадение имени, чтобы не получить
    # соседнюю/похожую категорию.
    try:
        admin_best = oc_category_from_admin_list(session, token, cleaned or category, log=log)
        if not admin_best and category_slug:
            slug_words = re.sub(r"[-_]+", " ", str(category_slug)).strip()
            if slug_words:
                admin_best = oc_category_from_admin_list(session, token, slug_words, log=log)
        if admin_best and oc_category_safe(cleaned or category, admin_best.get("name", "")):
            log(f"OpenCart category resolver: ADMIN FALLBACK -> {admin_best['category_id']} | {admin_best['name']}")
            return admin_best
    except Exception as e:
        log(f"OpenCart category resolver: ADMIN FALLBACK ошибка: {e}")

    return None

def oc_category_from_admin_list(session, token, category, log=lambda x: None):
    """Резервный поиск реального category_id через список категорий OpenCart.

    Используется, если штатный autocomplete не вернул кандидатов. Поиск идет
    по фильтру админского списка категорий, после чего проверяется точное
    соответствие имени. Случайный похожий category_id не принимается.
    """
    url = f"{ADMIN}?route=catalog/category&user_token={token}"
    best = None
    seen = set()

    for term in oc_category_terms(category):
        try:
            r = session.get(url, params={"filter_name": term}, timeout=30)
            if log:
                log(f"OpenCart category list: {term!r} -> HTTP {r.status_code}")
            if r.status_code != 200:
                continue

            soup = BeautifulSoup(r.text, "html.parser")
            candidates = []

            for tr in soup.find_all("tr"):
                row_text = tr.get_text(" ", strip=True)
                cid = ""
                for a in tr.find_all("a", href=True):
                    href = a.get("href", "")
                    if "catalog/category/edit" in href and "category_id=" in href:
                        m = re.search(r"[?&]category_id=(\d+)", href)
                        if m:
                            cid = m.group(1)
                            break
                if not cid:
                    for inp in tr.find_all("input"):
                        val = str(inp.get("value") or "").strip()
                        name = str(inp.get("name") or "")
                        if val.isdigit() and (name.startswith("selected") or name == "category_id"):
                            cid = val
                            break
                if cid and cid not in seen:
                    seen.add(cid)
                    candidates.append((cid, row_text))

            if not candidates:
                for a in soup.find_all("a", href=True):
                    href = a.get("href", "")
                    if "catalog/category/edit" not in href or "category_id=" not in href:
                        continue
                    m = re.search(r"[?&]category_id=(\d+)", href)
                    if not m:
                        continue
                    cid = m.group(1)
                    if cid in seen:
                        continue
                    seen.add(cid)
                    parent = a.parent
                    txt = parent.get_text(" ", strip=True) if parent else a.get_text(" ", strip=True)
                    candidates.append((cid, txt))

            for cid, name in candidates:
                score = oc_category_score(name, category)
                if oc_category_safe(category, name):
                    score += 2000
                candidate = {
                    "score": score,
                    "category_id": cid,
                    "name": name,
                    "term": f"admin-category-list:{term}",
                }
                if best is None or score > best["score"]:
                    best = candidate
        except Exception as e:
            if log:
                log(f"OpenCart category list: ошибка {term!r}: {e}")

    if best and oc_category_safe(category, best.get("name", "")):
        if log:
            log(f"OpenCart category resolver: ADMIN LIST -> {best['category_id']} | {best['name']}")
        return best
    return None


def resolve_public_category_id_via_admin(category_label, basic_user, basic_pass, oc_user, oc_pass, log=lambda x: None, category_slug=''):
    """Надежно определяет category_id из самой базы OpenCart.

    Публичный HTML может содержать чужие/родительские category_id (например,
    956), поэтому для категории из URL используем авторизацию OpenCart и
    штатный catalog/category/autocomplete. Возвращаем только действительно
    совпавшую категорию.
    """
    if not all(str(x or "").strip() for x in (basic_user, basic_pass, oc_user, oc_pass)):
        return None
    s = http_session()
    try:
        token = oc_login(s, basic_user, basic_pass, oc_user, oc_pass)

        # ПЕРВЫМ делом определяем category_id по точному SEO slug текущей
        # публичной категории. Это исключает ложные IDs, которые autocomplete
        # иногда возвращает для похожих категорий.
        if category_slug:
            direct_cid = oc_category_id_by_seo_keyword(s, token, category_slug, log=log)
            if direct_cid:
                log(f"OpenCart category resolver: DIRECT SEO SLUG -> {direct_cid} | slug={category_slug}")
                return str(direct_cid)

        best = oc_resolve_category(s, token, category_label, category_slug=category_slug, log=log)

        # Если H1 содержит регион или отличается от имени в базе,
        # пробуем название, восстановленное из URL slug.
        if not best and category_slug:
            slug_words = re.sub(r"[-_]+", " ", str(category_slug)).strip()
            if slug_words:
                best = oc_resolve_category(s, token, slug_words, category_slug=category_slug, log=log)

        if best and oc_category_safe(
            re.sub(r"\s+(?:в|во)\s+.*$", "", category_label, flags=re.I).strip()
            or category_label,
            best.get("name", "")
        ): 
            log(f"OpenCart category resolver: {category_label} -> {best['category_id']} | {best['name']}")
            return str(best["category_id"])
        if best:
            log(f"OpenCart resolver найден кандидат, но не прошел проверку: {best}")

        # Если autocomplete не дал точного ID, читаем административный список
        # категорий OpenCart. Это резервный путь для конкретных сборок, где
        # autocomplete может не возвращать нужную категорию.
        admin_best = oc_category_from_admin_list(s, token, category_label, log=log)
        if not admin_best and category_slug:
            slug_words = re.sub(r"[-_]+", " ", str(category_slug)).strip()
            if slug_words:
                admin_best = oc_category_from_admin_list(s, token, slug_words, log=log)
        if admin_best:
            expected_name = re.sub(r"\s+(?:в|во)\s+.*$", "", category_label, flags=re.I).strip() or category_label
            if oc_category_safe(expected_name, admin_best.get("name", "")):
                log(f"OpenCart category resolver: {category_label} -> {admin_best['category_id']} | {admin_best['name']}")
                return str(admin_best["category_id"])
    except Exception as e:
        log(f"OpenCart category resolver: ошибка {e}")
    return None


def oc_existing(session, token):
    url = f"{ADMIN}?route={PAGE_ROUTE}&user_token={token}"
    r = session.get(url, timeout=30)
    if r.status_code != 200:
        return set()
    soup = BeautifulSoup(r.text, "html.parser")
    found = set()
    for a in soup.find_all("a", href=True):
        if "editPage" not in a.get("href", ""):
            continue
        txt = a.get_text(" ", strip=True)
        if txt: found.add(txt.lower())
        m = re.search(r"(?:keyword|seo_keyword)=([^&]+)", a.get("href", ""))
        if m: found.add(m.group(1).lower())
    for tr in soup.find_all("tr"):
        txt = tr.get_text(" ", strip=True)
        if txt: found.add(txt.lower())
    return found



def oc_existing_page_records(session, token):
    """Возвращает записи существующих SEO-страниц OCFilter из списка.

    Для проверки дублей нам недостаточно проверять только предполагаемый URL.
    Старые SEO-страницы могли быть созданы с другим ЧПУ, например:
      /list-08h17t/
    вместо нового:
      /08h17t/

    Поэтому сначала забираем список OCFilter, а затем по подходящим строкам
    при необходимости открываем editPage и читаем реальный SEO keyword.
    """
    url = f"{ADMIN}?route={PAGE_ROUTE}&user_token={token}"
    r = session.get(url, timeout=30)
    if r.status_code != 200:
        return []
    soup = BeautifulSoup(r.text, "html.parser")
    records = []
    seen = set()
    for tr in soup.find_all("tr"):
        edit = None
        for a in tr.find_all("a", href=True):
            href = a.get("href", "")
            if "editPage" in href:
                edit = urljoin(r.url, href)
                break
        if not edit:
            continue
        text = tr.get_text(" ", strip=True)
        if not text:
            continue
        key = (edit, norm(text))
        if key in seen:
            continue
        seen.add(key)
        records.append({"edit_url": edit, "text": text})
    return records


def oc_read_existing_page(session, edit_url):
    """Читает реальный alias/category/params существующей OCFilter-страницы."""
    try:
        r = session.get(edit_url, timeout=30)
        if r.status_code != 200:
            return None
        soup = BeautifulSoup(r.text, "html.parser")
        def val(*names):
            for name in names:
                el = soup.find("input", {"name": name})
                if el and el.get("value") is not None:
                    return str(el.get("value")).strip()
        return {
            "alias": val("keyword", "seo_keyword"),
            "category_id": val("category_id"),
            "params": val("params"),
            "name": val("page_description[1][name]", "name"),
            "title": val("page_description[1][title]", "title"),
            "url": r.url,
        }
    except Exception:
        return None


def oc_find_existing_seo_page(session, token, row, records, category_url):
    """Ищет существующую SEO-страницу после 404.

    Сначала проверяем наиболее вероятные старые ЧПУ напрямую - это быстро и
    позволяет поймать типовой случай FE-RUS:
        /08h17t/
        /list-08h17t/

    Только если они не найдены, используем список OCFilter и открываем
    ограниченное число наиболее подходящих editPage.
    """
    value = str(row.get("value", "") or "").strip()
    value_kw = str(row.get("value_keyword", "") or "").strip()
    filter_kw = str(row.get("filter_keyword", "") or "").strip()
    if not value and not value_kw:
        return None

    raw_values = [value_kw, value]
    slugs = []
    for raw in raw_values:
        if raw:
            a = slugify(raw).strip("-")
            if a and a not in slugs:
                slugs.append(a)

    # Типовые старые варианты ЧПУ.
    aliases = []
    for slug in slugs:
        for a in (
            slug,
            f"list-{slug}",
            f"marka-{slug}",
            f"list-marka-{slug}",
        ):
            if a not in aliases:
                aliases.append(a)

    if filter_kw:
        fk = slugify(filter_kw).strip("-")
        for slug in slugs:
            for a in (
                f"{fk}-{slug}",
                f"list-{fk}-{slug}",
            ):
                if a not in aliases:
                    aliases.append(a)

    # Не более 10 прямых URL-проверок.
    for alias in aliases[:10]:
        candidate = expected_url(category_url, alias)
        try:
            code, loc, final_code, final_url, redirected = check_public_url(session, candidate)
        except Exception:
            continue
        if code in (200, 301):
            return {
                "alias": alias,
                "category_id": str(row.get("category_id", "") or ""),
                "params": "",
                "name": "",
                "source_text": "DIRECT_ALTERNATE_URL",
                "edit_url": "",
                "existing_url": final_url if redirected else candidate,
            }

    # Затем используем уже загруженный список OCFilter.
    value_n = norm(value)
    value_kw_n = norm(value_kw)
    category_n = norm(row.get("category", ""))
    expected_cid = str(row.get("category_id", "") or "").strip()
    filter_n = norm(filter_kw)

    candidates = []
    cat_words = [w for w in category_n.split() if len(w) >= 4]

    for rec in records:
        txt = norm(rec.get("text", ""))
        href = str(rec.get("edit_url", "") or "")
        hay = txt + " " + norm(href)

        if not ((value_n and value_n in hay) or (value_kw_n and value_kw_n in hay)):
            continue

        score = 1000
        score += sum(40 for w in cat_words if w in hay)
        if filter_n and filter_n in hay:
            score += 100
        if value_kw_n and value_kw_n in hay:
            score += 200
        candidates.append((score, rec))

    candidates.sort(key=lambda z: z[0], reverse=True)

    # Не открываем сотни editPage. Максимум 3 кандидата.
    for _, rec in candidates[:3]:
        page = oc_read_existing_page(session, rec["edit_url"])
        if not page:
            continue

        alias = str(page.get("alias") or "").strip().strip("/")
        if not alias:
            continue

        pcid = str(page.get("category_id") or "").strip()
        if expected_cid and pcid and expected_cid != pcid:
            continue

        old_url = expected_url(category_url, alias)

        try:
            code, loc, final_code, final_url, redirected = check_public_url(session, old_url)
        except Exception:
            continue

        if code in (200, 301):
            return {
                "alias": alias,
                "category_id": pcid,
                "params": page.get("params", ""),
                "name": page.get("name", ""),
                "source_text": rec.get("text", ""),
                "edit_url": rec.get("edit_url", ""),
                "existing_url": final_url if redirected else old_url,
            }

    return None

def check_public_url(session, url):
    """Проверяет исходный URL без автоматического скрытия 301.

    Правила:
    200 = EXISTS
    301 = EXISTS_301, даже если Location ведёт дальше на 200/404.
    404 = свободный кандидат, после чего можно искать старое ЧПУ.
    Таймауты короткие, чтобы один зависший URL не блокировал весь анализ.
    """
    url = str(url or "").strip()
    if not url:
        raise RuntimeError("Пустой URL")

    r = session.get(
        url,
        timeout=(5, 12),
        allow_redirects=False,
    )
    initial = r.status_code
    location = r.headers.get("Location", "") or ""
    final_url = r.url
    final_status = initial
    redirected = 300 <= initial < 400

    if redirected and location:
        redirect_url = urljoin(url, location)
        fr = session.get(
            redirect_url,
            timeout=(5, 12),
            allow_redirects=True,
        )
        final_url = fr.url
        final_status = fr.status_code

    return initial, location, final_status, final_url, redirected

def alias_known(existing, alias):
    a = alias.lower().strip()
    return any(a == x or a in x for x in existing)


def fallback_alias(alias, category):
    c = slugify(category)
    return f"{alias}-{c}" if c else f"{alias}-2"


def choose_alias(existing, base_alias, category):
    if not alias_known(existing, base_alias):
        return base_alias
    fb = fallback_alias(base_alias, category)
    if not alias_known(existing, fb):
        return fb
    for n in range(2, 100):
        a = f"{fb}-{n}"
        if not alias_known(existing, a):
            return a
    raise RuntimeError(f"Не удалось подобрать свободный alias: {base_alias}")


def post_ocfilter_page(session, token, row, category_id, language_ids, has_statusview, alias):
    category = first_value(row, "category")
    flt = first_value(row, "filter")
    value = first_value(row, "value")
    params = f"{flt}/{value}"
    keyword = alias
    title = first_value(row, "meta_title", "title", "keyword")
    meta_title = first_value(row, "meta_title", "title", "keyword")
    meta_description = first_value(row, "meta_description")
    meta_keyword = first_value(row, "meta_keyword", "keyword")
    description = first_value(row, "description")
    button = first_value(row, "button", "value")
    name = first_value(row, "name", "keyword")
    data = [
        ("category_id", str(category_id)),
        ("path", category),
        ("params", params),
        ("keyword", keyword),
        ("menu_status", MENU_STATUS),
        ("sort_order", "1"),
        ("status", "1"),
    ]
    if has_statusview:
        data.append(("statusview", "1"))
    for lang in language_ids:
        p = f"page_description[{lang}]"
        data.extend([
            (f"{p}[title]", title), (f"{p}[button]", button), (f"{p}[name]", name),
            (f"{p}[description]", description), (f"{p}[meta_title]", meta_title),
            (f"{p}[meta_description]", meta_description), (f"{p}[meta_keyword]", meta_keyword),
            (f"link_category[{lang}][title]", ""), (f"link_category[{lang}][link]", ""),
            (f"link_category[{lang}][product]", "0"),
        ])
    url = f"{ADMIN}?route={CREATE_ROUTE}&user_token={token}"
    r = session.post(url, data=data, timeout=30, allow_redirects=True)
    if "route=extension/module/ocfilter/page" in r.url:
        return True, "CREATED"
    text = re.sub(r"\s+", " ", BeautifulSoup(r.text, "html.parser").get_text(" ", strip=True))
    if "Данный SEO псевдоним уже используется" in text:
        return False, "SEO_ALIAS_EXISTS"
    if "Проверьте форму на наличие ошибок" in text:
        return False, text[-1500:]
    return False, f"Неожиданный ответ: {r.url} | {text[-1000:]}"


def matching_value_variants(value):
    """Безопасные варианты значения для сопоставления URL/H1 конкурента.

    Включает реальные варианты транслитерации и записи AISI/аиси, но не считает
    разные марки разными значениями.
    """
    raw = norm(value).replace("×", "x")
    vals = set()
    if raw:
        vals.add(raw)
        vals.add(slugify(raw))
        vals.add(slugify(raw).replace("-", ""))

    # AISI <-> аиси.
    m = re.fullmatch(r"(?i)(aisi|аиси)\s*([0-9]+[a-zа-яё]?)", raw)
    if m:
        code = m.group(2)
        for prefix in ("aisi", "аиси"):
            vals.add(f"{prefix} {code}")
            vals.add(f"{prefix}{code}")
            vals.add(slugify(f"{prefix} {code}"))
            vals.add(slugify(f"{prefix} {code}").replace("-", ""))

    # Варианты транслитерации. Генерируем только ограниченный набор.
    base = slugify(raw)
    variants = {base}
    substitutions = [
        ("ya", ("ja", "ia")),
        ("yu", ("ju", "iu")),
        ("kh", ("h", "x")),
        ("h", ("kh", "x")),
        ("ts", ("c",)),
        ("ch", ("c",)),
        ("sch", ("shch",)),
        ("zh", ("j",)),
        ("j", ("i", "y")),
        ("i", ("j",)),
    ]
    for src, alts in substitutions:
        current = list(variants)
        for s in current:
            if src not in s:
                continue
            for alt in alts:
                variants.add(s.replace(src, alt, 1))
                if len(variants) >= 64:
                    break
            if len(variants) >= 64:
                break
        if len(variants) >= 64:
            break
    vals.update(variants)
    vals.update(v.replace("-", "") for v in variants)
    return {str(v).lower() for v in vals if str(v).strip()}


def value_in_text(value, text):
    if not text or value is None:
        return False
    variants = matching_value_variants(value)
    original = norm(text).replace("×", "x").replace("х", "x")
    slug = slugify(norm(text).replace("×", "x"))
    compact = re.sub(r"[^a-z0-9а-я]+", "", original)
    compact_slug = re.sub(r"[^a-z0-9]+", "", slug)
    for v in variants:
        vv = re.sub(r"[^a-z0-9а-я]+", "", str(v).lower())
        if not vv:
            continue
        if vv.isdigit():
            if re.search(rf"(?<!\d){re.escape(vv)}(?!\d)", original) or re.search(rf"(?<!\d){re.escape(vv)}(?!\d)", slug):
                return True
        elif vv in compact or vv in compact_slug:
            return True
    # Для буквенно-цифровых значений с пробелами: ст 0 == ст0.
    vt = set(re.findall(r"[a-zа-я0-9]+", norm(value)))
    tt = set(re.findall(r"[a-zа-я0-9]+", original))
    if len(vt) > 1 and vt.issubset(tt):
        return True
    return False


def promotion_type_for_counts(product_domains, seo_domains, article_domains, unknown_domains):
    """Итоговый тип продвижения для одного поискового запроса.

    Считаем независимые домены, а не количество URL. Это защищает от ситуации,
    когда один сайт отдаёт много карточек товара и искусственно перевешивает
    интент.
    """
    p = int(product_domains or 0)
    s = int(seo_domains or 0)
    total = p + s
    if total == 0:
        if article_domains:
            return "ИНФОРМАЦИОННАЯ"
        return "НЕДОСТАТОЧНО ДАННЫХ"
    if p >= 2 and p >= s * 1.5 and p > s:
        return "ТОВАР"
    if s >= 2 and s >= p * 1.5 and s > p:
        return "SEO-СТРАНИЦА"
    if p >= 2 and s == 0:
        return "ТОВАР"
    if s >= 2 and p == 0:
        return "SEO-СТРАНИЦА"
    return "СМЕШАННЫЙ"


# -----------------------------------------------------------------------------
# PIPELINE
# -----------------------------------------------------------------------------

class Pipeline:
    def __init__(self, log=lambda x: None):
        self.log = log

    def category(self, url):
        if not url.startswith("http"):
            url = "https://" + url
        s = http_session()
        r = s.get(url, timeout=40)
        r.raise_for_status()
        obj = parse_ocfilter_data(r.text)
        if obj is None:
            raise RuntimeError("В HTML не найден ocFilterData.")
        opts = find_options(obj)
        cat = category_name(r.url)
        page_label = category_page_label(r.text)
        # Публичный HTML используем только как fallback. Для точного ID
        # сначала обращаемся к OpenCart, если сохранены обе авторизации.
        cid = find_category_id(r.text, r.url, obj)
        cfg = load_cfg()
        ocfg = cfg.get("ocfilter", {}) if isinstance(cfg.get("ocfilter", {}), dict) else {}
        basic_user = cfg.get("HTTP Basic login", ocfg.get("HTTP Basic login", ""))
        basic_pass = cfg.get("HTTP Basic password", ocfg.get("HTTP Basic password", ""))
        oc_user = cfg.get("OpenCart login", ocfg.get("OpenCart login", ""))
        oc_pass = cfg.get("OpenCart password", ocfg.get("OpenCart password", ""))
        if page_label and all(str(x or "").strip() for x in (basic_user, basic_pass, oc_user, oc_pass)):
            page_slug = urlparse(r.url).path.rstrip("/").split("/")[-1]
            exact_cid = resolve_public_category_id_via_admin(
                page_label, basic_user, basic_pass, oc_user, oc_pass, self.log,
                category_slug=page_slug,
            )
            if exact_cid:
                cid = exact_cid
            else:
                # Если авторизация есть, но точный ID через OpenCart не найден,
                # не доверяем случайному category_id из публичного HTML.
                cid = ""
                self.log("Category ID: точный ID через OpenCart не найден - публичный ID не используется")
        # Для WordKeeper используем русское название из H1, а не slug URL.
        # Например: «лист рифленый оцинкованный ст0», а не
        # «list riflenyj stal otsinkovannaya Ст0».
        search_category = page_label or cat
        rows = []
        for o in opts:
            for v in o["values"]:
                if not str(v.get("name", "")).strip():
                    continue
                alias = value_alias(o, v)
                rows.append({
                    "keyword": f"{search_category} {v['name']}",
                    "category": search_category,
                    "category_id": cid,
                    "filter": o["name"],
                    "filter_keyword": o.get("keyword", ""),
                    "option_id": o["option_id"],
                    "value": v["name"],
                    "value_id": v["value_id"],
                    "params": v["params"],
                    "value_keyword": v.get("keyword", ""),
                    "alias": alias,
                    "target_url": expected_url(r.url, alias),
                    "status": "UNVERIFIED",
                    "reason": "требует TOP-30 + проверки OCFilter",
                    "competitor_1": "", "competitor_2": "", "competitor_3": "", "competitor_4": "", "competitor_5": "",
                })
        self.log(f"Категория: {cat}")
        self.log(f"Название категории (H1): {page_label or 'не найдено'}")
        self.log(f"WordKeeper категория: {search_category}")
        self.log(f"Category ID: {cid or 'не найден'}")
        self.log(f"Фильтров: {len(opts)}")
        self.log(f"Значений: {len(rows)}")
        return {"url": r.url, "category": search_category, "category_slug": urlparse(r.url).path.rstrip("/").split("/")[-1], "search_category": search_category, "category_id": cid, "options": opts, "rows": rows}

    def merge_analysis(self, res, topdf, progress=None, stop_event=None, cache_path=None):
        """Сопоставляет TOP-30 с OCFilter и определяет тип продвижения.

        Для каждого URL TOP-30 определяется тип страницы:
        PRODUCT / SEO_PAGE / ARTICLE / OTHER.
        Затем по независимым доменам определяется promotion_type:
        ТОВАР / SEO-СТРАНИЦА / СМЕШАННЫЙ / ИНФОРМАЦИОННАЯ / НЕДОСТАТОЧНО ДАННЫХ.

        CREATE разрешается только для SEO-СТРАНИЦЫ и только при наличии
        минимум двух независимых подтвержденных SEO-конкурентов, которые
        действительно соответствуют текущему value.
        """
        if topdf is None or topdf.empty:
            for x in res["rows"]:
                x["promotion_type"] = "НЕДОСТАТОЧНО ДАННЫХ"
                x["status"] = "REVIEW"
                x["reason"] = "TOP-30 не загружен"
            return res

        kc = find_col(topdf, ["keyword", "query", "запрос"])
        uc = find_col(topdf, ["url", "link", "href"])
        dc = find_col(topdf, ["domain", "host"])
        if not kc or not uc:
            raise RuntimeError("WordKeeper CSV должен содержать keyword и url.")

        category = str(res.get("category") or "").strip()
        category_slug = slugify(category)
        category_slug_variants = [v for v in matching_value_variants(category) if re.fullmatch(r"[a-z0-9-]+", str(v)) and len(str(v)) >= 4]
        category_tokens = [t for t in category_slug.split("-") if len(t) >= 3]
        generic = {
            "truba","truby","list","lenta","krug","krugi","kruglyi","kruglaya",
            "kvadrat","kvadratnyi","ugolok","shveller","dvutavr","prutok",
            "shestigrannik","provoloka","polosa","stal","stalnoi","stalnaya",
            "alyum","alyuminievyi","alyuminievyy","nerzh","nerzhaveika"
        }

        def row_text(rr):
            vals = []
            for col in topdf.columns:
                name = str(col).lower()
                if any(k in name for k in ("title", "name", "snippet", "description", "заголов", "опис")):
                    v = str(rr.get(col, "") or "")
                    if v:
                        vals.append(v)
            return " ".join(vals)

        def score_url(url, title, keyword):
            path_slug = re.sub(r"[^a-z0-9а-яё]+", "-", urlparse(url).path.lower()).strip("-")
            score = 0
            if any(v and v in path_slug for v in category_slug_variants):
                score += 100
            score += sum(20 for t in category_tokens if t in path_slug)
            qslug = slugify(keyword)
            qtokens = [t for t in qslug.split("-") if len(t) >= 3 and t not in generic]
            score += min(sum(1 for t in qtokens if t in path_slug), 5) * 8
            if title:
                ts = slugify(title)
                score += min(sum(1 for t in category_tokens if t in ts), 3) * 10
            return score

        # Загружаем постоянный кэш классификации URL.
        if cache_path is None:
            cache_path = Path(os.getcwd()) / "seo_serp_page_classification.csv"
        cache = load_serp_classification_cache(cache_path)
        session = http_session()
        page_map = {}

        # Собираем только URL, реально относящиеся к текущим canonical keywords.
        # Помимо canonical keyword учитываем source_query и все актуальные варианты
        # WordKeeper. Это важно для накопительного CSV, созданного старой версией
        # программы, где keyword мог быть транслитерированным.
        canonical_keywords = [str(x.get("keyword", "")).strip() for x in res.get("rows", []) if str(x.get("keyword", "")).strip()]
        accepted_queries = set()
        for x in res.get("rows", []):
            canonical = str(x.get("keyword", "")).strip()
            if not canonical:
                continue
            accepted_queries.add(norm(canonical))
            for qv in (wordkeeper_query_variants(
                res.get("search_category") or res.get("category") or "",
                x.get("value", ""),
            ) or [canonical]):
                accepted_queries.add(norm(qv))

        keyword_norm_series = topdf[kc].astype(str).map(norm)
        relevant = topdf[keyword_norm_series.isin(accepted_queries)].copy()
        if "source_query" in topdf.columns:
            source_norm_series = topdf["source_query"].astype(str).map(norm)
            relevant = pd.concat([
                relevant,
                topdf[source_norm_series.isin(accepted_queries)].copy(),
            ], ignore_index=False).drop_duplicates()
        if relevant.empty:
            # Совместимость с CSV, где keyword/source_query содержат дополнительные
            # пробелы или регистр.
            relevant = topdf[
                keyword_norm_series.apply(lambda z: any(q in z or z in q for q in accepted_queries if q))
            ].copy()
            if "source_query" in topdf.columns:
                source_norm_series = topdf["source_query"].astype(str).map(norm)
                extra = topdf[
                    source_norm_series.apply(lambda z: any(q in z or z in q for q in accepted_queries if q))
                ].copy()
                relevant = pd.concat([relevant, extra], ignore_index=False).drop_duplicates()

        if progress:
            progress(f"TOP-30: найдено строк текущей категории/вариантов запроса: {len(relevant)}")

        unique_urls = []
        seen_urls = set()
        for _, rr in relevant.iterrows():
            u = str(rr.get(uc, "")).strip()
            if not u.startswith("http"):
                continue
            key = u.rstrip("/")
            if key in seen_urls:
                continue
            seen_urls.add(key)
            unique_urls.append((u, str(rr.get("title", "")), str(rr.get("snippet", ""))))

        total_urls = len(unique_urls)
        new_classified = 0

        # Сначала мгновенно забираем всё из кэша, а в сеть отправляем только
        # действительно новые URL. Раньше новые URL проверялись строго
        # последовательно, поэтому десятки медленных сайтов могли растянуть
        # этап анализа на многие минуты.
        pending_urls = []
        for u, title, snippet in unique_urls:
            key = u.rstrip("/")
            if key in cache:
                page_map[key] = cache[key]
            else:
                pending_urls.append((u, title, snippet))

        if progress:
            progress(
                f"Классификация TOP-30: уникальных URL {total_urls} | "
                f"кэшировано {total_urls - len(pending_urls)} | "
                f"новых {len(pending_urls)} | потоков {SERP_CLASSIFY_WORKERS}"
            )

        # requests.Session не является гарантированно thread-safe, поэтому
        # каждый поток получает собственную Session и переиспользует её для
        # всех своих URL. Это одновременно даёт параллельность и keep-alive.
        thread_local = threading.local()

        def classify_one(task):
            u, title, snippet = task
            sess = getattr(thread_local, "session", None)
            if sess is None:
                sess = http_session()
                thread_local.session = sess
            item = _classify_serp_page(sess, u, title, snippet)
            return u.rstrip("/"), item

        if pending_urls:
            max_workers = min(SERP_CLASSIFY_WORKERS, len(pending_urls))
            executor = ThreadPoolExecutor(max_workers=max_workers)
            futures = {executor.submit(classify_one, task): task for task in pending_urls}
            try:
                for future in as_completed(futures):
                    if stop_event is not None and stop_event.is_set():
                        for f in futures:
                            if not f.done():
                                f.cancel()
                        break
                    u, item = future.result()
                    cache[u] = item
                    page_map[u] = item
                    new_classified += 1
                    if progress:
                        progress(
                            f"Классификация URL: {new_classified}/{len(pending_urls)} "
                            f"(всего {total_urls}) | {item.get('page_type')} | {u}"
                        )
                    # Сохраняем регулярно, чтобы остановка/сбой не потеряли
                    # уже завершённые параллельные проверки.
                    if new_classified % 20 == 0:
                        save_serp_classification_cache(cache, cache_path)
            finally:
                executor.shutdown(wait=True, cancel_futures=True)

        save_serp_classification_cache(cache, cache_path)

        # Добавляем классификацию к исходному TOP-30 DataFrame для полного XLSX.
        classification_cols = ["page_type","page_confidence","page_signals","page_http_status","page_fetched","page_h1","page_breadcrumbs"]

        # Pandas 2.x может прочитать уже существующие колонки как StringDtype.
        # В таком DataFrame запись числового confidence (например 0.35) через
        # .at[] вызывает:
        # "Invalid value '0.35' for dtype 'str'".
        # Для результатов классификации храним все поля как обычный object/text.
        for col in classification_cols:
            if col not in topdf.columns:
                topdf[col] = pd.Series([""] * len(topdf), index=topdf.index, dtype=object)
            else:
                topdf[col] = topdf[col].astype(object)

        def classification_cell(value):
            if value is None:
                return ""
            try:
                if pd.isna(value):
                    return ""
            except Exception:
                pass
            return str(value)

        for idx, rr in topdf.iterrows():
            u = str(rr.get(uc, "")).strip().rstrip("/")
            info = page_map.get(u) or cache.get(u)
            if not info:
                continue
            topdf.at[idx, "page_type"] = classification_cell(info.get("page_type", ""))
            topdf.at[idx, "page_confidence"] = classification_cell(info.get("confidence", ""))
            topdf.at[idx, "page_signals"] = classification_cell(info.get("signals", ""))
            topdf.at[idx, "page_http_status"] = classification_cell(info.get("http_status", ""))
            topdf.at[idx, "page_fetched"] = classification_cell(info.get("fetched", ""))
            topdf.at[idx, "page_h1"] = classification_cell(info.get("h1", ""))
            topdf.at[idx, "page_breadcrumbs"] = classification_cell(info.get("breadcrumbs", ""))

        # Повторяем строки анализа.
        for x in res["rows"]:
            canonical = str(x.get("keyword", "")).strip()
            accepted_for_row = {norm(canonical)}
            for qv in (wordkeeper_query_variants(
                res.get("search_category") or res.get("category") or "",
                x.get("value", ""),
            ) or [canonical]):
                accepted_for_row.add(norm(qv))

            rows = topdf[topdf[kc].astype(str).map(norm).isin(accepted_for_row)].copy()
            if "source_query" in topdf.columns:
                rows = pd.concat([
                    rows,
                    topdf[topdf["source_query"].astype(str).map(norm).isin(accepted_for_row)].copy(),
                ], ignore_index=False).drop_duplicates()
            if rows.empty:
                rows = topdf[
                    topdf[kc].astype(str).map(norm).apply(
                        lambda z: any(q in z or z in q for q in accepted_for_row if q)
                    )
                ].copy()
                if "source_query" in topdf.columns:
                    rows = pd.concat([
                        rows,
                        topdf[topdf["source_query"].astype(str).map(norm).apply(
                            lambda z: any(q in z or z in q for q in accepted_for_row if q)
                        )].copy(),
                    ], ignore_index=False).drop_duplicates()

            value = x.get("value", "")
            filter_urls = []
            filter_domains = set()
            landing_urls = []
            landing_domains = set()
            rejected_filter_urls = []
            product_urls = []
            product_domains = set()
            seo_urls = []
            seo_domains = set()
            article_urls = []
            article_domains = set()
            other_urls = []
            unknown_domains = set()
            fe_filter = False

            seen_result_urls = set()
            for _, rr in rows.iterrows():
                u = str(rr.get(uc, "")).strip()
                if not u.startswith("http"):
                    continue
                uk = u.rstrip("/")
                if uk in seen_result_urls:
                    continue
                seen_result_urls.add(uk)
                d = (str(rr.get(dc, "")) if dc else urlparse(u).netloc).lower().split(":")[0].removeprefix("www.")
                if d in {"word-keeper.ru", "yandex.ru"} or "xtool.ru" in u.lower():
                    continue

                info = page_map.get(uk) or cache.get(uk) or {}
                ptype = str(info.get("page_type") or "UNKNOWN")
                title = row_text(rr)
                page_text = " ".join([
                    u,
                    title,
                    str(info.get("h1", "")),
                    str(info.get("breadcrumbs", "")),
                ])
                is_our = d == "fe-rus.ru" or d.endswith(".fe-rus.ru")
                path = urlparse(u).path.lower()
                is_filter_url = any(z in path for z in (
                    "/filter/", "/diametr/", "/diameter/", "/razmer/", "/size/",
                    "filter-", "diametr-", "diameter-", "razmer-", "size-", "/marka/"
                )) or ptype == "SEO_PAGE"
                value_match = value_in_text(value, page_text)

                if ptype == "PRODUCT":
                    product_urls.append(u)
                    product_domains.add(d)
                    continue
                if ptype == "ARTICLE":
                    article_urls.append(u)
                    article_domains.add(d)
                    continue
                if ptype == "SEO_PAGE":
                    seo_urls.append(u)
                    seo_domains.add(d)
                    if value_match:
                        landing_urls.append(u)
                        landing_domains.add(d)
                        if is_filter_url:
                            if is_our:
                                fe_filter = True
                            else:
                                filter_urls.append((score_url(u, title, x.get("keyword", "")), d, u))
                    continue

                if ptype in {"UNKNOWN", "OTHER"}:
                    other_urls.append(u)
                    unknown_domains.add(d)
                    # Старый URL-only сигнал оставляем как fallback, но не считаем его
                    # сильным конкурентом без подтверждения типа страницы.
                    if is_filter_url and value_match and not is_our:
                        filter_urls.append((score_url(u, title, x.get("keyword", "")), d, u))
                    elif value_match and not is_our:
                        landing_urls.append(u)
                        landing_domains.add(d)

            # Один домен = один сигнал.
            filter_urls.sort(key=lambda z: (-z[0], z[1], z[2]))
            unique_filter_urls = []
            for score, d, u in filter_urls:
                if d not in filter_domains:
                    filter_domains.add(d)
                    unique_filter_urls.append(u)

            # SEO URL без обязательного exact value участвуют в определении интента,
            # но CREATE даём только по exact-value конкурентам.
            unique_seo_urls = []
            seen_seo_domains = set()
            for u in seo_urls:
                d = urlparse(u).netloc.lower().split(":")[0].removeprefix("www.")
                if d not in seen_seo_domains:
                    seen_seo_domains.add(d)
                    unique_seo_urls.append(u)

            for u in unique_filter_urls:
                if u not in landing_urls:
                    landing_urls.append(u)

            promotion = promotion_type_for_counts(
                len(product_domains), len(seo_domains), len(article_domains), len(unknown_domains)
            )

            x["promotion_type"] = promotion
            x["product_count"] = len(product_urls)
            x["product_domains_count"] = len(product_domains)
            x["seo_page_count"] = len(seo_urls)
            x["seo_domains_count"] = len(seo_domains)
            x["article_count"] = len(article_urls)
            x["article_domains_count"] = len(article_domains)
            x["other_count"] = len(other_urls)
            x["unknown_domains_count"] = len(unknown_domains)
            x["filter_count"] = len(unique_filter_urls)
            x["filter_domains_count"] = len(filter_domains)
            x["fe_rus_filter"] = fe_filter
            x["competitor_filter_urls"] = " | ".join(unique_filter_urls)
            x["competitor_landing_urls"] = " | ".join(dict.fromkeys(landing_urls))
            x["rejected_filter_urls"] = " | ".join(dict.fromkeys(rejected_filter_urls))
            for i in range(1, 6):
                x[f"competitor_{i}"] = unique_filter_urls[i - 1] if len(unique_filter_urls) >= i else ""

            if fe_filter:
                x["status"] = "EXISTS"
                x["reason"] = "FE-RUS уже имеет filter URL для этого значения в TOP-30"
            elif promotion == "ТОВАР":
                x["status"] = "SKIP"
                x["reason"] = f"Товарный интент: PRODUCT-доменов={len(product_domains)}, SEO-доменов={len(seo_domains)}"
            elif promotion == "СМЕШАННЫЙ":
                x["status"] = "REVIEW"
                x["reason"] = f"Смешанный интент: PRODUCT-доменов={len(product_domains)}, SEO-доменов={len(seo_domains)}"
            elif promotion == "ИНФОРМАЦИОННАЯ":
                x["status"] = "SKIP"
                x["reason"] = f"Информационный интент: ARTICLE-доменов={len(article_domains)}"
            elif promotion == "НЕДОСТАТОЧНО ДАННЫХ":
                x["status"] = "REVIEW"
                x["reason"] = "Не удалось надежно определить товарный/каталожный интент TOP-30"
            elif len(filter_domains) >= 2:
                x["status"] = "CREATE"
                x["reason"] = f"SEO-интент + {len(filter_domains)} независимых конкурентов с подтвержденным filter URL для значения «{value}»"
            elif len(filter_domains) == 1:
                x["status"] = "REVIEW"
                x["reason"] = f"SEO-интент, но только 1 независимый конкурент с подтвержденным filter URL для значения «{value}»"
            elif len(seo_domains) >= 2:
                x["status"] = "REVIEW"
                x["reason"] = f"SEO-интент подтвержден {len(seo_domains)} доменами, но точных filter URL недостаточно для CREATE"
            elif len(landing_domains) >= 1:
                x["status"] = "REVIEW"
                x["reason"] = f"SEO-интент, найдено посадочных доменов={len(landing_domains)}, но недостаточно точных filter-конкурентов"
            else:
                x["status"] = "SKIP"
                x["reason"] = "SEO-интент не подтвержден независимыми каталожными/filter-конкурентами"

        if progress:
            progress("Классификация TOP-30 завершена.")
        return res

    def check_urls(self, res):
        s = http_session()
        rows = [x for x in res.get("rows", []) if x.get("status") in ("CREATE", "REVIEW")]
        total = len(rows)
        checked = 0
        for idx, x in enumerate(rows, 1):
            u = x.get("target_url") or ""
            if not u:
                x["url_status"] = "NO_URL"
                checked += 1
                continue
            try:
                # КРИТИЧНО: сначала смотрим ИСХОДНЫЙ HTTP-код.
                # allow_redirects=False нужен потому, что 301 сам по себе
                # означает, что URL уже занят и страницу создавать нельзя.
                initial, location, final_status, final_url, redirected = check_public_url(s, u)
                x["http_status"] = initial
                x["initial_http_status"] = initial
                x["location"] = location
                x["final_http_status"] = final_status
                x["final_url"] = final_url
                x["redirect"] = redirected

                if initial == 200:
                    x["url_status"] = "ALREADY_EXISTS"
                    if x.get("status") == "CREATE":
                        x["status"] = "EXISTS"
                elif initial == 301:
                    # По правилам проекта 301 = страница уже создана/URL занят.
                    x["url_status"] = "EXISTS_301"
                    x["status"] = "EXISTS"
                elif initial == 404:
                    x["url_status"] = "READY_FOR_OCFILTER"
                else:
                    x["url_status"] = f"HTTP_{initial}"
            except Exception as e:
                x["url_status"] = "ERROR"
                x["url_error"] = str(e)
            checked += 1
        return res

    def ocfilter(self, res, creds, create=False, progress=None):
        basic_user, basic_pass, oc_user, oc_pass = creds
        s = http_session()
        if progress:
            progress("Подключение к OpenCart / OCFilter...")
        token = oc_login(s, basic_user, basic_pass, oc_user, oc_pass)
        if progress:
            progress("OpenCart авторизация: OK. USER TOKEN получен.")
        langs, statusview = oc_form_info(s, token)
        if progress:
            progress(f"Форма OCFilter: LANGUAGE IDS={','.join(langs)} | STATUSVIEW={'YES' if statusview else 'NO'}")
        existing = oc_existing(s, token)
        if progress:
            progress(f"Найдено элементов в списке OCFilter: {len(existing)}")
        items = [x for x in res["rows"] if x.get("status") == "CREATE"]
        total = len(items)
        if progress:
            progress(f"CREATE к проверке OCFilter: {total}")
        results = []
        for idx, x in enumerate(items, 1):
            category = x["category"]
            base_alias = x.get("alias") or slugify(f"{x['filter']}-{x['value']}")
            if progress:
                progress(f"[{idx}/{total}] {x.get('keyword','')} | {x.get('filter','')}={x.get('value','')} | проверка категории...")
            alias = choose_alias(existing, base_alias, category)
            # Для вариантов одной родительской страницы сначала используем
            # сохраненный category_id. Если анализ категории был выполнен до
            # сохранения авторизаций или точный ID тогда не определился,
            # ОБЯЗАТЕЛЬНО повторно разрешаем родительскую категорию здесь,
            # уже имея рабочую авторизацию OpenCart.
            parent_cid = str(res.get("category_id") or "").strip()
            parent_slug = str(res.get("category_slug") or "").strip()
            if parent_cid and parent_cid.isdigit():
                resolved = {
                    "category_id": parent_cid,
                    "name": category,
                    "score": 9999,
                    "term": "PARENT_CATEGORY_ID",
                }
                if progress:
                    progress(f"[{idx}/{total}] CATEGORY ID={parent_cid} | из родительской категории")
            else:
                if progress:
                    progress(
                        f"[{idx}/{total}] CATEGORY ID у анализа отсутствует | "
                        f"повторно ищем родительскую категорию: {category}"
                    )
                log_fn = (lambda msg: progress(f"[{idx}/{total}] {msg}") if progress else None)

                # Если анализ не сохранил ID, сначала ищем его напрямую по
                # точному публичному SEO slug родительской категории.
                # Это надежнее autocomplete и не допускает подмены 476 на 474.
                resolved = None
                if parent_slug:
                    direct_cid = oc_category_id_by_seo_keyword(s, token, parent_slug, log=log_fn)
                    if direct_cid:
                        resolved = {
                            "category_id": str(direct_cid),
                            "name": category,
                            "score": 10000,
                            "term": "DIRECT_SEO_SLUG",
                        }
                        if progress:
                            progress(f"[{idx}/{total}] CATEGORY ID={direct_cid} | найден по точному SEO slug")

                # Если прямой поиск не дал ID, используем полный resolver.
                if not resolved:
                    resolved = oc_resolve_category(
                        s, token, category,
                        category_slug=parent_slug,
                        log=log_fn,
                    )
                if not resolved:
                    resolved = oc_resolve_category(
                        s, token, category,
                        category_slug="",
                        log=log_fn,
                    )

            if not resolved:
                x["oc_status"] = "CATEGORY_NOT_FOUND"
                x["oc_message"] = (
                    "Категория не найдена через OpenCart autocomplete; "
                    f"category={category}; category_slug={parent_slug or '-'}; "
                    f"category_id_from_analysis={parent_cid or '-'}"
                )
                results.append(x)
                if progress:
                    progress(
                        f"[{idx}/{total}] CATEGORY_NOT_FOUND | "
                        f"category={category} | slug={parent_slug or '-'} | "
                        f"analysis_category_id={parent_cid or '-'}"
                    )
                continue
            # Сохраняем реально найденный category_id в общей структуре.
            # После первого успешного разрешения остальные варианты используют
            # тот же проверенный ID, а экспорт получает его без повторного поиска.
            res["category_id"] = str(resolved["category_id"])
            x["category_id"] = str(resolved["category_id"])
            x["oc_category_id"] = resolved["category_id"]
            x["oc_category_found"] = resolved["name"]
            x["oc_category_score"] = resolved["score"]
            if not oc_category_safe(category, resolved["name"]):
                x["oc_status"] = "CATEGORY_REVIEW"
                x["oc_message"] = f"Найдено: {resolved['name']} | ожидается: {category}"
                results.append(x)
                if progress: progress(f"[{idx}/{total}] CATEGORY_REVIEW | {resolved['name']}")
                continue
            x["alias"] = alias
            x["target_url"] = expected_url(res["url"], alias)
            if not create:
                x["oc_status"] = "READY"
                results.append(x)
                if progress: progress(f"[{idx}/{total}] READY | category_id={resolved['category_id']} | alias={alias} | URL={x['target_url']}")
                continue
            ok, msg = post_ocfilter_page(s, token, x, resolved["category_id"], langs, statusview, alias)
            if ok:
                x["oc_status"] = "CREATED_FALLBACK_ALIAS" if alias != base_alias else "CREATED"
                existing.add(alias.lower())
            elif msg == "SEO_ALIAS_EXISTS":
                retry = fallback_alias(alias, category)
                ok2, msg2 = post_ocfilter_page(s, token, x, resolved["category_id"], langs, statusview, retry)
                if ok2:
                    x["alias"] = retry; x["target_url"] = expected_url(res["url"], retry); x["oc_status"] = "CREATED_FALLBACK_ALIAS"; existing.add(retry.lower())
                else:
                    x["oc_status"] = "ERROR"; x["oc_message"] = msg2
            else:
                x["oc_status"] = "ERROR"; x["oc_message"] = msg
            results.append(x)
            if progress: progress(f"[{idx}/{total}] {x.get('oc_status')} | {x.get('target_url','')}")
        return results


# -----------------------------------------------------------------------------
# EXPORT
# -----------------------------------------------------------------------------

def generator_rows(res):
    """Финальный набор строк для внешнего генератора контента."""
    out = []
    for x in res.get("rows", []):
        # Страница должна пройти оба независимых этапа:
        # URL = 404 (страницы нет) и OCFilter = READY (категория/alias проверены).
        if x.get("status") != "CREATE":
            continue
        # Внешний генератор получает только SEO-посадочные, не товарный интент.
        if x.get("promotion_type") != "SEO-СТРАНИЦА":
            continue
        if x.get("url_status") not in ("READY_FOR_OCFILTER", "HTTP_404"):
            continue
        if x.get("oc_status") != "READY":
            continue
        row = {
            "keyword": x.get("keyword", ""),
            "promotion_type": x.get("promotion_type", ""),
            "category": x.get("category", ""),
            "category_id": x.get("oc_category_id") or x.get("category_id", ""),
            "filter": x.get("filter", ""),
            "filter_keyword": x.get("filter_keyword", ""),
            "value": x.get("value", ""),
            "value_id": x.get("value_id", ""),
            "params": x.get("params", ""),
            "value_keyword": x.get("value_keyword", ""),
            "alias": x.get("alias", ""),
            "filter_result_url": x.get("filter_result_url") or filter_result_url(
                res.get("url", ""), x.get("filter_keyword", ""), x.get("value_keyword", ""), x.get("value", "")
            ),
            "target_url": x.get("target_url", ""),
            "competitor_1": x.get("competitor_1", ""),
            "competitor_2": x.get("competitor_2", ""),
            "competitor_3": x.get("competitor_3", ""),
            "competitor_4": x.get("competitor_4", ""),
            "competitor_5": x.get("competitor_5", ""),
            "competitor_filter_urls": x.get("competitor_filter_urls", ""),
            "competitor_landing_urls": x.get("competitor_landing_urls", ""),
            "competitor_count": x.get("filter_domains_count", 0),
            "url_status": x.get("url_status", ""),
            "existing_seo_url": x.get("existing_seo_url", ""),
            "existing_seo_alias": x.get("existing_seo_alias", ""),
            "oc_status": x.get("oc_status", ""),
            "generator_status": "READY",
        }
        out.append(row)
    return out


def export_generator_csv(res, path):
    rows = generator_rows(res)
    if not rows:
        raise RuntimeError(
            "Нет готовых строк для генератора. "
            "Сначала выполните Проверку URL и Проверку OCFilter (DRY-RUN)."
        )
    df = pd.DataFrame(rows)
    # Финальный файл для Excel/внешнего генератора: разделитель ';'.
    # UTF-8 BOM нужен Excel для корректного отображения кириллицы.
    df.to_csv(path, sep=";", index=False, encoding="utf-8-sig", lineterminator="\n")
    return len(df)


def export_excel(res, path, topdf=None):
    wb = Workbook()
    ws = wb.active; ws.title = "SUMMARY"
    c = Counter(x.get("status", "") for x in res.get("rows", []))
    ready = len(generator_rows(res))
    promo = Counter(x.get("promotion_type", "") for x in res.get("rows", []))
    summary = [
        ("Категория", res.get("category", "")), ("Category ID", res.get("category_id", "")),
        ("URL", res.get("url", "")), ("Фильтров", len(res.get("options", []))),
        ("Вариантов", len(res.get("rows", []))), ("ГОТОВО ДЛЯ ГЕНЕРАТОРА", ready),
        ("ТИП: SEO-СТРАНИЦА", promo.get("SEO-СТРАНИЦА", 0)),
        ("ТИП: ТОВАР", promo.get("ТОВАР", 0)),
        ("ТИП: СМЕШАННЫЙ", promo.get("СМЕШАННЫЙ", 0)),
        ("ТИП: ИНФОРМАЦИОННАЯ", promo.get("ИНФОРМАЦИОННАЯ", 0)),
    ] + [(k, c[k]) for k in ("CREATE", "SKIP", "REVIEW", "EXISTS")]
    for row in summary: ws.append(row)
    ws["A1"].font = Font(bold=True)
    headers = [
        "keyword","promotion_type","product_count","product_domains_count","seo_page_count","seo_domains_count","article_count","article_domains_count","other_count","unknown_domains_count",
        "category","category_id","filter","filter_keyword","option_id","value","value_id","params","value_keyword","alias","target_url",
        "status","reason","filter_count","filter_domains_count","competitor_1","competitor_2","competitor_3","competitor_4","competitor_5",
        "competitor_filter_urls","competitor_landing_urls","rejected_filter_urls","url_status","existing_seo_url","existing_seo_alias","http_status","final_url","redirect","oc_status","oc_category_id","oc_category_found","oc_category_score","oc_message"
    ]
    ws = wb.create_sheet("CREATE"); ws.append(headers)
    for x in res.get("rows", []):
        if x.get("status") == "CREATE" or x.get("oc_status"):
            ws.append([x.get(h, "") for h in headers])
    ws = wb.create_sheet("READY_FOR_GENERATOR")
    gen = generator_rows(res)
    gen_headers = list(gen[0].keys()) if gen else [
        "keyword","promotion_type","category","category_id","filter","filter_keyword","value","value_id","params","value_keyword","alias","filter_result_url","target_url",
        "competitor_1","competitor_2","competitor_3","competitor_4","competitor_5","competitor_filter_urls","competitor_landing_urls","competitor_count","url_status","oc_status","generator_status"
    ]
    ws.append(gen_headers)
    for x in gen: ws.append([x.get(h, "") for h in gen_headers])
    ws = wb.create_sheet("SKIP"); ws.append(headers)
    for x in res.get("rows", []):
        if x.get("status") == "SKIP": ws.append([x.get(h, "") for h in headers])
    ws = wb.create_sheet("REVIEW"); ws.append(headers)
    for x in res.get("rows", []):
        if x.get("status") == "REVIEW": ws.append([x.get(h, "") for h in headers])
    ws = wb.create_sheet("FILTERS"); ws.append(["filter","filter_keyword","option_id","value","value_id","value_keyword","params","alias"])
    for o in res.get("options", []):
        for v in o.get("values", []):
            ws.append([o.get("name"),o.get("keyword"),o.get("option_id"),v.get("name"),v.get("value_id"),v.get("keyword"),v.get("params"),value_alias(o,v)])
    if topdf is not None and not topdf.empty:
        ws = wb.create_sheet("WORDKEEPER_TOP30")
        for j,cname in enumerate(topdf.columns,1): ws.cell(1,j,cname)
        for i,row in enumerate(topdf.itertuples(index=False),2):
            for j,val in enumerate(row,1): ws.cell(i,j,str(val) if pd.notna(val) else "")
    wb.save(path)


# -----------------------------------------------------------------------------
# GUI
# -----------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FE-RUS SEO Manager v1.21.5")
        self.geometry("1280x860")
        self.q = queue.Queue()
        self.res = None
        self.topdf = pd.DataFrame()
        self.cfg = load_cfg()
        self.wk_stop_event = threading.Event()
        self.wk_running = False
        self.analysis_stop_event = threading.Event()
        self.analysis_running = False
        self.url_check_running = False
        self.build()
        self.load_cfg_to_ui()
        self.setup_clipboard_shortcuts()
        self.after(100, self.poll)

    def build(self):
        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tabs = {}
        names = ["1. Категория","2. WordKeeper","3. Анализ","4. Проверка URL","5. OCFilter","6. Экспорт"]
        for name in names:
            f = ttk.Frame(nb); nb.add(f, text=name); self.tabs[name] = f
        self.build_category(); self.build_wordkeeper(); self.build_analysis(); self.build_urls(); self.build_ocfilter(); self.build_export()

    def build_category(self):
        f=self.tabs["1. Категория"]
        box=ttk.LabelFrame(f,text="Родительская категория"); box.pack(fill="x",padx=10,pady=10)
        ttk.Label(box,text="URL:").grid(row=0,column=0,padx=8,pady=7,sticky="w")
        self.url=ttk.Entry(box); self.url.grid(row=0,column=1,columnspan=4,sticky="ew",padx=8)
        self.url.bind("<Control-c>", self.entry_copy)
        self.url.bind("<Control-C>", self.entry_copy)
        ttk.Label(box,text="Рабочая папка:").grid(row=1,column=0,padx=8,pady=7,sticky="w")
        self.folder=ttk.Entry(box); self.folder.grid(row=1,column=1,columnspan=3,sticky="ew",padx=8)
        ttk.Button(box,text="ОБЗОР",command=self.pickfolder).grid(row=1,column=4,padx=8)
        for c in range(5): box.columnconfigure(c,weight=1)
        ttk.Button(f,text="ПОЛУЧИТЬ КАТЕГОРИЮ И ФИЛЬТРЫ",command=self.analyze_category).pack(anchor="w",padx=10,pady=4)
        self.catstatus=ttk.Label(f,text="Нет анализа"); self.catstatus.pack(anchor="w",padx=12,pady=7)

        # Фильтры/характеристики категории. По умолчанию отмечены все.
        fb=ttk.LabelFrame(f,text="Характеристики для генерации вариантов")
        fb.pack(fill="x",padx=10,pady=4)
        self.filter_vars={}
        self.filter_selection={}
        self.filter_selection_url=""
        self.filter_frame=fb
        ttk.Label(fb,text="После получения категории можно снять ненужные характеристики.").pack(anchor="w",padx=8,pady=(5,2))
        self.filter_checks_frame=ttk.Frame(fb)
        self.filter_checks_frame.pack(fill="x",padx=8,pady=4)

        self.logtext=tk.Text(f,font=("Consolas",10),height=18); self.logtext.pack(fill="both",expand=True,padx=10,pady=8)

    def build_wordkeeper(self):
        f=self.tabs["2. WordKeeper"]
        box=ttk.LabelFrame(f,text="Авторизация WordKeeper"); box.pack(fill="x",padx=10,pady=8)
        self.wk={}
        specs=[("Логин / email",0,0,False),("Пароль",0,2,True)]
        for n,r,c,sec in specs:
            ttk.Label(box,text=n+":").grid(row=r,column=c,padx=8,pady=6,sticky="w")
            e=ttk.Entry(box,show="*" if sec else ""); e.grid(row=r,column=c+1,sticky="ew",padx=8); self.wk[n]=e
        box.columnconfigure(1,weight=1); box.columnconfigure(3,weight=1)
        # Основное управление сбором TOP-30.
        row=ttk.Frame(f); row.pack(fill="x",padx=10,pady=(6,2))
        ttk.Label(row,text="Регион:").pack(side="left")
        self.region=tk.IntVar(value=int(self.cfg.get("region",213) or 213))
        ttk.Spinbox(row,from_=0,to=9999,textvariable=self.region,width=7).pack(side="left",padx=8)
        self.wk_start_btn=ttk.Button(row,text="ПОЛУЧИТЬ TOP-30 WORDKEEPER",command=self.start_wordkeeper)
        self.wk_start_btn.pack(side="left")
        self.wk_stop_btn=ttk.Button(row,text="ОСТАНОВИТЬ АНАЛИЗ",command=self.stop_wordkeeper,state="disabled")
        self.wk_stop_btn.pack(side="left",padx=8)
        ttk.Button(row,text="ЗАГРУЗИТЬ CSV",command=self.load_wk_csv).pack(side="left",padx=8)
        self.wkstatus=ttk.Label(row,text="TOP-30 ещё не загружен"); self.wkstatus.pack(side="left",padx=8)

        # Пересборка вынесена в отдельную строку, чтобы кнопка не исчезала
        # за пределами окна при небольшой ширине интерфейса.
        rebuild_row=ttk.Frame(f); rebuild_row.pack(fill="x",padx=10,pady=(2,6))
        self.wk_rebuild_btn=ttk.Button(
            rebuild_row,
            text="ПЕРЕСОБРАТЬ TOP-30 ТЕКУЩЕГО РАЗДЕЛА",
            command=self.rebuild_wordkeeper_current_section,
        )
        self.wk_rebuild_btn.pack(side="left")
        ttk.Label(
            rebuild_row,
            text="Удаляет только TOP-30 текущего раздела и запускает его заново",
        ).pack(side="left",padx=10)
        self.wklog=tk.Text(f,font=("Consolas",10)); self.wklog.pack(fill="both",expand=True,padx=10,pady=8)

    def build_analysis(self):
        f=self.tabs["3. Анализ"]
        row=ttk.Frame(f); row.pack(fill="x",padx=10,pady=8)
        self.compare_btn=ttk.Button(row,text="СОПОСТАВИТЬ С WORDKEEPER",command=self.compare); self.compare_btn.pack(side="left")
        self.analysis_stop_btn=ttk.Button(row,text="ОСТАНОВИТЬ АНАЛИЗ",command=self.stop_analysis,state="disabled"); self.analysis_stop_btn.pack(side="left",padx=8)
        ttk.Button(row,text="ПОКАЗАТЬ CREATE",command=lambda:self.show_status("CREATE")).pack(side="left",padx=8)
        ttk.Button(row,text="ПОКАЗАТЬ SKIP",command=lambda:self.show_status("SKIP")).pack(side="left")
        ttk.Button(row,text="ПОКАЗАТЬ REVIEW",command=lambda:self.show_status("REVIEW")).pack(side="left",padx=8)
        ttk.Button(row,text="ПОКАЗАТЬ ТОВАРНЫЕ",command=lambda:self.show_promotion_type("ТОВАР")).pack(side="left",padx=8)
        ttk.Button(row,text="ПОКАЗАТЬ SEO",command=lambda:self.show_promotion_type("SEO-СТРАНИЦА")).pack(side="left")
        ttk.Button(row,text="ПОКАЗАТЬ СМЕШАННЫЕ",command=lambda:self.show_promotion_type("СМЕШАННЫЙ")).pack(side="left",padx=8)
        self.stats=tk.Text(f,font=("Consolas",10)); self.stats.pack(fill="both",expand=True,padx=10,pady=8)

    def build_urls(self):
        f=self.tabs["4. Проверка URL"]
        row=ttk.Frame(f); row.pack(fill="x",padx=10,pady=10)
        self.url_check_btn=ttk.Button(row,text="ПРОВЕРИТЬ CREATE URL",command=self.check_urls)
        self.url_check_btn.pack(side="left")
        self.url_stop_btn=ttk.Button(row,text="ОСТАНОВИТЬ ПРОВЕРКУ",command=self.stop_url_check,state="disabled")
        self.url_stop_btn.pack(side="left",padx=8)
        self.urlstatus=ttk.Label(row,text="Проверка не запускалась")
        self.urlstatus.pack(side="left",padx=8)
        self.urllog=tk.Text(f,font=("Consolas",10)); self.urllog.pack(fill="both",expand=True,padx=10,pady=8)

    def build_ocfilter(self):
        f=self.tabs["5. OCFilter"]
        box=ttk.LabelFrame(f,text="Авторизация OpenCart / OCFilter"); box.pack(fill="x",padx=10,pady=8)
        self.oc={}
        specs=[("HTTP Basic login",0,0,False),("HTTP Basic password",0,2,True),("OpenCart login",1,0,False),("OpenCart password",1,2,True)]
        for n,r,c,sec in specs:
            ttk.Label(box,text=n+":").grid(row=r,column=c,padx=8,pady=6,sticky="w")
            e=ttk.Entry(box,show="*" if sec else ""); e.grid(row=r,column=c+1,sticky="ew",padx=8); self.oc[n]=e
        box.columnconfigure(1,weight=1); box.columnconfigure(3,weight=1)
        row=ttk.Frame(f); row.pack(fill="x",padx=10,pady=8)
        ttk.Button(row,text="СОХРАНИТЬ АВТОРИЗАЦИИ",command=self.save_settings).pack(side="left")
        self.oc_check_btn=ttk.Button(row,text="ПРОВЕРИТЬ OCFILTER (DRY-RUN)",command=self.run_oc)
        self.oc_check_btn.pack(side="left",padx=8)
        ttk.Label(f,text="Сохраняются обе авторизации: HTTP Basic + OpenCart. Файл: %APPDATA%\\FE-RUS SEO Manager\\config.json").pack(anchor="w",padx=12,pady=(0,4))
        ttk.Label(f,text="Показывать в ТОП меню: всегда ОТКЛЮЧЕНО (menu_status=0)").pack(anchor="w",padx=12)
        self.oclog=tk.Text(f,font=("Consolas",10)); self.oclog.pack(fill="both",expand=True,padx=10,pady=8)

    def build_export(self):
        f=self.tabs["6. Экспорт"]
        ttk.Label(f,text="Финальный экспорт для внешней программы генерации контента").pack(anchor="w",padx=10,pady=12)
        ttk.Button(f,text="СОХРАНИТЬ CSV ДЛЯ ГЕНЕРАТОРА",command=self.export_generator).pack(anchor="w",padx=10,pady=4)
        ttk.Button(f,text="ЭКСПОРТИРОВАТЬ ПОЛНЫЙ XLSX",command=self.export).pack(anchor="w",padx=10,pady=4)
        ttk.Label(f,text="CSV содержит только готовые SEO-страницы: promotion_type=SEO-СТРАНИЦА + status=CREATE + URL/OCFilter=READY. target_url - будущий SEO URL, filter_result_url - старый URL фильтрации.").pack(anchor="w",padx=10,pady=8)
        self.export_log=tk.Text(f,font=("Consolas",10),height=24)
        self.export_log.pack(fill="both",expand=True,padx=10,pady=8)
        ttk.Button(f,text="ПОКАЗАТЬ ГОТОВЫЕ URL ФИЛЬТРАЦИИ",command=self.show_ready_filter_urls).pack(anchor="w",padx=10,pady=4)
        ttk.Button(f,text="СКОПИРОВАТЬ ВСЕ ССЫЛКИ",command=self.copy_ready_filter_urls).pack(anchor="w",padx=10,pady=4)

    # Clipboard
    def setup_clipboard_shortcuts(self):
        # ttk.Entry имеет отдельный bindtag TEntry. Старый bindtag Entry
        # поэтому не срабатывал в первом окне. Подключаем оба класса.
        for cls in ("Entry", "TEntry"):
            for seq, fn in (
                ("<Control-a>", self.entry_select_all), ("<Control-A>", self.entry_select_all),
                ("<Control-c>", self.entry_copy), ("<Control-C>", self.entry_copy),
                ("<Control-v>", self.entry_paste), ("<Control-V>", self.entry_paste),
                ("<Control-x>", self.entry_cut), ("<Control-X>", self.entry_cut),
                ("<Shift-Insert>", self.entry_paste),
            ):
                self.bind_class(cls, seq, fn)

        # Дополнительно вешаем обработчики непосредственно на все Entry.
        # Это особенно надёжно в собранном PyInstaller EXE.
        entries = []
        if hasattr(self, "url"): entries.append(self.url)
        if hasattr(self, "folder"): entries.append(self.folder)
        entries.extend(getattr(self, "wk", {}).values())
        entries.extend(getattr(self, "oc", {}).values())
        for widget in entries:
            for seq, fn in (
                ("<Control-a>", self.entry_select_all),
                ("<Control-c>", self.entry_copy),
                ("<Control-v>", self.entry_paste),
                ("<Control-x>", self.entry_cut),
                ("<Shift-Insert>", self.entry_paste),
            ):
                widget.bind(seq, fn, add="+")
            # <Control-c>/<Control-v>/... определяются Tk по СИМВОЛУ, который
            # печатает клавиша - при русской раскладке физическая Ctrl+C/V/X/A
            # часто не даёт нужный символ, и привязки выше просто не
            # срабатывают. Добавляем резервный обработчик по ФИЗИЧЕСКОМУ коду
            # клавиши (не зависит от раскладки): 65=A, 67=C, 86=V, 88=X -
            # вызывает те же entry_copy/entry_paste/entry_cut/entry_select_all.
            widget.bind("<Key>", self.entry_key_by_keycode, add="+")
            # Контекстное меню правой кнопкой - не зависит вообще ни от
            # раскладки, ни от кода клавиши: напрямую дёргает те же функции
            # копирования/вставки по клику мышью. Самый надёжный резерв.
            self._add_entry_context_menu(widget)

        self.bind_class("Text", "<Control-a>", self.text_select_all)
        self.bind_class("Text", "<Control-A>", self.text_select_all)
        self.bind_class("Text", "<Control-c>", self.text_copy)
        self.bind_class("Text", "<Control-C>", self.text_copy)

    def text_select_all(self, e):
        e.widget.tag_add("sel", "1.0", "end-1c")
        e.widget.mark_set("insert", "end-1c")
        return "break"

    def text_copy(self, e):
        try:
            value = e.widget.get("sel.first", "sel.last")
        except Exception:
            value = e.widget.get("1.0", "end-1c")
        if not win_clipboard_set(value):
            try:
                self.clipboard_clear(); self.clipboard_append(value); self.update()
            except Exception:
                pass
        return "break"

    def entry_select_all(self, e):
        e.widget.select_range(0, "end")
        e.widget.icursor("end")
        return "break"

    def _add_entry_context_menu(self, widget):
        fake_event = types.SimpleNamespace(widget=widget)
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="Вырезать", command=lambda: self.entry_cut(fake_event))
        menu.add_command(label="Копировать", command=lambda: self.entry_copy(fake_event))
        menu.add_command(label="Вставить", command=lambda: self.entry_paste(fake_event))
        menu.add_separator()
        menu.add_command(label="Выделить всё", command=lambda: self.entry_select_all(fake_event))

        def show_menu(event):
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        widget.bind("<Button-3>", show_menu, add="+")

    def entry_key_by_keycode(self, e):
        """Резервная обработка Ctrl+A/C/V/X по физическому коду клавиши -
        не зависит от раскладки клавиатуры (см. комментарий в
        setup_clipboard_shortcuts). Срабатывает, только если зажат Control
        (e.state & 0x4) и код клавиши совпадает; иначе не мешает обычному
        вводу текста."""
        if not (e.state & 0x4):
            return None
        if e.keycode == 65:
            return self.entry_select_all(e)
        if e.keycode == 67:
            return self.entry_copy(e)
        if e.keycode == 86:
            return self.entry_paste(e)
        if e.keycode == 88:
            return self.entry_cut(e)
        return None

    def entry_copy(self, e):
        try:
            value = e.widget.selection_get()
        except Exception:
            value = e.widget.get()
        if not win_clipboard_set(value):
            try:
                self.clipboard_clear(); self.clipboard_append(value); self.update()
            except Exception:
                pass
        return "break"

    def entry_paste(self, e):
        value = win_clipboard_get()
        if not value:
            try:
                value = self.clipboard_get()
            except Exception:
                value = ""
        if value:
            try:
                e.widget.delete("sel.first", "sel.last")
            except Exception:
                pass
            try:
                e.widget.insert("insert", value.replace("\x00", ""))
            except Exception:
                pass
        return "break"

    def entry_cut(self, e):
        try:
            value = e.widget.selection_get()
        except Exception:
            value = ""
        if value:
            if not win_clipboard_set(value):
                try:
                    self.clipboard_clear(); self.clipboard_append(value); self.update()
                except Exception:
                    pass
            try:
                e.widget.delete("sel.first", "sel.last")
            except Exception:
                pass
        return "break"

    def load_cfg_to_ui(self):
        self.url.insert(0,self.cfg.get("url","https://fe-rus.ru/tavr-alyuminievyj/"))
        self.folder.insert(0,self.cfg.get("folder",r"D:\wordkeeper"))
        self.wk["Логин / email"].insert(0,self.cfg.get("wk_login",""))
        self.wk["Пароль"].insert(0,self.cfg.get("wk_password",""))
        ocfg=self.cfg.get("ocfilter",{}) if isinstance(self.cfg.get("ocfilter",{}),dict) else {}
        for k,e in self.oc.items():
            e.insert(0,self.cfg.get(k,ocfg.get(k,"")))

    def save_settings(self):
        d={
            "url":self.url.get(),
            "folder":self.folder.get(),
            "region":int(self.region.get()),
            "wk_login":self.wk["Логин / email"].get(),
            "wk_password":self.wk["Пароль"].get(),
            **{k:e.get() for k,e in self.oc.items()}
        }
        d["ocfilter"]={k:e.get() for k,e in self.oc.items()}
        save_cfg(d); self.cfg=d

    def pickfolder(self):
        p=filedialog.askdirectory(initialdir=self.folder.get() or os.getcwd())
        if p:self.folder.delete(0,"end");self.folder.insert(0,p);self.save_settings()

    def remember_filter_selection(self):
        """Запоминает текущие состояния чекбоксов перед повторным получением категории."""
        if not getattr(self, "filter_vars", None):
            return
        self.filter_selection = {
            str(k): bool(v.get()) for k, v in self.filter_vars.items()
        }
        self.filter_selection_url = str((getattr(self, "res", {}) or {}).get("url", ""))

    def render_filter_checkboxes(self, options):
        # Не сбрасываем выбор пользователя при повторном получении той же категории.
        previous = dict(getattr(self, "filter_selection", {}) or {})
        # Выбор переносим только внутри одной и той же категории.
        # При переходе на новую категорию все её характеристики включаются.
        same_category = (
            str(getattr(self, "filter_selection_url", ""))
            == str((getattr(self, "url", None).get() if getattr(self, "url", None) else ""))
        )
        if not same_category:
            previous = {}

        for w in self.filter_checks_frame.winfo_children():
            w.destroy()
        self.filter_vars = {}
        current_selection = {}

        for i, o in enumerate(options):
            name = str(o.get("name") or o.get("keyword") or "").strip()
            if not name:
                continue
            key = str(o.get("option_id") or name)

            # Существующая характеристика сохраняет выбор. Новая включена по умолчанию.
            value = previous.get(key, True)
            var = tk.BooleanVar(value=bool(value))
            self.filter_vars[key] = var
            current_selection[key] = bool(value)

            cb = ttk.Checkbutton(
                self.filter_checks_frame,
                text=name,
                variable=var,
                command=self.remember_filter_selection,
            )
            cb.grid(row=i // 4, column=i % 4, sticky="w", padx=8, pady=3)

        self.filter_selection = current_selection
        self.filter_selection_url = str((getattr(self, "url", None).get() if getattr(self, "url", None) else ""))
        for c in range(4):
            self.filter_checks_frame.columnconfigure(c, weight=1)

    def selected_option_ids(self):
        return {k for k,v in self.filter_vars.items() if v.get()}

    def apply_selected_filters(self):
        if not self.res or not self.filter_vars:
            return
        selected=self.selected_option_ids()
        if not selected:
            messagebox.showwarning("Характеристики","Выберите хотя бы одну характеристику.")
            return
        opts=[o for o in self.res.get("options",[]) if str(o.get("option_id") or o.get("name")) in selected]
        base_url=self.res.get("url","")
        rows=[]
        for o in opts:
            for v in o.get("values",[]):
                if not str(v.get("name","")).strip():
                    continue
                alias=value_alias(o,v)
                rows.append({
                    "keyword":f"{self.res['category']} {v['name']}",
                    "category":self.res["category"],"category_id":self.res.get("category_id",""),
                    "filter":o.get("name",""),"filter_keyword":o.get("keyword",""),
                    "option_id":o.get("option_id",""),"value":v.get("name",""),
                    "value_id":v.get("value_id",""),"params":v.get("params",""),
                    "value_keyword":v.get("keyword",""),"alias":alias,
                    "filter_result_url":filter_result_url(base_url, o.get("keyword",""), v.get("keyword",""), v.get("name","")),
                    "target_url":expected_url(base_url,alias),"status":"UNVERIFIED",
                    "reason":"требует TOP-30 + проверки OCFilter",
                    "competitor_1":"","competitor_2":"","competitor_3":"","competitor_4":"","competitor_5":"",
                })
        self.res["options"]=opts
        self.res["rows"]=rows
        self.catstatus.config(text=f"Выбрано характеристик: {len(opts)} | вариантов: {len(rows)}")
        self.logtext.insert("end",f"\nПрименены характеристики: {', '.join(o.get('name','') for o in opts)}\nВариантов: {len(rows)}\n")

    def analyze_category(self):
        # Важно: пользователь мог снять характеристики и сразу нажать кнопку.
        # Сохраняем их выбор до того, как новый ответ API заменит self.res.
        self.remember_filter_selection()
        self.save_settings(); self.logtext.delete("1.0","end")
        threading.Thread(target=self.worker_category,daemon=True).start()
    def worker_category(self):
        try:self.q.put(("cat",Pipeline(self.log).category(self.url.get())))
        except Exception as e:self.q.put(("err","Категория: "+str(e)))
    def log(self,x):self.q.put(("log",x))

    def load_wk_csv(self):
        p=filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")],initialdir=self.folder.get() or os.getcwd())
        if not p:return
        try:
            self.topdf=pd.read_csv(p,encoding="utf-8-sig")
            self.wkstatus.config(text=f"Загружено: {Path(p).name} | строк: {len(self.topdf)}")
            self.save_settings()
        except Exception as e:messagebox.showerror("WordKeeper",str(e))

    def start_wordkeeper(self):
        if not self.res:
            messagebox.showwarning("Категория","Сначала получите категорию и фильтры.")
            return
        try:
            self.apply_selected_filters()
        except Exception:
            return
        if self.wk_running:
            return
        self.save_settings()
        self.wk_stop_event.clear()
        self.wk_running=True
        self.wk_start_btn.config(state="disabled")
        self.wk_stop_btn.config(state="normal")
        self.wklog.delete("1.0","end")
        threading.Thread(target=self.worker_wordkeeper,daemon=True).start()

    def stop_wordkeeper(self):
        if not self.wk_running:
            return
        self.wk_stop_event.set()
        self.wkstatus.config(text="Остановка WordKeeper... текущий запрос завершится, затем анализ остановится")
        self.wklog.insert("end","\n!!! Запрошена остановка. Последний завершённый запрос уже сохранён.\n")
        self.wklog.see("end")

    def rebuild_wordkeeper_current_section(self):
        """Удаляет накопленные TOP-30 только для текущего раздела и запускает его заново."""
        if not self.res or not self.res.get("rows"):
            messagebox.showwarning("TOP-30", "Сначала получите категорию и фильтры текущего раздела.")
            return
        if self.wk_running:
            messagebox.showwarning("TOP-30", "Сначала дождитесь завершения или остановите текущий сбор TOP-30.")
            return

        category = str(self.res.get("category") or self.res.get("search_category") or "").strip()
        current_keywords = {
            str(x.get("keyword", "")).strip()
            for x in self.res.get("rows", [])
            if str(x.get("keyword", "")).strip()
        }
        current_variants = set()
        for row in self.res.get("rows", []):
            canonical = str(row.get("keyword", "")).strip()
            if not canonical:
                continue
            variants = wordkeeper_query_variants(
                self.res.get("search_category") or self.res.get("category") or "",
                row.get("value", ""),
            ) or [canonical]
            current_variants.update(norm(q) for q in variants if str(q).strip())

        if not current_keywords:
            messagebox.showwarning("TOP-30", "У текущего раздела нет запросов для пересборки.")
            return

        ok = messagebox.askyesno(
            "Пересобрать TOP-30",
            f"Пересобрать TOP-30 заново только для текущего раздела?\n\n"
            f"Раздел: {category or 'текущий'}\n"
            f"Значений/запросов: {len(current_keywords)}\n\n"
            "Старые результаты этого раздела будут удалены из накопительного CSV и checkpoint. "
            "Результаты других разделов сохранятся.",
        )
        if not ok:
            return

        folder = Path(self.folder.get() or os.getcwd())
        out_path = folder / "seo_top30_result.csv"
        progress_path = folder / "seo_top30_progress.json"
        removed_rows = 0
        removed_queries = 0

        # Удаляем только строки текущего раздела из накопительного CSV.
        if out_path.exists() and out_path.stat().st_size:
            try:
                old = pd.read_csv(out_path, encoding="utf-8-sig")
                if not old.empty:
                    mask = pd.Series(False, index=old.index)
                    if "source_query" in old.columns:
                        mask = mask | old["source_query"].astype(str).map(norm).isin(current_variants)
                    if "keyword" in old.columns:
                        mask = mask | old["keyword"].astype(str).str.strip().isin(current_keywords)
                    removed_rows = int(mask.sum())
                    old = old.loc[~mask].copy()
                    old.to_csv(out_path, index=False, encoding="utf-8-sig")
            except Exception as e:
                messagebox.showerror("TOP-30", f"Не удалось очистить текущий раздел из CSV:\n{e}")
                return

        # Из checkpoint удаляем только текущие canonical keywords и их варианты.
        if progress_path.exists() and progress_path.stat().st_size:
            try:
                progress = json.loads(progress_path.read_text(encoding="utf-8"))
                old_done_keywords = [str(x).strip() for x in progress.get("done_keywords", [])]
                old_done_queries = [str(x).strip() for x in progress.get("done_queries", [])]
                new_done_keywords = [x for x in old_done_keywords if x not in current_keywords]
                new_done_queries = [x for x in old_done_queries if norm(x) not in current_variants]
                removed_queries = len(old_done_queries) - len(new_done_queries)
                progress["done_keywords"] = new_done_keywords
                progress["done_queries"] = new_done_queries
                progress["updated_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
                progress_path.write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")
            except Exception as e:
                messagebox.showerror("TOP-30", f"Не удалось обновить checkpoint:\n{e}")
                return

        self.topdf = pd.DataFrame()
        self.wkstatus.config(text="Старый TOP-30 текущего раздела сброшен")
        self.wklog.delete("1.0", "end")
        self.wklog.insert("end", f"Пересборка текущего раздела: {category or 'текущий раздел'}\n")
        self.wklog.insert("end", f"Удалено старых строк TOP-30: {removed_rows}\n")
        self.wklog.insert("end", f"Сброшено checkpoint-запросов: {removed_queries}\n")
        self.wklog.insert("end", "Запускаем новый сбор TOP-30...\n")
        self.wklog.see("end")
        self.save_settings()
        self.start_wordkeeper()

    def worker_wordkeeper(self):
        out_path=Path(self.folder.get() or os.getcwd())/"seo_top30_result.csv"
        progress_path=Path(self.folder.get() or os.getcwd())/"seo_top30_progress.json"
        try:
            login=self.wk["Логин / email"].get().strip()
            password=self.wk["Пароль"].get()
            region=int(self.region.get())
            if not login or not password:
                raise RuntimeError("Заполните логин и пароль WordKeeper.")
            s=http_session()
            self.q.put(("wklog","WordKeeper: открываем авторизацию..."))
            r=s.get(LOGIN_WK,timeout=60); r.raise_for_status(); csrf=get_csrf(r.text)
            if not csrf:
                raise RuntimeError("CSRF-токен WordKeeper не найден.")
            r=s.post(LOGIN_WK,data={"login":login,"password":password,"auth":"Войти","csrf_token":csrf},timeout=60,allow_redirects=True)
            if "/dashboard" not in r.url:
                raise RuntimeError("Авторизация WordKeeper не удалась.")

            # Отдельная постоянная сессия для AJAX: не трогаем cookie-jar
            # авторизационной сессии, но переиспользуем HTTP-соединение между
            # десятками/сотнями TOP-30 запросов.
            ajax_session = requests.Session()
            ajax_session.headers.update({"User-Agent": UA, "Accept-Language": "ru-RU,ru;q=0.9"})
            cookie_header = "; ".join(
                f"{str(c.name)}={str(c.value)}"
                for c in s.cookies
                if getattr(c, "name", None) is not None
            )

            # Запросы именно ТЕКУЩЕЙ категории.
            # CSV/checkpoint являются накопительным кэшем, но для текущей категории
            # считаем только её канонические keywords. Один keyword может иметь
            # несколько вариантов запроса WordKeeper.
            keywords=list(dict.fromkeys(
                str(x.get("keyword", "")).strip()
                for x in self.res["rows"]
                if x.get("keyword") and str(x.get("keyword", "")).strip()
            ))
            # Формируем карту АКТУАЛЬНЫХ запросов текущей категории до загрузки кэша.
            # Старые seo_top30_result.csv / progress.json могут содержать результаты
            # других категорий. Их нельзя считать выполненными только по факту наличия
            # source_query: учитываем только запросы, которые реально относятся к
            # текущей категории.
            current_variants_by_keyword = {}
            variant_to_canonical = {}
            for row in self.res["rows"]:
                canonical = str(row.get("keyword", "")).strip()
                if not canonical:
                    continue
                variants = wordkeeper_query_variants(
                    self.res.get("search_category") or self.res.get("category") or "",
                    row.get("value", ""),
                ) or [canonical]
                current_variants_by_keyword[canonical] = variants
                for q in variants:
                    variant_to_canonical[norm(q)] = canonical

            keywords = list(current_variants_by_keyword.keys())
            current_keyword_set = set(keywords)
            results=[]
            done=set()
            done_queries=set()
            cached_result_query_keys=set()

            if out_path.exists():
                old=pd.read_csv(out_path,encoding="utf-8-sig")
                if not old.empty:
                    results=old.to_dict("records")
                    # ВАЖНО: если старый CSV был создан предыдущей версией, его
                    # canonical keyword мог быть транслитерированным. Перепривязываем
                    # строки по source_query к текущему русскому canonical keyword.
                    if "source_query" in old.columns:
                        for rr in results:
                            sq = str(rr.get("source_query", "")).strip()
                            canonical = variant_to_canonical.get(norm(sq))
                            if canonical:
                                rr["keyword"] = canonical
                                done_queries.add(sq)
                                cached_result_query_keys.add(norm(sq))
                    # Старые canonical keywords учитываем только если они относятся
                    # к текущей категории.
                    for rr in results:
                        k = str(rr.get("keyword", "")).strip()
                        if k in current_keyword_set:
                            done.add(k)

            if progress_path.exists():
                try:
                    progress=json.loads(progress_path.read_text(encoding="utf-8"))
                    # done_keywords из старого checkpoint не переносим целиком:
                    # оставляем только keywords текущей категории.
                    done.update(
                        str(x).strip() for x in progress.get("done_keywords",[])
                        if str(x).strip() in current_keyword_set
                    )
                    # Аналогично done_queries: только актуальные варианты текущей
                    # категории могут блокировать новый запрос.
                    for x in progress.get("done_queries", []):
                        q = str(x).strip()
                        if q and norm(q) in variant_to_canonical:
                            done_queries.add(q)
                except Exception:
                    pass

            # Нормализованный набор уже выполненных запросов. Это устраняет ложные
            # повторы из-за регистра/лишних пробелов.
            done_query_keys = {norm(q) for q in done_queries if str(q).strip()}

            # Если checkpoint говорит «запрос выполнен», но в накопительном CSV
            # вообще нет ни одной строки с этим source_query, не доверяем старому
            # checkpoint: такой запрос повторяем, иначе 0-результат/старый checkpoint
            # может навсегда скрыть реальные TOP-30.
            done_query_keys = {
                qk for qk in done_query_keys
                if qk in cached_result_query_keys
            } | {
                qk for qk in done_query_keys
                if qk in cached_result_query_keys
            }

            query_plan=[]
            for canonical, variants in current_variants_by_keyword.items():
                pending=[q for q in variants if norm(q) not in done_query_keys]
                if pending:
                    query_plan.append((canonical, pending))

            current_done={
                canonical for canonical, variants in current_variants_by_keyword.items()
                if not any(norm(q) not in done_query_keys for q in variants)
            }
            pending_queries=sum(len(v) for _,v in query_plan)
            self.q.put(("wklog",f"TOP-30 CSV: {out_path}"))
            self.q.put(("wklog",f"Checkpoint: {progress_path}"))
            self.q.put(("wklog",f"Всего значений текущей категории: {len(keywords)} | уже полностью обработано: {len(current_done)} | осталось значений: {len(keywords)-len(current_done)}"))
            self.q.put(("wklog",f"Запросов WordKeeper к выполнению: {pending_queries} | кэшем подтверждено запросов текущей категории: {len(done_query_keys)}"))
            self.q.put(("wklog",f"Строк TOP-30 в накопительном CSV: {len(results)} | source_query с результатами текущей категории: {len(cached_result_query_keys)}"))

            # Один раз строим индекс уже сохранённых пар (source_query, URL).
            # Раньше этот set пересобирался из всего CSV на каждом запросе,
            # что давало лишнюю O(N^2) работу на больших накопительных файлах.
            existing_keys = {
                (str(rr.get("source_query", "")).strip(), str(rr.get("url", "")).strip())
                for rr in results
                if str(rr.get("source_query", "")).strip() and str(rr.get("url", "")).strip()
            }

            query_index=0
            total_pending=pending_queries
            for canonical, variants in query_plan:
                if self.wk_stop_event.is_set():
                    break
                all_variants_done=True
                for actual_query in variants:
                    if self.wk_stop_event.is_set():
                        all_variants_done=False
                        break
                    query_index += 1
                    try:
                        rr=wordkeeper_ajax_post(ajax_session, actual_query, region, cookie_header)
                        rr_status=http_status_code(rr)
                        rr_text=http_response_text(rr)
                        if rr_status in (301,302,303,307,308):
                            location = ""
                            try:
                                location = str(rr.headers.get("Location") or "")
                            except Exception:
                                pass
                            if location and "/login" in location:
                                raise RuntimeError("WordKeeper вернул редирект на /login: сессия авторизации потеряна")
                        if not rr_text:
                            raise RuntimeError(
                                f"WordKeeper вернул пустой ответ (HTTP={rr_status}, тип={type(rr).__name__})"
                            )
                        rows=parse_top30(rr_text,canonical)
                        for rrow in rows:
                            rrow["source_query"]=actual_query
                        done_queries.add(actual_query)
                        done_query_keys.add(norm(actual_query))
                        # Не дублируем одинаковый URL одного и того же source_query.
                        for rr in rows:
                            key = (str(rr.get("source_query", "")).strip(), str(rr.get("url", "")).strip())
                            if key not in existing_keys:
                                results.append(rr)
                                existing_keys.add(key)
                        # Не переписываем весь накопительный CSV после каждого
                        # запроса. Сохраняем каждые WK_SAVE_EVERY запросов, а также
                        # немедленно при остановке/последнем запросе.
                        force_save = (
                            query_index % WK_SAVE_EVERY == 0
                            or self.wk_stop_event.is_set()
                            or query_index == total_pending
                        )
                        if force_save:
                            pd.DataFrame(results).to_csv(out_path,index=False,encoding="utf-8-sig")
                            progress_path.write_text(json.dumps({
                                "done_keywords":safe_sorted_strings(done),
                                "done_queries":safe_sorted_strings(done_queries),
                                "updated_at":time.strftime("%Y-%m-%d %H:%M:%S")
                            },ensure_ascii=False,indent=2),encoding="utf-8")
                            save_note = " | СОХРАНЕНО"
                        else:
                            save_note = ""
                        self.q.put(("wklog",f"[{query_index}/{total_pending}] {canonical} | запрос: {actual_query} | HTTP {rr_status} | результатов: {len(rows)}{save_note}"))
                        if self.wk_stop_event.is_set():
                            all_variants_done=False
                            break
                        time.sleep(WK_DELAY)
                    except Exception as e:
                        all_variants_done=False
                        pd.DataFrame(results).to_csv(out_path,index=False,encoding="utf-8-sig")
                        progress_path.write_text(json.dumps({
                            "done_keywords":safe_sorted_strings(done),
                            "done_queries":safe_sorted_strings(done_queries),
                            "updated_at":time.strftime("%Y-%m-%d %H:%M:%S")
                        },ensure_ascii=False,indent=2),encoding="utf-8")
                        error_log = Path(self.folder.get() or os.getcwd()) / "wordkeeper_error.log"
                        try:
                            error_log.write_text(
                                traceback.format_exc(),
                                encoding="utf-8"
                            )
                        except Exception:
                            pass
                        self.q.put(("wklog",f"ОШИБКА {actual_query}: {e} | текущая категория: обработано {len(set(keywords) & done)} из {len(keywords)} | traceback: {error_log}"))
                        break

                if all_variants_done:
                    done.add(canonical)
                    # CSV/checkpoint сохраняются пакетно выше. Здесь только фиксируем
                    # завершение canonical keyword в памяти, чтобы не переписывать
                    # весь файл после каждого значения. Финальное сохранение выполняется
                    # после цикла, а при ошибке/остановке — принудительно.
                    self.q.put(("wklog",f"ЗАВЕРШЕНО ЗНАЧЕНИЕ: {canonical} | вариантов запроса: {len(variants)}"))

            pd.DataFrame(results).to_csv(out_path,index=False,encoding="utf-8-sig")
            progress_path.write_text(json.dumps({
                "done_keywords":safe_sorted_strings(done),
                "done_queries":safe_sorted_strings(done_queries),
                "updated_at":time.strftime("%Y-%m-%d %H:%M:%S")
            },ensure_ascii=False,indent=2),encoding="utf-8")
            if self.wk_stop_event.is_set():
                self.q.put(("wkstop",{"path":out_path,"progress":progress_path,"done":len(set(keywords) & done),"total":len(keywords)}))
            else:
                self.q.put(("wkdone",out_path))
        except Exception as e:
            self.q.put(("wkerr",str(e)))
        finally:
            self.wk_running=False

    def compare(self):
        if not self.res:
            messagebox.showwarning("Сначала категория", "Сначала выполните анализ категории.")
            return
        if self.analysis_running:
            return
        try:
            # Выбор характеристик фиксируем до запуска фонового анализа.
            self.apply_selected_filters()
            if not self.res.get("rows"):
                return
        except Exception:
            return
        if self.topdf is None or self.topdf.empty:
            messagebox.showwarning("WordKeeper", "Сначала загрузите или получите TOP-30 WordKeeper.")
            return
        self.analysis_stop_event.clear()
        self.analysis_running=True
        self.compare_btn.config(state="disabled")
        self.analysis_stop_btn.config(state="normal")
        self.stats.delete("1.0","end")
        self.stats.insert("end", "Запуск анализа TOP-30...\n\n")
        threading.Thread(target=self.worker_compare,daemon=True).start()

    def stop_analysis(self):
        if self.analysis_running:
            self.analysis_stop_event.set()
            self.stats.insert("end", "\nЗапрошена остановка анализа. Уже классифицированные URL сохранены.\n")
            self.stats.see("end")

    def worker_compare(self):
        try:
            folder=Path(self.folder.get() or os.getcwd())
            cache_path=folder / "seo_serp_page_classification.csv"
            res=Pipeline().merge_analysis(
                self.res,
                self.topdf,
                progress=lambda msg:self.q.put(("analysislog",msg)),
                stop_event=self.analysis_stop_event,
                cache_path=cache_path,
            )
            self.q.put(("analysisdone", {"res":res, "cache":str(cache_path), "stopped":self.analysis_stop_event.is_set()}))
        except Exception as e:
            self.q.put(("analysiserr",str(e)))
        finally:
            self.analysis_running=False

    def _render_analysis(self):
        if not self.res:
            return
        c=Counter(x.get("status") for x in self.res["rows"])
        p=Counter(x.get("promotion_type") for x in self.res["rows"])
        self.stats.delete("1.0","end")
        self.stats.insert("end",f"Категория: {self.res['category']}\nCategory ID: {self.res['category_id']}\nВариантов: {len(self.res['rows'])}\n\n")
        self.stats.insert("end", "ТИП ПРОДВИЖЕНИЯ:\n")
        for k in ("SEO-СТРАНИЦА","ТОВАР","СМЕШАННЫЙ","ИНФОРМАЦИОННАЯ","НЕДОСТАТОЧНО ДАННЫХ"):
            self.stats.insert("end",f"{k}: {p[k]}\n")
        self.stats.insert("end", "\nСТАТУС SEO-СТРАНИЦ:\n")
        for k in ("CREATE","SKIP","REVIEW","EXISTS"):
            self.stats.insert("end",f"{k}: {c[k]}\n")
        self.stats.insert("end", "\nCREATE = только SEO-интент + минимум 2 независимых точных filter-конкурента.\n")
        self.stats.insert("end", "ТОВАР = запрос преимущественно ведет на карточки товаров; отдельную SEO-посадочную не создаем.\n\n")
        self.stats.insert("end", "СПИСОК:\n")
        for x in self.res["rows"]:
            self.stats.insert(
                "end",
                f"- {x['keyword']} | {x['filter']} | {x['value']} | ТИП={x.get('promotion_type','')} | STATUS={x.get('status','')} | PRODUCT={x.get('product_domains_count',0)} | SEO={x.get('seo_domains_count',0)} | конкуренты={x.get('filter_domains_count',0)} | {x.get('reason','')}\n"
            )
        self.stats.see("1.0")

    def show_status(self,status):
        if not self.res:return
        self.stats.delete("1.0","end")
        for x in self.res["rows"]:
            if x.get("status")==status:
                self.stats.insert("end",f"{x['keyword']} | {x['filter']} | {x['value']} | ТИП={x.get('promotion_type','')} | {x.get('reason','')}\n")

    def show_promotion_type(self,promotion_type):
        if not self.res:return
        self.stats.delete("1.0","end")
        for x in self.res["rows"]:
            if x.get("promotion_type")==promotion_type:
                self.stats.insert("end",f"{x['keyword']} | {x['filter']} | {x['value']} | ТИП={x.get('promotion_type','')} | STATUS={x.get('status','')} | PRODUCT={x.get('product_domains_count',0)} | SEO={x.get('seo_domains_count',0)} | {x.get('reason','')}\n")

    def url_audit_path(self):
        return Path(self.folder.get() or os.getcwd()) / "seo_url_check_audit.csv"

    def _load_url_audit(self):
        path = self.url_audit_path()
        if not path.exists():
            return {}
        try:
            df = pd.read_csv(path, sep=";", encoding="utf-8-sig", dtype=str).fillna("")
            required = {"category_url", "keyword", "target_url", "url_status"}
            if not required.issubset(df.columns):
                return {}
            result = {}
            for _, r in df.iterrows():
                key = (str(r.get("category_url", "")).strip(), str(r.get("keyword", "")).strip(), str(r.get("target_url", "")).strip())
                if not key[1] or not key[2]:
                    continue
                result[key] = r.to_dict()
            return result
        except Exception:
            return {}

    def _save_url_audit(self, audit):
        path = self.url_audit_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        rows = list(audit.values())
        if not rows:
            return
        df = pd.DataFrame(rows)
        cols = [
            "category_url", "keyword", "category", "category_id", "filter", "value",
            "target_url", "url_status", "http_status", "final_url", "redirect",
            "checked_at", "verification_state", "error"
        ]
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        df = df[cols]
        df.to_csv(path, sep=";", index=False, encoding="utf-8-sig", lineterminator="\n")

    def _apply_url_audit(self, items, audit):
        applied = 0
        for x in items:
            key = (str(self.res.get("url", "")).strip(), str(x.get("keyword", "")).strip(), str(x.get("target_url", "")).strip())
            saved = audit.get(key)
            if not saved:
                continue
            state = str(saved.get("verification_state", "")).strip()
            if state != "VERIFIED" or str(saved.get("verification_version", "")) != "5":
                continue
            for field in ("url_status", "http_status", "redirect_location", "final_http_status", "final_url", "redirect", "checked_at", "verification_version", "existing_seo_url", "existing_seo_alias", "existing_seo_source"):
                if field in saved:
                    x[field] = saved.get(field, "")
            saved_status = str(saved.get("url_status", "")).strip()
            if saved_status.startswith("EXISTS"):
                x["status"] = "EXISTS"
            applied += 1
        return applied

    def check_urls(self):
        if not self.res:
            messagebox.showwarning("Нет данных", "Сначала выполните анализ.")
            return
        if self.url_check_running:
            return
        self.url_check_running=True
        self.url_check_stop_event=threading.Event()
        self.url_check_btn.config(state="disabled")
        self.url_stop_btn.config(state="normal")
        self.urllog.delete("1.0","end")
        self.urllog.insert("end", "Запуск проверки CREATE URL...\n")
        threading.Thread(target=self.worker_urls,daemon=True).start()

    def stop_url_check(self):
        if self.url_check_running:
            self.url_check_stop_event.set()
            self.urlstatus.config(text="Остановка проверки...")
            self.urllog.insert("end", "\nЗапрошена остановка. Уже проверенные строки сохранены.\n")
            self.urllog.see("end")

    def worker_urls(self):
        try:
            s=http_session()
            items=[x for x in self.res["rows"] if x.get("status")=="CREATE"]
            total=len(items)
            audit=self._load_url_audit()
            reused=self._apply_url_audit(items, audit)
            pending=[]
            category_url=str(self.res.get("url", "")).strip()

            # Авторизованный список OCFilter нужен для поиска старых SEO-страниц
            # с другим ЧПУ. Например /list-08h17t/ вместо /08h17t/.
            oc_records=[]
            oc_session=None
            oc_token=None
            try:
                cfg=load_cfg()
                ocfg=cfg.get("ocfilter",{}) if isinstance(cfg.get("ocfilter",{}),dict) else {}
                bu=cfg.get("HTTP Basic login",ocfg.get("HTTP Basic login",""))
                bp=cfg.get("HTTP Basic password",ocfg.get("HTTP Basic password",""))
                ou=cfg.get("OpenCart login",ocfg.get("OpenCart login",""))
                op=cfg.get("OpenCart password",ocfg.get("OpenCart password",""))
                if all(str(v or "").strip() for v in (bu,bp,ou,op)):
                    oc_session=http_session()
                    oc_token=oc_login(oc_session,bu,bp,ou,op)
                    oc_records=oc_existing_page_records(oc_session,oc_token)
                    self.q.put(("urllog", f"OCFilter: загружено существующих SEO-страниц: {len(oc_records)}"))
                else:
                    self.q.put(("urllog", "OCFilter: авторизация не заполнена - проверяется только предполагаемый URL."))
            except Exception as e:
                self.q.put(("urllog", f"OCFilter: не удалось получить список существующих страниц: {e}"))

            for x in items:
                key=(category_url, str(x.get("keyword","")).strip(), str(x.get("target_url","")).strip())
                saved=audit.get(key)
                if saved and str(saved.get("verification_state","")).strip()=="VERIFIED" and str(saved.get("verification_version","")) == "5":
                    continue
                pending.append(x)

            path=self.url_audit_path()
            self.q.put(("urllog", f"Файл проверки: {path}"))
            self.q.put(("urllog", f"CREATE всего: {total} | уже проверено ранее: {reused} | осталось проверить: {len(pending)}"))

            if not pending:
                self.q.put(("urldone", {"checked": total, "total": total, "stopped": False, "reused": reused, "path": str(path)}))
                return

            checked_new=0
            for x in pending:
                if self.url_check_stop_event.is_set():
                    break
                u=x.get("target_url") or ""
                now=time.strftime("%Y-%m-%d %H:%M:%S")
                try:
                    if not u:
                        x["url_status"]="NO_URL"
                        x["verification_state"]="VERIFIED"
                        x["verification_version"]="5"
                        x["checked_at"]=now
                        checked_new += 1
                        self.q.put(("urllog", f"[{reused+checked_new}/{total}] {x['keyword']} -> URL отсутствует | СОХРАНЕНО"))
                        continue

                    self.q.put(("urllog", f"[{reused+checked_new+1}/{total}] ПРОВЕРЯЮ: {u}"))
                    initial_status, redirect_location, final_status, final_url, redirected = check_public_url(s,u)
                    x["http_status"]=initial_status
                    x["redirect_location"]=redirect_location
                    x["final_http_status"]=final_status
                    x["final_url"]=final_url
                    x["redirect"]=redirected
                    x["existing_seo_url"]=""
                    x["existing_seo_alias"]=""
                    x["existing_seo_source"]=""

                    # 1. Прямой URL.
                    # Критическое правило: HTTP 301 означает, что URL уже занят.
                    # Неважно, какой статус отдаёт конечный URL после редиректа.
                    if initial_status == 200:
                        x["url_status"]="ALREADY_EXISTS"
                        x["status"]="EXISTS"
                    elif initial_status == 301:
                        x["url_status"]="EXISTS_301"
                        x["status"]="EXISTS"
                    elif initial_status == 404:
                        x["url_status"]="READY_FOR_OCFILTER"

                        # 2. Важная дополнительная проверка:
                        # существующая SEO-страница может иметь другое ЧПУ.
                        if oc_session and oc_token and oc_records:
                            old=oc_find_existing_seo_page(oc_session,oc_token,x,oc_records,category_url)
                            if old:
                                old_url=old.get("existing_url","")
                                self.q.put(("urllog", f"[{reused+checked_new+1}/{total}] ПРОВЕРЯЮ СТАРОЕ ЧПУ: {old_url}"))
                                try:
                                    os0, oloc, ofs, ofinal, ored=check_public_url(s,old_url)
                                except Exception:
                                    os0,oloc,ofs,ofinal,ored=(0,"",0,old_url,False)
                                # Для найденной старой SEO-страницы 301 также означает, что страница существует.
                                if os0 in (200, 301):
                                    x["url_status"]="EXISTS_ALTERNATE_SEO_301" if os0 == 301 else "EXISTS_ALTERNATE_SEO"
                                    x["status"]="EXISTS"
                                    x["existing_seo_url"]=ofinal if ored else old_url
                                    x["existing_seo_alias"]=old.get("alias","")
                                    x["existing_seo_source"]=old.get("source_text","")
                                    self.q.put(("urllog", f"[{reused+checked_new+1}/{total}] {x['keyword']} -> НАЙДЕНА СУЩЕСТВУЮЩАЯ SEO-СТРАНИЦА: {x['existing_seo_url']} | alias={x['existing_seo_alias']}"))
                    elif redirected and final_status == 404:
                        x["url_status"]="REDIRECT_TARGET_404"
                    elif redirected:
                        x["url_status"]=f"REDIRECT_TARGET_HTTP_{final_status}"
                    else:
                        x["url_status"]=f"HTTP_{initial_status}"

                    x["verification_state"]="VERIFIED"
                    x["verification_version"]="5"
                    x["checked_at"]=now
                    x["url_error"]=""
                    checked_new += 1

                    key=(category_url, str(x.get("keyword","")).strip(), str(x.get("target_url","")).strip())
                    audit[key]={
                        "category_url":category_url,"keyword":x.get("keyword",""),"category":x.get("category",""),
                        "category_id":x.get("category_id",""),"filter":x.get("filter",""),"value":x.get("value",""),
                        "target_url":u,"url_status":x.get("url_status",""),"http_status":initial_status,
                        "redirect_location":redirect_location,"final_http_status":final_status,"final_url":final_url,
                        "redirect":redirected,"existing_seo_url":x.get("existing_seo_url",""),
                        "existing_seo_alias":x.get("existing_seo_alias",""),"existing_seo_source":x.get("existing_seo_source",""),
                        "checked_at":now,"verification_state":"VERIFIED","verification_version":"5","error":""
                    }
                    self._save_url_audit(audit)
                    redir_txt=f" | Location: {redirect_location}" if redirect_location else ""
                    extra=f" | OLD SEO: {x.get('existing_seo_url')}" if x.get("existing_seo_url") else ""
                    self.q.put(("urllog", f"[{reused+checked_new}/{total}] {x['keyword']} -> {u} | HTTP {initial_status} | FINAL {final_status} | {x['url_status']}{redir_txt}{extra} | СОХРАНЕНО"))
                except Exception as e:
                    x["url_status"]="ERROR"
                    x["url_error"]=str(e)
                    x["verification_state"]="ERROR"
                    checked_new += 1
                    self.q.put(("urllog", f"[{reused+checked_new}/{total}] {x['keyword']} -> ERROR: {e} | НЕ ПОМЕЧЕНО КАК ПРОВЕРЕНО"))
                self.q.put(("urlstatus", f"Проверено: {min(reused+checked_new,total)} из {total}"))

            self._save_url_audit(audit)
            stopped=self.url_check_stop_event.is_set()
            self.q.put(("urldone", {"checked": min(reused+checked_new,total), "total": total, "stopped": stopped, "reused": reused, "path": str(path)}))
        except Exception as e:
            self.q.put(("urlerr",str(e)))
        finally:
            self.url_check_running=False


    def ocfilter_audit_path(self):
        return Path(self.folder.get() or os.getcwd()) / "seo_ocfilter_check_audit.json"

    def _save_ocfilter_audit(self, results):
        """Сохраняет результат DRY-RUN OCFilter, чтобы он не терялся между запусками."""
        try:
            path = self.ocfilter_audit_path()
            old = {}
            if path.exists():
                try:
                    old = json.loads(path.read_text(encoding="utf-8"))
                except Exception:
                    old = {}
            if not isinstance(old, dict):
                old = {}
            category_url = str(self.res.get("url", "")).strip() if self.res else ""
            now = time.strftime("%Y-%m-%d %H:%M:%S")
            for r in results:
                key = "|".join([
                    category_url,
                    norm(r.get("keyword", "")),
                    norm(r.get("filter", "")),
                    canonical_value(r.get("value", "")),
                ])
                old[key] = {
                    "category_url": category_url,
                    "keyword": r.get("keyword", ""),
                    "filter": r.get("filter", ""),
                    "value": r.get("value", ""),
                    "target_url": r.get("target_url", ""),
                    "url_status": r.get("url_status", ""),
                    "oc_status": r.get("oc_status", ""),
                    "oc_message": r.get("oc_message", ""),
                    "oc_category_id": r.get("oc_category_id", ""),
                    "oc_category_found": r.get("oc_category_found", ""),
                    "oc_category_score": r.get("oc_category_score", ""),
                    "alias": r.get("alias", ""),
                    "verified_at": now,
                    "verification_state": "VERIFIED",
                    "verification_version": "2",
                }
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(old, ensure_ascii=False, indent=2), encoding="utf-8")
            return path
        except Exception:
            return None

    def _apply_ocfilter_audit(self):
        """Подтягивает сохраненные результаты OCFilter в текущий анализ."""
        if not self.res:
            return 0
        try:
            path = self.ocfilter_audit_path()
            if not path.exists():
                return 0
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                return 0
            category_url = str(self.res.get("url", "")).strip()
            applied = 0
            for row in self.res.get("rows", []):
                key = "|".join([
                    category_url,
                    norm(row.get("keyword", "")),
                    norm(row.get("filter", "")),
                    canonical_value(row.get("value", "")),
                ])
                saved = data.get(key)
                if not isinstance(saved, dict):
                    continue
                # Не переносим старый результат на другой target URL.
                saved_target = str(saved.get("target_url", "")).strip()
                current_target = str(row.get("target_url", "")).strip()
                if saved_target and current_target and saved_target != current_target:
                    continue
                for field in (
                    "url_status", "oc_status", "oc_message", "oc_category_id",
                    "oc_category_found", "oc_category_score", "alias", "target_url"
                ):
                    if field in saved and saved.get(field) not in (None, ""):
                        row[field] = saved.get(field)
                applied += 1
            return applied
        except Exception:
            return 0

    def run_oc(self):
        if not self.res:
            messagebox.showwarning("Нет данных","Сначала выполните анализ и проверку URL.")
            return
        if getattr(self, "oc_running", False):
            return
        self.save_settings()
        creds=(self.oc["HTTP Basic login"].get().strip(),self.oc["HTTP Basic password"].get(),self.oc["OpenCart login"].get().strip(),self.oc["OpenCart password"].get())
        if not all(creds):
            messagebox.showwarning("OCFilter","Заполните все 4 поля авторизации OCFilter.")
            return
        self.oc_running=True
        self.oc_check_btn.config(state="disabled")
        self.oclog.delete("1.0","end")
        self.oclog.insert("end","DRY-RUN OCFilter запущен. Ничего на сайте не изменяется.\n\n")
        threading.Thread(target=self.worker_oc,args=(creds,),daemon=True).start()

    def worker_oc(self,creds):
        try:
            results=Pipeline().ocfilter(self.res,creds,create=False,progress=lambda msg:self.q.put(("oclog",msg)))
            self.q.put(("ocdone",results))
        except Exception as e:
            self.q.put(("ocerr",str(e)))
        finally:
            self.oc_running=False

    def worker_urls(self):
        try:
            s=http_session()
            items=[x for x in self.res["rows"] if x.get("status")=="CREATE"]
            total=len(items)
            audit=self._load_url_audit()
            reused=self._apply_url_audit(items, audit)
            pending=[]
            category_url=str(self.res.get("url", "")).strip()

            # Авторизованный список OCFilter нужен для поиска старых SEO-страниц
            # с другим ЧПУ. Например /list-08h17t/ вместо /08h17t/.
            oc_records=[]
            oc_session=None
            oc_token=None
            try:
                cfg=load_cfg()
                ocfg=cfg.get("ocfilter",{}) if isinstance(cfg.get("ocfilter",{}),dict) else {}
                bu=cfg.get("HTTP Basic login",ocfg.get("HTTP Basic login",""))
                bp=cfg.get("HTTP Basic password",ocfg.get("HTTP Basic password",""))
                ou=cfg.get("OpenCart login",ocfg.get("OpenCart login",""))
                op=cfg.get("OpenCart password",ocfg.get("OpenCart password",""))
                if all(str(v or "").strip() for v in (bu,bp,ou,op)):
                    oc_session=http_session()
                    oc_token=oc_login(oc_session,bu,bp,ou,op)
                    oc_records=oc_existing_page_records(oc_session,oc_token)
                    self.q.put(("urllog", f"OCFilter: загружено существующих SEO-страниц: {len(oc_records)}"))
                else:
                    self.q.put(("urllog", "OCFilter: авторизация не заполнена - проверяется только предполагаемый URL."))
            except Exception as e:
                self.q.put(("urllog", f"OCFilter: не удалось получить список существующих страниц: {e}"))

            for x in items:
                key=(category_url, str(x.get("keyword","")).strip(), str(x.get("target_url","")).strip())
                saved=audit.get(key)
                if saved and str(saved.get("verification_state","")).strip()=="VERIFIED" and str(saved.get("verification_version","")) == "5":
                    continue
                pending.append(x)

            path=self.url_audit_path()
            self.q.put(("urllog", f"Файл проверки: {path}"))
            self.q.put(("urllog", f"CREATE всего: {total} | уже проверено ранее: {reused} | осталось проверить: {len(pending)}"))

            if not pending:
                self.q.put(("urldone", {"checked": total, "total": total, "stopped": False, "reused": reused, "path": str(path)}))
                return

            checked_new=0
            for x in pending:
                if self.url_check_stop_event.is_set():
                    break
                u=x.get("target_url") or ""
                now=time.strftime("%Y-%m-%d %H:%M:%S")
                try:
                    if not u:
                        x["url_status"]="NO_URL"
                        x["verification_state"]="VERIFIED"
                        x["verification_version"]="5"
                        x["checked_at"]=now
                        checked_new += 1
                        self.q.put(("urllog", f"[{reused+checked_new}/{total}] {x['keyword']} -> URL отсутствует | СОХРАНЕНО"))
                        continue

                    self.q.put(("urllog", f"[{reused+checked_new+1}/{total}] ПРОВЕРЯЮ: {u}"))
                    initial_status, redirect_location, final_status, final_url, redirected = check_public_url(s,u)
                    x["http_status"]=initial_status
                    x["redirect_location"]=redirect_location
                    x["final_http_status"]=final_status
                    x["final_url"]=final_url
                    x["redirect"]=redirected
                    x["existing_seo_url"]=""
                    x["existing_seo_alias"]=""
                    x["existing_seo_source"]=""

                    # 1. Прямой URL.
                    # Критическое правило: HTTP 301 означает, что URL уже занят.
                    # Неважно, какой статус отдаёт конечный URL после редиректа.
                    if initial_status == 200:
                        x["url_status"]="ALREADY_EXISTS"
                        x["status"]="EXISTS"
                    elif initial_status == 301:
                        x["url_status"]="EXISTS_301"
                        x["status"]="EXISTS"
                    elif initial_status == 404:
                        x["url_status"]="READY_FOR_OCFILTER"

                        # 2. Важная дополнительная проверка:
                        # существующая SEO-страница может иметь другое ЧПУ.
                        if oc_session and oc_token and oc_records:
                            old=oc_find_existing_seo_page(oc_session,oc_token,x,oc_records,category_url)
                            if old:
                                old_url=old.get("existing_url","")
                                self.q.put(("urllog", f"[{reused+checked_new+1}/{total}] ПРОВЕРЯЮ СТАРОЕ ЧПУ: {old_url}"))
                                try:
                                    os0, oloc, ofs, ofinal, ored=check_public_url(s,old_url)
                                except Exception:
                                    os0,oloc,ofs,ofinal,ored=(0,"",0,old_url,False)
                                # Для найденной старой SEO-страницы 301 также означает, что страница существует.
                                if os0 in (200, 301):
                                    x["url_status"]="EXISTS_ALTERNATE_SEO_301" if os0 == 301 else "EXISTS_ALTERNATE_SEO"
                                    x["status"]="EXISTS"
                                    x["existing_seo_url"]=ofinal if ored else old_url
                                    x["existing_seo_alias"]=old.get("alias","")
                                    x["existing_seo_source"]=old.get("source_text","")
                                    self.q.put(("urllog", f"[{reused+checked_new+1}/{total}] {x['keyword']} -> НАЙДЕНА СУЩЕСТВУЮЩАЯ SEO-СТРАНИЦА: {x['existing_seo_url']} | alias={x['existing_seo_alias']}"))
                    elif redirected and final_status == 404:
                        x["url_status"]="REDIRECT_TARGET_404"
                    elif redirected:
                        x["url_status"]=f"REDIRECT_TARGET_HTTP_{final_status}"
                    else:
                        x["url_status"]=f"HTTP_{initial_status}"

                    x["verification_state"]="VERIFIED"
                    x["verification_version"]="5"
                    x["checked_at"]=now
                    x["url_error"]=""
                    checked_new += 1

                    key=(category_url, str(x.get("keyword","")).strip(), str(x.get("target_url","")).strip())
                    audit[key]={
                        "category_url":category_url,"keyword":x.get("keyword",""),"category":x.get("category",""),
                        "category_id":x.get("category_id",""),"filter":x.get("filter",""),"value":x.get("value",""),
                        "target_url":u,"url_status":x.get("url_status",""),"http_status":initial_status,
                        "redirect_location":redirect_location,"final_http_status":final_status,"final_url":final_url,
                        "redirect":redirected,"existing_seo_url":x.get("existing_seo_url",""),
                        "existing_seo_alias":x.get("existing_seo_alias",""),"existing_seo_source":x.get("existing_seo_source",""),
                        "checked_at":now,"verification_state":"VERIFIED","verification_version":"5","error":""
                    }
                    self._save_url_audit(audit)
                    redir_txt=f" | Location: {redirect_location}" if redirect_location else ""
                    extra=f" | OLD SEO: {x.get('existing_seo_url')}" if x.get("existing_seo_url") else ""
                    self.q.put(("urllog", f"[{reused+checked_new}/{total}] {x['keyword']} -> {u} | HTTP {initial_status} | FINAL {final_status} | {x['url_status']}{redir_txt}{extra} | СОХРАНЕНО"))
                except Exception as e:
                    x["url_status"]="ERROR"
                    x["url_error"]=str(e)
                    x["verification_state"]="ERROR"
                    checked_new += 1
                    self.q.put(("urllog", f"[{reused+checked_new}/{total}] {x['keyword']} -> ERROR: {e} | НЕ ПОМЕЧЕНО КАК ПРОВЕРЕНО"))
                self.q.put(("urlstatus", f"Проверено: {min(reused+checked_new,total)} из {total}"))

            self._save_url_audit(audit)
            stopped=self.url_check_stop_event.is_set()
            self.q.put(("urldone", {"checked": min(reused+checked_new,total), "total": total, "stopped": stopped, "reused": reused, "path": str(path)}))
        except Exception as e:
            self.q.put(("urlerr",str(e)))
        finally:
            self.url_check_running=False

    def show_ready_filter_urls(self):
        """Показывает только исходные URL результатов фильтрации,
        которые относятся к страницам, прошедшим проверку и готовым
        для внешнего генератора. Один URL в одной строке, без нумерации.
        """
        self.export_log.delete("1.0", "end")
        if not self.res:
            return

        self._apply_ocfilter_audit()
        urls = []
        seen = set()
        for row in generator_rows(self.res):
            url = str(row.get("filter_result_url") or "").strip()
            if not url or url in seen:
                continue
            seen.add(url)
            urls.append(url)

        if urls:
            self.export_log.insert("1.0", "\n".join(urls))
        else:
            rows = self.res.get("rows", [])
            create = sum(1 for x in rows if x.get("status") == "CREATE")
            seo_create = sum(
                1 for x in rows
                if x.get("status") == "CREATE"
                and x.get("promotion_type") == "SEO-СТРАНИЦА"
            )
            url_ready = sum(
                1 for x in rows
                if x.get("status") == "CREATE"
                and x.get("promotion_type") == "SEO-СТРАНИЦА"
                and x.get("url_status") in ("READY_FOR_OCFILTER", "HTTP_404")
            )
            oc_ready = sum(
                1 for x in rows
                if x.get("status") == "CREATE"
                and x.get("promotion_type") == "SEO-СТРАНИЦА"
                and x.get("url_status") in ("READY_FOR_OCFILTER", "HTTP_404")
                and x.get("oc_status") == "READY"
            )
            oc_statuses = {}
            for x in rows:
                if x.get("status") != "CREATE" or x.get("promotion_type") != "SEO-СТРАНИЦА":
                    continue
                st = str(x.get("oc_status", "")) or "EMPTY"
                oc_statuses[st] = oc_statuses.get(st, 0) + 1
            oc_status_text = ", ".join(f"{k}={v}" for k,v in sorted(oc_statuses.items())) or "нет данных"
            self.export_log.insert(
                "1.0",
                "Готовых URL фильтрации пока нет.\n\n"
                f"CREATE: {create}\n"
                f"SEO + CREATE: {seo_create}\n"
                f"URL прошли проверку: {url_ready}\n"
                f"OCFilter READY: {oc_ready}\n"
                f"OCFilter статусы: {oc_status_text}\n\n"
                "Правило экспорта: status=CREATE + promotion_type=SEO-СТРАНИЦА "
                "+ URL=404/READY + OCFilter=READY.\n"
                "Если OCFilter READY=0, ссылки намеренно не выводятся."
            )

    def copy_ready_filter_urls(self):
        """Копирует только готовые исходные URL фильтрации, по одному в строке."""
        self.show_ready_filter_urls()
        value=self.export_log.get("1.0", "end-1c").strip()
        if not value:
            messagebox.showwarning("Нет URL", "Нет URL результатов фильтрации, прошедших проверку.")
            return
        self.clipboard_clear()
        self.clipboard_append(value)
        self.update()

    def export_generator(self):
        if not self.res:
            messagebox.showwarning("Нет данных","Сначала выполните анализ категории.")
            return
        self._apply_ocfilter_audit()
        ready = generator_rows(self.res)
        if not ready:
            messagebox.showwarning("Нет готовых страниц","Сначала выполните Проверку URL и Проверку OCFilter (DRY-RUN).")
            return
        p=filedialog.asksaveasfilename(defaultextension=".csv",initialfile="seo_pages_for_generator.csv",filetypes=[("CSV","*.csv")])
        if not p:
            return
        try:
            n=export_generator_csv(self.res,p)
            messagebox.showinfo("Готово",f"Сохранено страниц: {n}\n\nФайл для генератора:\n{p}")
        except Exception as e:
            messagebox.showerror("Экспорт",str(e))

    def export(self):
        if not self.res:messagebox.showwarning("Нет данных","Сначала выполните анализ.");return
        self._apply_ocfilter_audit()
        p=filedialog.asksaveasfilename(defaultextension=".xlsx",initialfile="fe_rus_seo_pages.xlsx",filetypes=[("Excel","*.xlsx")])
        if p:
            export_excel(self.res,p,self.topdf); messagebox.showinfo("Готово","Файл сохранен:\n"+p)

    def poll(self):
        try:
            while True:
                t,x=self.q.get_nowait()
                if t=="log":self.logtext.insert("end",str(x)+"\n");self.logtext.see("end")
                elif t=="cat":
                    self.res=x
                    self.render_filter_checkboxes(x.get("options", []))
                    self.catstatus.config(text=f"Готово: {x['category']} | category_id={x['category_id']} | вариантов={len(x['rows'])}")
                    self.logtext.insert("end","\nГотово. Отметьте нужные характеристики и переходите в WordKeeper.\n")
                elif t=="err":messagebox.showerror("Ошибка",x)
                elif t=="wklog":self.wklog.insert("end",str(x)+"\n");self.wklog.see("end")
                elif t=="wkdone":
                    p=x; self.topdf=pd.read_csv(p,encoding="utf-8-sig") if Path(p).exists() and Path(p).stat().st_size else pd.DataFrame(); self.wkstatus.config(text=f"Готово: {Path(p).name} | строк: {len(self.topdf)}"); self.wk_start_btn.config(state="normal"); self.wk_stop_btn.config(state="disabled"); self.wklog.insert("end",f"\nФайл TOP-30: {p}\n")
                elif t=="wkstop":
                    p=x["path"]; self.topdf=pd.read_csv(p,encoding="utf-8-sig") if Path(p).exists() and Path(p).stat().st_size else pd.DataFrame(); self.wkstatus.config(text=f"ОСТАНОВЛЕНО: {len(self.topdf)} строк | {Path(p).name}"); self.wk_start_btn.config(state="normal"); self.wk_stop_btn.config(state="disabled"); self.wklog.insert("end",f"\nАНАЛИЗ ОСТАНОВЛЕН.\nTOP-30: {p}\nCheckpoint: {x['progress']}\nЗавершено запросов: {x['done']} из {x['total']}\nМожно нажать ПОЛУЧИТЬ TOP-30 снова - продолжит с последнего сохранённого запроса.\n")
                elif t=="wkerr":
                    self.wk_running=False; self.wk_start_btn.config(state="normal"); self.wk_stop_btn.config(state="disabled"); messagebox.showerror("WordKeeper",x)
                elif t=="analysislog":
                    self.stats.insert("end",str(x)+"\n"); self.stats.see("end")
                elif t=="analysisdone":
                    self.res=x["res"]
                    self.analysis_running=False
                    self.compare_btn.config(state="normal")
                    self.analysis_stop_btn.config(state="disabled")
                    self._render_analysis()
                    cache_txt=x.get("cache","")
                    if x.get("stopped"):
                        self.stats.insert("end",f"\nАНАЛИЗ ОСТАНОВЛЕН. Кэш классификации сохранен: {cache_txt}\n")
                    else:
                        self.stats.insert("end",f"\nКлассификация сохранена: {cache_txt}\n")
                    self.stats.see("end")
                elif t=="analysiserr":
                    self.analysis_running=False
                    self.compare_btn.config(state="normal")
                    self.analysis_stop_btn.config(state="disabled")
                    messagebox.showerror("Анализ",x)
                elif t=="urllog":self.urllog.insert("end",str(x)+"\n");self.urllog.see("end")
                elif t=="urlstatus":self.urlstatus.config(text=str(x))
                elif t=="urldone":
                    self.url_check_running=False
                    self.url_check_btn.config(state="normal")
                    self.url_stop_btn.config(state="disabled")
                    msg=f"Проверено: {x['checked']} из {x['total']}"
                    if x.get("stopped"): msg += " | ОСТАНОВЛЕНО"
                    self.urlstatus.config(text=msg)
                    self.urllog.insert("end",f"\n{msg}\n")
                elif t=="urlerr":
                    self.url_check_running=False
                    self.url_check_btn.config(state="normal")
                    self.url_stop_btn.config(state="disabled")
                    messagebox.showerror("Проверка URL",x)
                elif t=="oclog":
                    self.oclog.insert("end",str(x)+"\n"); self.oclog.see("end")
                elif t=="ocdone":
                    self.oc_running=False
                    self.oc_check_btn.config(state="normal")
                    rows_by_key = {}
                    for rr in self.res.get("rows", []):
                        rows_by_key.setdefault(result_row_key(rr), []).append(rr)
                    matched = 0
                    for r in x:
                        candidates = rows_by_key.get(result_row_key(r), [])
                        if len(candidates) == 1:
                            candidates[0].update(r)
                            matched += 1
                        elif candidates:
                            # Резерв: если ключ совпал несколько раз, target_url
                            # позволяет выбрать конкретную строку.
                            target = str(r.get("target_url", "")).strip()
                            found = next((rr for rr in candidates if str(rr.get("target_url", "")).strip() == target), None)
                            if found is not None:
                                found.update(r)
                                matched += 1
                    audit_path = self._save_ocfilter_audit(x)
                    self._apply_ocfilter_audit()
                    ready = len(generator_rows(self.res))
                    statuses = {}
                    for rr in self.res.get("rows", []):
                        st = str(rr.get("oc_status", "")) or "EMPTY"
                        statuses[st] = statuses.get(st, 0) + 1
                    status_text = ", ".join(f"{k}={v}" for k,v in sorted(statuses.items()))
                    self.oclog.insert("end",f"\nГотово. Проверено CREATE: {len(x)} | сопоставлено с анализом: {matched}\nГОТОВО ДЛЯ ГЕНЕРАТОРА: {ready}\nOCFilter статусы: {status_text}\n")
                    if audit_path:
                        self.oclog.insert("end",f"Проверка OCFilter сохранена: {audit_path}\n")
                    self.oclog.insert("end","\n")
                    for r in x:self.oclog.insert("end",f"{r.get('keyword')} -> {r.get('oc_status')} | {r.get('target_url','')}\n")
                    self.oclog.insert("end",f"\nСледующий шаг: вкладка 6. Экспорт -> СОХРАНИТЬ CSV ДЛЯ ГЕНЕРАТОРА.\n")
                    self.oclog.see("end")
                elif t=="ocerr":
                    self.oc_running=False
                    self.oc_check_btn.config(state="normal")
                    self.oclog.insert("end",f"\nОШИБКА: {x}\n")
                    self.oclog.see("end")
                    messagebox.showerror("OCFilter",x)
        except queue.Empty:pass
        self.after(100,self.poll)


if __name__ == "__main__":
    App().mainloop()
