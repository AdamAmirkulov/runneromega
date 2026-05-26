# -*- coding: utf-8 -*-
import os
import re
import time
import base64
import queue
import shutil
import tempfile
import threading
from html import escape

import extract_msg
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import sys
import io

import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--workdir', default=None)
parser.add_argument('--source_folder', default=None)
parser.add_argument('--company_id', default=None)
parser.add_argument('--workers', default=None,
                    help="Сколько параллельных Chrome (по умолчанию 4)")
args, _ = parser.parse_known_args()
if not args.company_id or not str(args.company_id).strip():
    print("❌ ОШИБКА: не передан --company_id — компания не определена, запуск остановлен.")
    sys.exit(1)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    N_WORKERS = max(1, min(8, int(str(args.workers).strip())))
except (TypeError, ValueError):
    N_WORKERS = 4

# --- Пути ---
MSG_FOLDER = args.source_folder
OUT_FOLDER = MSG_FOLDER  # PDF сохраняются туда же

# --- Отправитель для блока "От кого" в зависимости от компании ---
COMPANY_SENDERS = {
    '1': u'Коллекторское агенство "A-Омега" <noreply@a-omega.kz>',
    '2': u'Коллекторское агенство "KPI" <no-reply@kkpi.kz>',
    '3': u'Коллекторское агенство "ORION" <no-reply@ka-orion.kz>',
    '4': u'Специальная финансовая компания "Invest Way KZ" <info@investway.kz>',
}
SENDER_NAME = COMPANY_SENDERS.get(str(args.company_id), COMPANY_SENDERS['1'])

os.makedirs(OUT_FOLDER, exist_ok=True)

# --- Драйвер Chrome ---
_user_data_dirs = []
_driver_path = [None]
_driver_path_lock = threading.Lock()


def _resolve_driver_path():
    """Путь к chromedriver резолвим ОДИН раз и лениво (не при импорте, не на
    каждый инстанс). cache_valid_range=7 — не ходить в интернет, если драйвер
    уже качали за последние 7 дней."""
    with _driver_path_lock:
        if _driver_path[0] is None:
            try:
                _driver_path[0] = ChromeDriverManager(cache_valid_range=7).install()
            except TypeError:
                _driver_path[0] = ChromeDriverManager().install()
        return _driver_path[0]


def make_driver():
    """Отдельный headless-Chrome со своим user-data-dir — параллельные
    инстансы не мешают друг другу."""
    udd = tempfile.mkdtemp(prefix="cvmsg_")
    _user_data_dirs.append(udd)
    opts = Options()
    opts.add_argument('--headless=new')
    opts.add_argument('--disable-gpu')
    opts.add_argument('--no-sandbox')
    opts.add_argument('--disable-dev-shm-usage')
    opts.add_argument('--user-data-dir=' + udd)
    return webdriver.Chrome(service=Service(_resolve_driver_path()), options=opts)


def render_pdf(driver, full_html):
    """Рендер статичного HTML в PDF без временных файлов и фиксированных
    пауз: грузим через data:-URI и ждём готовности документа (обычно
    мгновенно, максимум ~1 сек)."""
    b64 = base64.b64encode(full_html.encode('utf-8')).decode('ascii')
    driver.get("data:text/html;charset=utf-8;base64," + b64)
    for _ in range(50):
        try:
            if driver.execute_script("return document.readyState") == "complete":
                break
        except Exception:
            pass
        time.sleep(0.02)
    pdf_data = driver.execute_cdp_cmd("Page.printToPDF", {
        "printBackground": True,
        "marginTop": 0.6, "marginBottom": 0.6,
        "marginLeft": 0.6, "marginRight": 0.6,
        "paperWidth": 8.27, "paperHeight": 11.69,  # A4
    })
    return base64.b64decode(pdf_data['data'])

NBSP = u"\u00A0"

# --- Утилиты ---
def clean_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name or '').strip()

def html_to_text_with_lines(html):
    if not html:
        return ""
    html = html.replace(NBSP, " ")
    soup = BeautifulSoup(html, 'html.parser')
    for tag in soup.find_all(['p', 'div', 'span']):
        if not tag.get_text(strip=True):
            tag.decompose()
    for tag in soup.find_all(True):
        style = tag.get('style')
        if style and re.search(r'height\s*:\s*\d{2,4}px', style):
            tag.decompose()
    text = soup.get_text(separator='\n', strip=True).replace('\r', '\n')
    text = "\n".join(re.sub(r'\s+', ' ', ln).strip() for ln in text.split('\n'))
    return re.sub(r'\n{2,}', '\n', text).strip()

# Алфавит (рус/каз)
UPPER = r"A-ZА-ЯЁӘІҢҮҰҚӨҺ"
LOWER = r"a-zа-яёәіңүұқөһ"
LETTER = UPPER + LOWER
NAME_TOKEN = r"[{0}][{0}\-''`]+".format(LETTER)
NAME_SEQ   = r"(?:{0}(?:\s+{0}){{1,4}})".format(NAME_TOKEN)

def smart_title(s):
    s = re.sub(r'\s+', ' ', (s or '').strip())
    parts = re.split(r'(\s+|-)', s.lower())
    out = []
    for p in parts:
        if not p or p.isspace() or p == '-':
            out.append(p)
        else:
            out.append(p[0].upper() + p[1:])
    return ''.join(out)

def pick_iin_from_chunk(s):
    digits = re.sub(r'\D+', '', s or '')
    if not digits:
        return None
    if len(digits) < 12:
        digits = digits.zfill(12)
    else:
        digits = digits[:12]
    return digits

def normalize_iin_in_html(html):
    def repl(m):
        label = m.group(1)
        chunk = m.group(2)
        norm = pick_iin_from_chunk(chunk) or chunk
        return u"%s: %s" % (label, norm)
    return re.sub(r'((?:ИИН|ЖСН))\s*[:\-]?\s*((?:\d[\s\u00A0\-]?){5,40})',
                  repl, html or '', flags=re.IGNORECASE)

def extract_name_iin(text_with_lines):
    if not text_with_lines:
        return (None, None)
    text_flat = text_with_lines.replace('\n', ' ')

    pat_close_iin = r"(" + NAME_SEQ + r")\s*(?:ИИН|ЖСН)\s*[:\-]?\s*((?:\d[\s\u00A0\-]?){8,40})"
    m = re.search(pat_close_iin, text_flat, flags=re.IGNORECASE)
    if m:
        fio = m.group(1).strip()
        iin = pick_iin_from_chunk(m.group(2))
        return (fio, iin)

    lines = text_with_lines.split('\n')
    extracted_iin_any = None
    for idx, ln in enumerate(lines):
        mm = re.search(r'(?:ИИН|ЖСН)\s*[:\-]?\s*((?:\d[\s\u00A0\-]?){5,40})',
                       ln, flags=re.IGNORECASE)
        if not mm:
            continue
        iin_candidate = pick_iin_from_chunk(mm.group(1))
        if iin_candidate:
            extracted_iin_any = iin_candidate
        for j in (idx - 1, idx - 2):
            if 0 <= j < len(lines):
                if len(lines[j].strip()) <= 2:
                    continue
                mname = re.search(r"(" + NAME_SEQ + r")", lines[j])
                if mname:
                    fio = mname.group(1).strip()
                    return (fio, iin_candidate)

    m = re.search(r"(?:кому|кімге)\s*[:\-]?\s*(" + NAME_SEQ + r")",
                  text_flat, flags=re.IGNORECASE)
    if m:
        fio = m.group(1).strip()
        return (fio, extracted_iin_any)

    m = re.search(r"между\s+должником\s*[-–—]\s*(" + NAME_SEQ + r")\s+и\s+",
                  text_flat, flags=re.IGNORECASE)
    if m:
        fio = m.group(1).strip()
        return (fio, extracted_iin_any)

    if not extracted_iin_any:
        mm2 = re.search(r'(?:ИИН|ЖСН)\s*[:\-]?\s*((?:\d[\s\u00A0\-]?){5,40})',
                        text_flat, flags=re.IGNORECASE)
        if mm2:
            extracted_iin_any = pick_iin_from_chunk(mm2.group(1))

    return (None, extracted_iin_any)

def build_header_block(sender, recipient, date_sent, subject, cc=""):
    cc_line = u"<b>Копия:</b> {0}<br>".format(escape(cc)) if cc else ""
    return u"""<div style="font-size: 9pt; color: #333; margin-bottom: 10px;
         font-family: Arial, sans-serif;">
        <p><b>От кого:</b> {0}<br>
        <b>Кому:</b> {1}<br>
        {4}<b>Дата и время:</b> {2}<br>
        <b>Тема:</b> {3}</p>
    </div><hr>""".format(escape(sender or ""), escape(recipient or ""),
                          escape(date_sent or ""), escape(subject or ""), cc_line)

def inject_header(html, header_html):
    """Вставляет header_html сразу после открывающего <body>, либо в начало."""
    m = re.search(r'<body[^>]*>', html, re.IGNORECASE)
    if m:
        idx = m.end()
        return html[:idx] + header_html + html[idx:]
    return header_html + html

def clean_html_spacing(html):
    soup = BeautifulSoup(html or "", 'html.parser')
    for tag in soup.find_all(['p', 'div', 'span']):
        if not tag.get_text(strip=True):
            tag.decompose()
    for tag in soup.find_all(True):
        style = tag.get('style')
        if style and re.search(r'height\s*:\s*\d{2,4}px', style):
            tag.decompose()
    word_div = soup.find('div', {'class': 'WordSection1'})
    if word_div:
        html = str(word_div)
    else:
        html = str(soup)
    html = re.sub(r'(<br\s*/?>\s*){2,}', '<br>', html, flags=re.IGNORECASE)
    return html.replace('&nbsp;', ' ')

# --- Чтение .msg через extract_msg (корректно достаёт отправителя, дату,
#     получателей — без ручного разбора OLE-потоков) ---
def _s(v):
    """К строке, аккуратно (bytes -> декод, None -> '')."""
    if v is None:
        return ''
    if isinstance(v, bytes):
        for enc in ('utf-8', 'utf-16-le', 'cp1251', 'latin-1'):
            try:
                return v.decode(enc)
            except Exception:
                continue
        return v.decode('utf-8', 'replace')
    return str(v)


def read_msg_safe(msg_path):
    """Возвращает dict: from / to / cc / date / subject / body / htmlBody.

    Содержимое письма НЕ меняется — берём тело как есть из .msg."""
    result = {
        'from': '', 'to': '', 'cc': '',
        'date': '', 'subject': '', 'body': '', 'htmlBody': '',
    }

    with extract_msg.openMsg(msg_path) as msg:
        result['from'] = _s(msg.sender).strip()
        result['to'] = _s(msg.to).strip()
        result['cc'] = _s(msg.cc).strip()
        result['subject'] = _s(msg.subject).strip()
        result['body'] = _s(msg.body)

        html = msg.htmlBody
        result['htmlBody'] = _s(html)

        # Дата: extract_msg отдаёт datetime (из свойств письма или из
        # заголовков). Формат — «ДД.ММ.ГГГГ ЧЧ:ММ».
        d = msg.date
        if d is not None:
            try:
                result['date'] = d.strftime('%d.%m.%Y %H:%M')
            except Exception:
                result['date'] = _s(d)

    return result


# --- Обработка одного .msg (вызывается из воркеров) ---
_out_lock = threading.Lock()
_print_lock = threading.Lock()


def log(*a):
    with _print_lock:
        print(*a, flush=True)


def reserve_out_path(out_name):
    """Под локом «столбит» имя PDF: если такой файл уже есть — удаляет его
    (перезапись, как раньше), создаёт пустышку. Два параллельных воркера с
    одинаковым именем не затрут запись друг друга на полпути."""
    with _out_lock:
        p = os.path.join(OUT_FOLDER, out_name)
        if os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass
        open(p, "wb").close()
        return p


def process_one(driver, fname):
    msg_path = os.path.join(MSG_FOLDER, fname)
    data = read_msg_safe(msg_path)

    # «От кого» — реальный отправитель из письма; компания-заглушка только
    # если в .msg отправитель не указан.
    sender    = data['from'] or SENDER_NAME
    recipient = data['to'] or "Без получателя"
    cc        = data['cc']
    date_sent = data['date'] or "Без даты"
    subject   = data['subject'] or "Без темы"
    raw_text  = data['body'] or ""
    html_body = data['htmlBody'] or ""

    # убираем недоступные вне Outlook ресурсы (иначе в PDF будут битые
    # картинки), остальную вёрстку письма не трогаем — PDF должен быть
    # чистой копией письма
    html_render = re.sub(r'(src|href|background)=["\'](cid:|file:|data:)[^"\']*["\']',
                         '', html_body, flags=re.IGNORECASE)
    html_render = re.sub(r'<(img|link)[^>]*>', '', html_render, flags=re.IGNORECASE)

    # Word помечает тело письма как "div.WordSection1 { page:WordSection1 }".
    # CSS-свойство `page` форсит разрыв страницы перед этим блоком — из-за
    # этого наша шапка оставалась одна на 1-й странице, а само письмо
    # уезжало на 2-ю (большой пустой отступ). Убираем только это свойство,
    # вёрстку/поля письма не трогаем.
    html_render = re.sub(r'\bpage\s*:\s*(?:Word)?Section\d+\s*;?', '',
                         html_render, flags=re.IGNORECASE)

    # отдельная, агрессивно очищенная копия — только для извлечения ФИО/ИИН,
    # на итоговый PDF не влияет
    html_for_parsing = normalize_iin_in_html(clean_html_spacing(html_render or raw_text))

    text_for_parse = html_to_text_with_lines(html_for_parsing or raw_text)
    text_lower_flat = re.sub(r'\s+', ' ', text_for_parse).lower()
    fio, iin = extract_name_iin(text_for_parse)
    if not iin:
        mm = re.search(r'(?:ИИН|ЖСН)\s*[:\-]?\s*((?:\d[\s \-]?){5,40})',
                       text_lower_flat, re.IGNORECASE)
        if mm:
            iin = pick_iin_from_chunk(mm.group(1))

    if fio:
        fio = smart_title(fio)
    fio_for_file = clean_filename((fio or "Без_ФИО")).replace(" ", "_")
    iin_for_file = (iin or "Без_ИИН")

    if (re.search(r'уведомлен\w*\s+об\s+уступк', text_lower_flat) or
        "уведомление об уступке" in text_lower_flat or
        "уведомление" in fname.lower()):
        out_name = u"Уведомление об уступки, {0}, {1}.pdf".format(fio_for_file, iin_for_file)
    elif "досудебная претензия" in text_lower_flat:
        out_name = u"Досудебная претензия, {0}, {1}.pdf".format(fio_for_file, iin_for_file)
    else:
        out_name = u"{0}_{1}.pdf".format(fio_for_file, iin_for_file)
    out_name = clean_filename(out_name)

    if html_render.strip():
        full_html = html_render
        if not re.search(r'<meta[^>]+charset', full_html, re.IGNORECASE):
            if re.search(r'<head[^>]*>', full_html, re.IGNORECASE):
                full_html = re.sub(r'(<head[^>]*>)', r'\1<meta charset="utf-8">',
                                    full_html, count=1, flags=re.IGNORECASE)
            else:
                full_html = u'<meta charset="utf-8">' + full_html
        if not re.search(r'<body[^>]*>', full_html, re.IGNORECASE):
            full_html = (u'<html><head><meta charset="utf-8"><style>'
                         u'body { font-family: Arial, sans-serif; font-size: 10pt; '
                         u'padding: 1.5cm; }</style></head><body>' + full_html +
                         u'</body></html>')
    else:
        body_text = escape(raw_text or "").replace('\n', '<br>')
        full_html = (u'<html><head><meta charset="utf-8"><style>'
                     u'body { font-family: Arial, sans-serif; font-size: 10pt; '
                     u'line-height: 1.4; padding: 1.5cm; white-space: pre-wrap; }'
                     u'</style></head><body>' + body_text + u'</body></html>')

    header_block = build_header_block(sender, recipient, date_sent, subject, cc)
    full_html = inject_header(full_html, header_block)

    pdf_bytes = render_pdf(driver, full_html)

    out_pdf = reserve_out_path(out_name)
    with open(out_pdf, "wb") as f:
        f.write(pdf_bytes)
    log("✅", os.path.basename(out_pdf))

    # PDF на месте и не пустой — удаляем исходный .msg, чтобы он не
    # конвертировался повторно.
    if os.path.getsize(out_pdf) > 0:
        try:
            os.remove(msg_path)
        except Exception as e:
            log("⚠ Не удалось удалить .msg:", fname, e)


# --- Параллельная обработка ---
msg_files = [f for f in os.listdir(MSG_FOLDER) if f.lower().endswith('.msg')]
if not msg_files:
    print("Нет .msg файлов в папке — нечего конвертировать.", flush=True)
    sys.exit(0)

n_workers = min(N_WORKERS, len(msg_files))
print("Файлов .msg: {0} | параллельных Chrome: {1}".format(len(msg_files), n_workers),
      flush=True)

work_q = queue.Queue()
for f in msg_files:
    work_q.put(f)

_stats = {'ok': 0, 'err': 0}
_stats_lock = threading.Lock()


def worker(wid):
    try:
        driver = make_driver()
    except Exception as e:
        log("[w{0}] не удалось запустить Chrome: {1}".format(wid, e))
        return
    try:
        while True:
            try:
                fname = work_q.get_nowait()
            except queue.Empty:
                break
            try:
                process_one(driver, fname)
                with _stats_lock:
                    _stats['ok'] += 1
            except Exception as e:
                log("⛔ Ошибка при обработке {0}: {1}".format(fname, e))
                with _stats_lock:
                    _stats['err'] += 1
            finally:
                work_q.task_done()
    finally:
        try:
            driver.quit()
        except Exception:
            pass


_threads = [threading.Thread(target=worker, args=(i + 1,), daemon=True)
            for i in range(n_workers)]
for _t in _threads:
    _t.start()
for _t in _threads:
    _t.join()

for _d in _user_data_dirs:
    shutil.rmtree(_d, ignore_errors=True)

print("\n\U0001F389 Готово. Успешно: {0}, ошибок: {1}".format(_stats['ok'], _stats['err']),
      flush=True)
