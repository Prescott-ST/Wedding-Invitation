# -*- coding: utf-8 -*-
"""从 QQ 邮箱拉取 FormSubmit 回执邮件，汇总生成 嘉宾回执统计.xlsx"""
import imaplib
import email
import os
import re
import sys

from email.header import decode_header
from email.utils import parsedate_to_datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HOST = 'imap.qq.com'
USER = os.environ.get('RSVP_EMAIL', '')
CODE = os.environ.get('RSVP_MAILCODE', '')
OUT = '嘉宾回执统计.xlsx'

if not USER or not CODE:
    print('缺少 RSVP_EMAIL / RSVP_MAILCODE 环境变量（GitHub Secrets 未配置）')
    sys.exit(1)


def decode_subject(msg):
    parts = decode_header(msg.get('Subject', ''))
    out = ''
    for payload, enc in parts:
        if isinstance(payload, bytes):
            out += payload.decode(enc or 'utf-8', 'ignore')
        else:
            out += payload
    return out


def get_body(msg):
    plain, html = '', ''
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        ct = part.get_content_type()
        if ct not in ('text/plain', 'text/html'):
            continue
        payload = part.get_payload(decode=True) or b''
        charset = part.get_content_charset() or 'utf-8'
        if ct == 'text/plain' and not plain:
            plain = payload.decode(charset, 'ignore')
        elif ct == 'text/html' and not html:
            html = payload.decode(charset, 'ignore')
    return plain or re.sub(r'<[^>]+>', ' ', html)


rows = []
mail = imaplib.IMAP4_SSL(HOST, 993)
mail.login(USER, CODE)
mail.select('INBOX', readonly=True)
typ, data = mail.search(None, 'ALL')
for uid in data[0].split():
    typ, msg_data = mail.fetch(uid, '(RFC822)')
    if not msg_data or not msg_data[0]:
        continue
    msg = email.message_from_bytes(msg_data[0][1])
    subject = decode_subject(msg)
    if '婚礼回执' not in subject:
        continue
    text = subject + '\n' + get_body(msg)
    m_name = re.search(r'姓名\s*[:：]?\s*([^\s<，,、；;]+)', text)
    if not m_name:
        continue
    m_cnt = re.search(r'出席人数\s*[:：]?\s*(\d+)', text)
    try:
        when = parsedate_to_datetime(msg.get('Date', '')).strftime('%Y-%m-%d %H:%M')
    except Exception:
        when = msg.get('Date', '')
    rows.append({
        'id': msg.get('Message-ID', ''),
        'time': when,
        'name': m_name.group(1),
        'count': int(m_cnt.group(1)) if m_cnt else None,
    })
mail.logout()

# 按邮件 Message-ID 去重，按提交时间排序
seen, uniq = set(), []
for r in sorted(rows, key=lambda x: x['time']):
    if r['id'] and r['id'] in seen:
        continue
    seen.add(r['id'])
    uniq.append(r)

wb = Workbook()
ws = wb.active
ws.title = '回执统计'

green_fill = PatternFill('solid', fgColor='2F4A3C')
gold_fill = PatternFill('solid', fgColor='F5EFE2')
thin = Side(style='thin', color='D9C48A')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal='center', vertical='center')

headers = ['序号', '姓名', '出席人数', '提交时间']
for col, width in zip('ABCD', (8, 18, 12, 20)):
    ws.column_dimensions[col].width = width
ws.append(headers)
for cell in ws[1]:
    cell.font = Font(bold=True, color='F5EFE2')
    cell.fill = green_fill
    cell.alignment = center
    cell.border = border

total = 0
for i, r in enumerate(uniq, 1):
    cnt = r['count'] if r['count'] is not None else ''
    if r['count']:
        total += r['count']
    ws.append([i, r['name'], cnt, r['time']])

for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    for cell in row:
        cell.alignment = center
        cell.border = border
        if row[0].row % 2 == 1:
            cell.fill = gold_fill

ws.append([])
ws.append(['合计', '%d 份回执' % len(uniq), total, ''])
for cell in ws[ws.max_row]:
    cell.font = Font(bold=True, color='2F4A3C')
    cell.alignment = center
    cell.border = border

wb.save(OUT)
print('回执 %d 份，出席合计 %d 人，已生成 %s' % (len(uniq), total, OUT))
