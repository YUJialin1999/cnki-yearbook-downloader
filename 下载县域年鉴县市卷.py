# -*- coding: utf-8 -*-
"""
下载《中国县域统计年鉴》县市卷各年度 Excel 统计表格。

保存目录：<脚本目录>/<年鉴年份>（如 2024、2025）。
文件命名：条目题名-页码.xls（页码取自目录接口的 ysym 字段），
          如 “北京市-3.xls”。

运行：
    D:\\Python\\python.exe 下载县域年鉴县市卷.py          # 下载 TARGET_YEARS
    D:\\Python\\python.exe 下载县域年鉴县市卷.py 2024    # 只下载 2024

说明：校园网/机构 IP 可直接访问知网，无需登录；每次启动干净的临时浏览器，
     不保存配置目录。若个别条目触发拼图验证，完成后重跑脚本即可。
"""

import json
import os
import re
import sys
import time

from playwright.sync_api import sync_playwright

# ========================== 配置区 ==========================

# 任一卷册的已知代码，只用于拉取“全部年度+卷别”列表
ROOT_CODE = 'N2024030156'

# 目录接口需要的拼音码
PINYIN_CODE = 'YZGCA'

# 目标卷册（平台名称形如 “2024(县市卷)”）
TARGET_YEARS = ['2024', '2025']

# 下载节奏（秒）
INTERVAL = 3

# 保存根目录（脚本所在目录）
SAVE_ROOT = os.path.dirname(os.path.abspath(__file__))

API_BASE = 'https://data.cnki.net/api/csyd'
HOME = 'https://data.cnki.net/yearBook/single'
UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

# ===========================================================


def dataclean(name):
    """清理文件名中的换行与 Windows 非法字符。"""
    name = re.sub(r'[\r\n\t]+', ' ', str(name))
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name


def sniff_ext(body):
    if body[:4] == b'\xd0\xcf\x11\xe0':
        return '.xls'
    if body[:2] == b'PK':
        return '.xlsx'
    return None


def already_exists(out_dir, file_code, new_name):
    """断点跳过：目录里已有同条目代码的文件，或同名新文件。"""
    try:
        for n in os.listdir(out_dir):
            if n.lower().endswith(('.xls', '.xlsx')) and (
                    file_code in n or os.path.splitext(n)[0] == new_name):
                return True
    except OSError:
        pass
    return False


class Yb:
    def __init__(self, ctx):
        self.ctx = ctx

    def post(self, path, payload):
        r = self.ctx.request.post(
            API_BASE + path, data=json.dumps(payload),
            headers={'Content-Type': 'application/json',
                     'User-Agent': UA, 'Referer': HOME}, timeout=60000)
        return json.loads(r.text())

    def year_codes(self):
        r = self.post('/StatisticalData/GetYbYearAndCode', {'code': ROOT_CODE})
        mapping = {}
        for it in (r.get('data') or {}).get('list') or []:
            if it.get('year') and it.get('code'):
                mapping[it['year']] = it['code']
        return mapping

    def chapters(self, encrypted_code):
        r = self.post('/StatisticalData/GetContentByCode',
                      {'code': encrypted_code, 'pinyinCode': PINYIN_CODE})
        items = (r['data'] or {}).get('data') or []
        plain = items[0]['code']
        return plain, [(it['mlbm'], it['title']) for it in items]

    def entries(self, plain_code, par_node):
        out, cur = [], 1
        while True:
            r = self.post('/StatisticalData/GetEntResByCodeAndParNode',
                          {'code': plain_code, 'parNode': par_node,
                           'current': cur, 'size': 100})
            d = (r.get('data') or {}).get('data') or {}
            lst = d.get('list') or []
            out.extend(lst)
            if cur * 100 >= (d.get('total') or 0) or not lst:
                return out
            cur += 1
            time.sleep(1)

    def download(self, file_code, dest_path):
        """返回 'ok' 或失败原因字符串。"""
        r = self.post('/StatisticalData/PdfAndCajDownload',
                      {'fileName': file_code, 'downLoadFlag': '2'})
        url = r.get('data')
        if not url or not str(url).startswith('http'):
            return '获取下载链接失败: %s' % (r.get('message') or url)

        resp = self.ctx.request.get(
            url, headers={'User-Agent': UA, 'Referer': 'https://data.cnki.net/'},
            timeout=180000, fail_on_status_code=False)
        body = resp.body() if callable(resp.body) else resp.body
        ctype = resp.headers.get('content-type', '')

        if resp.status == 200 and ('excel' in ctype or sniff_ext(body)):
            ext = sniff_ext(body) or '.xls'
            path = os.path.splitext(dest_path)[0] + ext
            with open(path, 'wb') as f:
                f.write(body)
            return 'ok'

        head = body[:3000].decode('utf-8', errors='ignore')
        if 'login.cnki.net' in resp.url or 'login.cnki.net' in head:
            return '未登录：请在浏览器窗口登录知网后重跑脚本'
        if '/bar/verify/' in resp.url or '安全验证' in head or '拼图' in head:
            return '需要拼图验证：请在浏览器窗口完成验证后重跑脚本'
        plain = re.sub(r'<[^>]+>', ' ', head)
        plain = re.sub(r'\s+', ' ', plain).strip()
        return '下载异常 status=%s: %s' % (resp.status, plain[:120])


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    years = args or TARGET_YEARS

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            channel='chrome', headless=False,
            args=['--disable-blink-features=AutomationControlled'])
        ctx = browser.new_context(accept_downloads=True)
        yb = Yb(ctx)
        try:
            page = ctx.new_page()
            page.goto(HOME, timeout=60000, wait_until='domcontentloaded')

            codes = yb.year_codes()
            print('平台收录: %s' % ', '.join(sorted(codes, reverse=True)))

            for y in years:
                label = '%s(县市卷)' % y
                if label not in codes:
                    print('[!] 平台上没有 %s，跳过' % label)
                    continue
                plain, ch_list = yb.chapters(codes[label])
                print('\n============ %s (代码 %s) ============' % (label, plain))

                tables = []
                for mlbm, ch_title in ch_list:
                    es = [e for e in yb.entries(plain, mlbm)
                          if str(e.get('czexcel')) == '1']
                    print('  [%s] %s : %d 张' % (mlbm, ch_title, len(es)))
                    for e in es:
                        tables.append((e['title'], e.get('ysym'), e['fileCode']))
                    time.sleep(0.5)
                print('  可下载 Excel 共 %d 张' % len(tables))

                out_dir = os.path.join(SAVE_ROOT, y)
                os.makedirs(out_dir, exist_ok=True)
                ok = skip = fail = 0
                for i, (title, page_no, fc) in enumerate(tables, 1):
                    base = dataclean('%s-%s' % (title, page_no))
                    if already_exists(out_dir, fc, base):
                        skip += 1
                        continue
                    dest = os.path.join(out_dir, base)
                    print('  (%d/%d) %s-%s [%s]' % (i, len(tables), title, page_no, fc),
                          flush=True)
                    st = yb.download(fc, dest)
                    if st == 'ok':
                        ok += 1
                    else:
                        fail += 1
                        print('    [X] %s' % st)
                    time.sleep(INTERVAL)
                print('  %s 完成：成功 %d，跳过 %d，失败 %d' % (label, ok, skip, fail))
        finally:
            try:
                ctx.close()
                browser.close()
            except Exception:
                pass


if __name__ == '__main__':
    main()
