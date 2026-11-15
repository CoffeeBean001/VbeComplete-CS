# -*- coding: utf-8 -*-
"""把「真实函数清单 + 逐文件讲解文案」编译成逐页课件源文件。
用法：python render.py [--limit N]
"""
import json, io, os, sys, re, math

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
sys.path.insert(0, BASE)
from content_meta import STAGES, FILES, INTRO, OUTRO
from content_notes import NOTES

INV = json.load(open(os.path.join(BASE, 'ppt_inventory.json'), encoding='utf-8'))
TOTAL_METHODS = sum(len(v['methods']) for v in INV.values())
TOTAL_FILES = len(INV)

# ---- 全局视觉约定 ----------------------------------------------------------
P = '#2563EB'        # 主色
PD = '#1E3A8A'       # 深主色
CY = '#06B6D4'       # 辅色
AC = '#F59E0B'       # 强调
BG = '#FFFFFF'
INK = '#1E293B'
MU = '#64748B'
LINE = '#E5E7EB'
LIGHT = '#EFF6FF'
SOFT = '#F8FAFC'
FONT = 'Microsoft YaHei, Arial'
ROWS = 12            # 每页函数行数

# 0 方法 / 少方法的文件：把类与静态成员也列进同一张表，避免空页
EXTRA = {
    'BuiltinTables.g.cs': [
        ('Functions（数组）', '171 条内建函数名字表', 9),
        ('TypeNames（数组）', '内建类型名字表', 43),
        ('Keywords（数组）', 'VBA 语言关键字全集', 51),
        ('KeywordsWantSpace（数组）', '补完之后要跟空格的关键字', 70),
        ('BuiltinTables（静态类）', '由 Python 版导出，两边必须逐字一致', 6),
    ],
    'ComCall.cs': [
        ('ComCall（静态类）', '只为带 ByRef 出参的 VBA 方法而存在', 10),
        ('CallFourOut', '取回 CodePane.GetSelection 的四个出参', 55),
        ('VARIANT（结构体）', '用 FieldOffset 表示 COM 变体类型', 19),
        ('DISPPARAMS（结构体）', '描述参数数组与命名参数个数', 26),
        ('IDispatchRaw（接口）', '直接调 IDispatch::Invoke 的原生签名', 36),
    ],
}
TIPS = [
    '按行号顺序读：先看判定函数，再看动作函数',
    '函数名带 Is / Has / Wants 的多半返回 bool',
    '函数名带 On / Do 的多半是事件或按钮回调',
    '带 ForTest / Probe / Test 的是给自检与探针留的口子',
    '改完这一层的判据，第一件事是跑 --selftest',
    '拿不准某个函数干什么，就去 SelfTest.cs 找同名断言',
    '行号是 src 文件里的真实行号，可以直接跳过去看',
    '返回 string / int 的通常是列号、行号、文本换算工具',
    '这一层不碰 COM，所以能在没有 Office 的机器上验证',
    '先看函数签名猜用途，再回来看这句说明对不对',
    '源文件顶部的注释往往比函数注释更值得先读',
    '一个文件读不下去，先跳到下一站，回头再补',
]


def J(s):
    """JSX 字符串表达式子节点；文案里的尖括号换成安全字符。"""
    s = str(s).replace('<', '‹').replace('>', '›').replace('&', '＆')
    return '{ ' + json.dumps(s, ensure_ascii=False) + ' }'


def clean(s, n):
    s = re.sub(r'\s+', ' ', str(s or '')).strip()
    s = s.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
    s = re.sub(r'</?[A-Za-z][A-Za-z0-9_ .=\"\-]*>', '', s)
    s = s.replace('★', '').replace('◆', '').strip()
    s = re.sub(r'^\s*v\d+[:：]\s*', '', s)
    if len(s) > n:
        s = s[:n - 1] + '…'
    return s


def desc_of(fname, m):
    """一行说明：优先手写白话 → 源码注释 → 按签名兜底。"""
    name = m['name']
    hand = NOTES.get(fname, {}).get(name)
    if hand:
        return clean(hand, 22)
    doc = clean(m.get('doc', ''), 22)
    if doc:
        return doc
    ret = (m.get('ret') or '').lower()
    if ret == 'bool':
        return '判定一个条件是否成立，返回 true / false'
    if ret in ('int', 'long'):
        return '算出行号、列号或计数一类的整数'
    if ret == 'string':
        return '返回处理之后的文本'
    if ret == 'void':
        return '执行一次状态更新或界面动作'
    if ret.startswith('list') or ret.startswith('dict'):
        return '收集并返回一组结果'
    return '返回 ' + (m.get('ret') or '一个') + ' 类型的结果'


# ---- 页面计划 --------------------------------------------------------------
pages = []


def add(kind, **kw):
    kw['kind'] = kind
    pages.append(kw)


add('cover')
add('catalog')
for ci, (num, title, lead, flist) in enumerate(STAGES):
    add('section', ch=ci, num=num, title=title, lead=lead)
    if ci == 0:
        add('roadmap', ch=ci)
        for i in range(len(INTRO)):
            add('intro', ch=ci, idx=i)
    else:
        for fname in flist:
            methods = INV[fname]['methods']
            chunks = [methods[i:i + ROWS] for i in range(0, len(methods), ROWS)] or [[]]
            for k, chunk in enumerate(chunks):
                add('file', ch=ci, file=fname, part=k + 1, parts=len(chunks), methods=chunk)
    if ci == len(STAGES) - 1:
        for i in range(len(OUTRO)):
            add('outro', ch=ci, idx=i)
add('ending')

TOTAL_PAGES = len(pages)
sec_start = {}
for i, pg in enumerate(pages):
    if pg['kind'] == 'section':
        sec_start[pg['ch']] = i + 1

# 全局文件序号
file_no = {}
n = 0
for ci in range(1, len(STAGES)):
    for fname in STAGES[ci][3]:
        n += 1
        file_no[fname] = n


# ---- 通用零件 --------------------------------------------------------------
def slide_open():
    return ("<Slide style={{ width: '1280px', height: '720px', padding: '60px', "
            "background: '" + BG + "', display: 'flex', flexDirection: 'column', "
            "fontFamily: '" + FONT + "' }}>")


def footer(idx):
    return ("<Box style={{ position: 'absolute', bottom: 30, left: 60, right: 60, height: 20, "
            "flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center' }}>"
            "<Text style={{ fontSize: 14, color: '" + MU + "' }}>"
            + J('VbeComplete-CS · 逐文件精读 · 面向零基础读者') + "</Text>"
            "<Text style={{ fontSize: 14, color: '" + MU + "' }}>"
            + J('P%02d / %d' % (idx + 1, TOTAL_PAGES)) + "</Text></Box>")


def header(eyebrow, title):
    return (
        "<Box style={{ height: 76, flexDirection: 'column', gap: 4, justifyContent: 'center' }}>"
        "<Text style={{ fontSize: 16, color: '" + P + "', fontWeight: 'bold', letterSpacing: 1 }}>" + J(eyebrow) + "</Text>"
        "<Text style={{ fontSize: 34, fontWeight: 'bold', color: '" + INK + "', lineHeight: 1.2 }}>" + J(clean(title, 30)) + "</Text>"
        "</Box>"
        "<Box style={{ height: 1, background: '" + LINE + "', marginTop: 6, marginBottom: 12 }} />"
    )


def strip(text, accent=False, height=40):
    return ("<Box style={{ height: " + str(height) + ", background: '" + (LIGHT if not accent else PD) + "', "
            "borderRadius: 8, padding: '0 16px', flexDirection: 'row', alignItems: 'center' }}>"
            "<Text style={{ fontSize: 16, color: '" + (P if not accent else '#FFFFFF') + "', fontWeight: 'bold' }}>"
            + J(clean(text, 62)) + "</Text></Box>")


def info_card(fname, pg, variant, dark=False):
    """文件信息卡：职责 / 为什么这时候读 / 先读哪几个。"""
    meta = FILES[fname]
    v = INV[fname]
    bg = PD if dark else LIGHT
    fg = '#FFFFFF' if dark else INK
    sub = '#DBEAFE' if dark else MU
    emp = '#FFFFFF' if dark else P
    div = 'rgba(255,255,255,0.28)' if dark else '#CBD5E1'
    no = file_no[fname]
    parts = []
    parts.append("<Box style={{ flexDirection: 'row', alignItems: 'baseline', gap: 12 }}>"
                 "<Text style={{ fontSize: 32, fontWeight: 'bold', color: '" + ('#93C5FD' if dark else P) + "' }}>" + J('FILE %02d' % no) + "</Text>"
                 "</Box>")
    parts.append("<Text style={{ fontSize: 21, fontWeight: 'bold', color: '" + fg + "' }}>" + J(fname) + "</Text>")
    parts.append("<Text style={{ fontSize: 14, color: '" + sub + "' }}>" +
                 J('%d 行 · %d 个方法 · 第 %d/%d 页' % (v['lines'], len(v['methods']), pg['part'], pg['parts'])) + "</Text>")
    parts.append("<Box style={{ height: 1, background: '" + div + "' }} />")
    if pg['part'] == 1:
        parts.append("<Text style={{ fontSize: 17, fontWeight: 'bold', color: '" + emp + "' }}>" + J('这个文件干什么') + "</Text>")
        parts.append("<Text style={{ fontSize: 16, color: '" + fg + "', lineHeight: 1.45 }}>" + J(clean(meta['role'], 30)) + "</Text>")
        parts.append("<Text style={{ fontSize: 17, fontWeight: 'bold', color: '" + emp + "' }}>" + J('为什么安排在这儿读') + "</Text>")
        parts.append("<Text style={{ fontSize: 16, color: '" + fg + "', lineHeight: 1.45 }}>" + J(clean(meta['why'], 30)) + "</Text>")
        parts.append("<Text style={{ fontSize: 17, fontWeight: 'bold', color: '" + emp + "' }}>" + J('先读这 3 个') + "</Text>")
        for nm, why in meta['first'][:3]:
            parts.append("<Text style={{ fontSize: 14, color: '" + sub + "', lineHeight: 1.35 }}>" +
                         J(clean(nm + ' · ' + why, 20)) + "</Text>")
    else:
        parts.append("<Text style={{ fontSize: 17, fontWeight: 'bold', color: '" + emp + "' }}>" + J('本页 12 个函数在讲什么') + "</Text>")
        parts.append("<Text style={{ fontSize: 16, color: '" + fg + "', lineHeight: 1.45 }}>" +
                     J(clean('接着上一页：先看 bool / Is 判定函数，再看 void / On 动作函数，最后看文本与数值换算。', 40)) + "</Text>")
        parts.append("<Text style={{ fontSize: 17, fontWeight: 'bold', color: '" + emp + "' }}>" + J('按用途分三组') + "</Text>")
        chunks = pg['methods']
        nb = sum(1 for m in chunks if (m.get('ret') or '').lower() == 'bool')
        nv = sum(1 for m in chunks if (m.get('ret') or '').lower() == 'void')
        no_ = len(chunks) - nb - nv
        for label, cnt in [('条件判定（bool）', nb), ('状态动作（void）', nv), ('文本 / 数值 / 其它', no_)]:
            parts.append("<Text style={{ fontSize: 16, color: '" + fg + "', lineHeight: 1.4 }}>" +
                         J('%s：%d 个' % (label, cnt)) + "</Text>")
        parts.append("<Text style={{ fontSize: 17, fontWeight: 'bold', color: '" + emp + "' }}>" + J('别跳过的一句') + "</Text>")
        parts.append("<Text style={{ fontSize: 16, color: '" + fg + "', lineHeight: 1.45 }}>" + J(clean(meta['takeaway'], 34)) + "</Text>")
    inner = "<Box style={{ width: 300, height: 450, background: '" + bg + "', borderRadius: 8, padding: 14, flexDirection: 'column', gap: 4 }}>" + ''.join(parts) + "</Box>"
    return inner


def func_table(methods, fname, part=1):
    def row(i, name, desc, line, tip=False):
        zebra = SOFT if i % 2 else BG
        col = MU if tip else PD
        return ("<Box style={{ height: 32, flexDirection: 'row', alignItems: 'center', gap: 10, "
                "background: '" + zebra + "', padding: '0 8px' }}>"
                "<Text style={{ width: 248, fontSize: 16, fontWeight: 'bold', color: '" + col + "' }}>" + J(name) + "</Text>"
                "<Text style={{ flex: 1, fontSize: 16, color: '" + (MU if tip else INK) + "' }}>" + J(clean(desc, 22)) + "</Text>"
                "<Text style={{ width: 50, fontSize: 14, color: '" + MU + "', textAlign: 'right' }}>" + J(line) + "</Text>"
                "</Box>")
    rows = ["<Box style={{ height: 28, flexDirection: 'row', alignItems: 'center', gap: 10, "
            "background: '" + LIGHT + "', borderRadius: 4, padding: '0 8px' }}>"
            "<Text style={{ width: 248, fontSize: 15, fontWeight: 'bold', color: '" + PD + "' }}>" + J('函数 / 成员') + "</Text>"
            "<Text style={{ flex: 1, fontSize: 15, fontWeight: 'bold', color: '" + PD + "' }}>" + J('这一行在讲什么') + "</Text>"
            "<Text style={{ width: 50, fontSize: 14, color: '" + MU + "', textAlign: 'right' }}>" + J('行号') + "</Text>"
            "</Box>"]
    items = [(m['name'], desc_of(fname, m), 'L' + str(m['line']), False) for m in methods]
    if part == 1:
        for name, desc, ln in EXTRA.get(fname, []):
            items.append((name, desc, 'L' + str(ln), False))
    k = 0
    while len(items) < ROWS:
        items.append(('读法提示', TIPS[k % len(TIPS)], '—', True))
        k += 1
    for i, (name, desc, line, tip) in enumerate(items[:ROWS]):
        rows.append(row(i, name, desc, line, tip))
    return ("<Box style={{ flex: 1, height: 450, border: '1px solid " + LINE + "', borderRadius: 8, "
            "padding: 14, background: '" + BG + "', flexDirection: 'column' }}>" + ''.join(rows) + "</Box>")


def file_page(idx, pg):
    fname = pg['file']
    meta = FILES[fname]
    stage = STAGES[pg['ch']]
    variant = idx % 3
    eyebrow = '第 %d 站 · %s   /   FILE %02d / %d   /   %s' % (
        pg['ch'], stage[1].split('：')[1] if '：' in stage[1] else stage[1],
        file_no[fname], TOTAL_FILES, ('%d/%d' % (pg['part'], pg['parts'])) if pg['parts'] > 1 else '全览')
    title = fname + ' · ' + clean(meta['role'], 18)
    card = info_card(fname, pg, variant, dark=(variant == 1))
    table = func_table(pg['methods'], fname, pg['part'])
    if variant == 1:      # 深色卡在右
        body = table + card
    elif variant == 2:    # 双卡在右（上下叠），表格在左
        body = table + card
    else:                 # 浅卡在左
        body = card + table
    return (slide_open() + header(eyebrow, title) +
            "<Box style={{ height: 450, flexDirection: 'row', gap: 24 }}>" + body + "</Box>" +
            strip('带走：' + meta['takeaway']) + footer(idx) + "</Slide>")


# ---- 结构页 ----------------------------------------------------------------
def cover():
    return (slide_open().replace("padding: '60px'", "padding: '80px'") +
            "<Box style={{ height: 560, flexDirection: 'column', justifyContent: 'center', gap: 26 }}>"
            "<Text style={{ fontSize: 18, color: '" + P + "', letterSpacing: 3, fontWeight: 'bold' }}>" + J('从入口到探针 · 新手学习版') + "</Text>"
            "<Text style={{ fontSize: 76, fontWeight: 'bold', color: '" + INK + "', lineHeight: 1.15 }}>" + J('VbeComplete-CS') + "</Text>"
            "<Text style={{ fontSize: 52, fontWeight: 'bold', color: '" + P + "', lineHeight: 1.2 }}>" + J('逐文件精读') + "</Text>"
            "<Box style={{ height: 1, background: '" + LINE + "', width: 900 }} />"
            "<Text style={{ fontSize: 22, color: '" + MU + "', lineHeight: 1.5 }}>" +
            J('按阅读顺序，把每个源文件、每个函数讲一遍；先读哪个、后读哪个，都有理由。') + "</Text>"
            "<Text style={{ fontSize: 16, color: '" + MU + "' }}>" + J('VbeComplete-CS 项目 · 2026-11-07') + "</Text>"
            "</Box></Slide>")


def catalog():
    items = []
    for ci, (num, title, lead, flist) in enumerate(STAGES):
        items.append("<Box style={{ height: 62, flexDirection: 'row', alignItems: 'center', gap: 20, "
                     "borderBottom: '1px solid " + LINE + "' }}>"
                     "<Text style={{ width: 70, fontSize: 32, fontWeight: 'bold', color: '" + P + "' }}>" + J(num) + "</Text>"
                     "<Text style={{ flex: 1, fontSize: 24, fontWeight: 'bold', color: '" + INK + "' }}>" + J(title) + "</Text>"
                     "<Text style={{ width: 90, fontSize: 16, color: '" + MU + "', textAlign: 'right' }}>" +
                     J('P%02d 起' % sec_start[ci]) + "</Text></Box>")
    return (slide_open() +
            "<Box style={{ flexDirection: 'row', gap: 40, height: 600 }}>"
            "<Box style={{ width: 260, flexDirection: 'column', justifyContent: 'center', gap: 14 }}>"
            "<Text style={{ fontSize: 56, fontWeight: 'bold', color: '" + INK + "' }}>" + J('目录') + "</Text>"
            "<Text style={{ fontSize: 20, color: '" + P + "', letterSpacing: 2 }}>" + J('CONTENTS') + "</Text>"
            "<Box style={{ height: 1, background: '" + LINE + "' }} />"
            "<Text style={{ fontSize: 16, color: '" + MU + "', lineHeight: 1.5 }}>" +
            J('六章按「先看哪个文件」排序，跟着走一遍就能把项目串起来。') + "</Text>"
            "</Box>"
            "<Box style={{ flex: 1, flexDirection: 'column', justifyContent: 'center' }}>" + ''.join(items) + "</Box>"
            "</Box></Slide>")


def section_page(idx, pg):
    num, title, lead, flist = STAGES[pg['ch']]
    return ("<Slide style={{ width: '1280px', height: '720px', padding: '80px', background: '" + PD + "', "
            "display: 'flex', flexDirection: 'column', justifyContent: 'center', gap: 18, fontFamily: '" + FONT + "' }}>"
            "<Text style={{ fontSize: 150, fontWeight: 'bold', color: 'rgba(255,255,255,0.18)', lineHeight: 1 }}>" + J(num) + "</Text>"
            "<Text style={{ fontSize: 52, fontWeight: 'bold', color: '#FFFFFF', lineHeight: 1.2 }}>" + J(title) + "</Text>"
            "<Box style={{ height: 1, background: 'rgba(255,255,255,0.35)', width: 900 }} />"
            "<Text style={{ fontSize: 20, color: '#DBEAFE', lineHeight: 1.5 }}>" + J(lead) + "</Text>"
            "</Slide>")


def roadmap(idx):
    stops = [('01', 'Program.cs', '入口'), ('02', 'Rules / Editing', '纯判据'),
             ('03', 'Sqlite / 词库', '数据'), ('04', 'VbeBackend / App', '运行时'),
             ('05', 'SelfTest / Probes', '验证')]
    nodes = []
    for i, (n_, f, label) in enumerate(stops):
        nodes.append("<Box style={{ flex: 1, height: 190, background: '" + (LIGHT if i % 2 == 0 else SOFT) + "', "
                     "borderRadius: 8, padding: 16, flexDirection: 'column', gap: 10, justifyContent: 'center' }}>"
                     "<Text style={{ fontSize: 34, fontWeight: 'bold', color: '" + P + "' }}>" + J(n_) + "</Text>"
                     "<Text style={{ fontSize: 19, fontWeight: 'bold', color: '" + INK + "' }}>" + J(f) + "</Text>"
                     "<Text style={{ fontSize: 16, color: '" + MU + "' }}>" + J(label) + "</Text></Box>")
        if i < len(stops) - 1:
            nodes.append("<Box style={{ width: 26, height: 190, justifyContent: 'center', alignItems: 'center' }}>"
                         "<Text style={{ fontSize: 26, color: '" + AC + "', fontWeight: 'bold' }}>" + J('→') + "</Text></Box>")
    return (slide_open() + header('读之前 · 全局路线', '先看哪个文件，再看哪个文件：一条 5 站路线') +
            "<Box style={{ height: 450, flexDirection: 'column', gap: 22 }}>"
            "<Box style={{ height: 190, flexDirection: 'row', alignItems: 'center' }}>" + ''.join(nodes) + "</Box>"
            "<Box style={{ flexDirection: 'row', gap: 22, flex: 1 }}>"
            "<Box style={{ flex: 1, background: '" + SOFT + "', borderRadius: 8, padding: 20, flexDirection: 'column', gap: 12 }}>"
            "<Text style={{ fontSize: 22, fontWeight: 'bold', color: '" + INK + "' }}>" + J('为什么是这个顺序') + "</Text>"
            "<Text style={{ fontSize: 18, color: '" + INK + "', lineHeight: 1.5 }}>" +
            J('先读不依赖 Windows、不依赖 Office 的纯字符串判据，它们最好懂也最好测；'
              '等你知道「该不该弹」之后，再去读真正去碰 VBE 的那层。') + "</Text></Box>"
            "<Box style={{ width: 330, background: '" + PD + "', borderRadius: 8, padding: 20, flexDirection: 'column', gap: 12 }}>"
            "<Text style={{ fontSize: 22, fontWeight: 'bold', color: '#FFFFFF' }}>" + J('一句话记住') + "</Text>"
            "<Text style={{ fontSize: 18, color: '#DBEAFE', lineHeight: 1.5 }}>" +
            J('钩子线程快但绝不碰 COM；UI 线程慢但能读真代码。后面一半的怪设计都是这句的代价。') + "</Text></Box>"
            "</Box></Box>" + strip('每一站的每一页，都会列出这个文件的全部函数、一句说明和源码行号。') +
            footer(idx) + "</Slide>")


def intro_page(idx, pg):
    i = pg['idx']
    title, text = INTRO[i]
    eyebrow = '读之前 · 先建立地图   /   %d / 3' % (i + 1)
    if i == 0:
        body = ("<Box style={{ height: 450, flexDirection: 'row', gap: 24 }}>"
                "<Box style={{ flex: 1, background: '" + SOFT + "', borderRadius: 8, padding: 26, flexDirection: 'column', gap: 16, justifyContent: 'center' }}>"
                "<Text style={{ fontSize: 24, fontWeight: 'bold', color: '" + INK + "' }}>" + J('VBE 自带的列表') + "</Text>"
                "<Text style={{ fontSize: 18, color: '" + INK + "', lineHeight: 1.6 }}>" +
                J('只认「已经存在的成员」：对象、方法、枚举常量。你刚敲完 Dim 变量名，'
                  '或自己写了一个函数，它并不会因为你在打字就把它列出来。') + "</Text></Box>"
                "<Box style={{ width: 60, justifyContent: 'center', alignItems: 'center' }}>"
                "<Text style={{ fontSize: 30, color: '" + AC + "', fontWeight: 'bold' }}>" + J('VS') + "</Text></Box>"
                "<Box style={{ flex: 1, background: '" + LIGHT + "', borderRadius: 8, padding: 26, flexDirection: 'column', gap: 16, justifyContent: 'center' }}>"
                "<Text style={{ fontSize: 24, fontWeight: 'bold', color: '" + PD + "' }}>" + J('这个工具补的那一半') + "</Text>"
                "<Text style={{ fontSize: 18, color: '" + INK + "', lineHeight: 1.6 }}>" +
                J('后台读你的模块全文，解析出自己声明的名字、过程、窗体控件，'
                  '再在光标下方画一个不抢焦点的候选窗。') + "</Text></Box></Box>")
    elif i == 1:
        tech = [('语言 / 运行时', 'C# · .NET 10 · WinExe', '不带控制台，托盘常驻'),
                ('界面', 'WinForms（全部自绘）', '候选窗不抢焦点，只能自己画'),
                ('Office 接入', 'dynamic 后期绑定 + COM', '不绑 Excel / WPS 版本，一个 exe 通吃'),
                ('数据库', 'P/Invoke 系统 winsqlite3.dll', '零 NuGet 依赖，双击就能跑'),
                ('发布', 'PublishSingleFile → dist\\VbeComplete.exe', '资源必须内嵌，旁边没有散文件')]
        rows = []
        for k, (a, b, c) in enumerate(tech):
            rows.append("<Box style={{ height: 56, flexDirection: 'row', alignItems: 'center', gap: 16, "
                        "background: '" + (SOFT if k % 2 else BG) + "', padding: '0 14px' }}>"
                        "<Text style={{ width: 200, fontSize: 18, fontWeight: 'bold', color: '" + PD + "' }}>" + J(a) + "</Text>"
                        "<Text style={{ width: 380, fontSize: 17, color: '" + INK + "' }}>" + J(b) + "</Text>"
                        "<Text style={{ flex: 1, fontSize: 16, color: '" + MU + "' }}>" + J(c) + "</Text></Box>")
        body = ("<Box style={{ height: 450, flexDirection: 'row', gap: 24 }}>"
                "<Box style={{ width: 210, background: '" + PD + "', borderRadius: 8, padding: 20, "
                "flexDirection: 'column', justifyContent: 'center', gap: 10 }}>"
                "<Text style={{ fontSize: 96, fontWeight: 'bold', color: '#FFFFFF', lineHeight: 1 }}>" + J('0') + "</Text>"
                "<Text style={{ fontSize: 20, color: '#DBEAFE', lineHeight: 1.4 }}>" + J('个第三方 NuGet 包') + "</Text>"
                "<Text style={{ fontSize: 16, color: '#93C5FD', lineHeight: 1.4 }}>" +
                J('SQLite 和 COM 都直接用系统自带能力') + "</Text></Box>"
                "<Box style={{ flex: 1, border: '1px solid " + LINE + "', borderRadius: 8, padding: 14, flexDirection: 'column' }}>" +
                ''.join(rows) + "</Box></Box>")
    else:
        layers = [('纯判据层', 'Rules / Editing / Parser / Scope', LIGHT),
                  ('引擎层', 'Engine（攒池子、匹配、排序）', SOFT),
                  ('宿主层', 'VbeBackend（唯一大面积碰 COM）', LIGHT),
                  ('主流程层', 'App（钩子 + 轮询 + 按键路由）', SOFT),
                  ('界面层', 'PopupForm / SettingsForm / SnippetForm', LIGHT)]
        rows = []
        for k, (a, b, bg) in enumerate(layers):
            rows.append("<Box style={{ height: 60, flexDirection: 'row', alignItems: 'center', gap: 18, "
                        "background: '" + bg + "', borderRadius: 6, padding: '0 16px' }}>"
                        "<Text style={{ width: 150, fontSize: 19, fontWeight: 'bold', color: '" + PD + "' }}>" + J(a) + "</Text>"
                        "<Text style={{ flex: 1, fontSize: 17, color: '" + INK + "' }}>" + J(b) + "</Text></Box>")
        body = ("<Box style={{ height: 450, flexDirection: 'row', gap: 24 }}>"
                "<Box style={{ flex: 1, flexDirection: 'column', gap: 12, padding: 8 }}>" + ''.join(rows) + "</Box>"
                "<Box style={{ width: 330, flexDirection: 'column', gap: 16 }}>"
                "<Box style={{ flex: 1, background: '" + PD + "', borderRadius: 8, padding: 18, flexDirection: 'column', gap: 10 }}>"
                "<Text style={{ fontSize: 21, fontWeight: 'bold', color: '#FFFFFF' }}>" + J('钩子线程') + "</Text>"
                "<Text style={{ fontSize: 16, color: '#DBEAFE', lineHeight: 1.5 }}>" +
                J('快；只读内存状态。绝不碰 COM、绝不 Sleep——它卡一下，用户整个键盘就卡住。') + "</Text></Box>"
                "<Box style={{ flex: 1, background: '" + LIGHT + "', borderRadius: 8, padding: 18, flexDirection: 'column', gap: 10 }}>"
                "<Text style={{ fontSize: 21, fontWeight: 'bold', color: '" + PD + "' }}>" + J('UI 线程') + "</Text>"
                "<Text style={{ fontSize: 16, color: '" + INK + "', lineHeight: 1.5 }}>" +
                J('慢；能读真 VBE 代码、能画窗口。两条线程靠一份内存快照接力。') + "</Text></Box>"
                "</Box></Box>")
    return (slide_open() + header(eyebrow, title) + body +
            strip('本页结论：' + text, accent=(i == 2)) + footer(idx) + "</Slide>")


def outro_page(idx, pg):
    i = pg['idx']
    title, text = OUTRO[i]
    eyebrow = '第 5 站 · 收尾   /   %d / 3' % (i + 1)
    if i == 0:
        rows = [('第一层', '587 条纯逻辑自检 · --selftest', '不依赖桌面和 Office，最快，判据的标准答案'),
                ('第二层', '合成后端探针 · --enginetest', '假模块跑真 Engine，A/B 开关能确定性复现 bug'),
                ('第三层', '真机探针 · --declprobe / --fmtprobe', '真 VBE、真钩子、真按 Ctrl+S，最贵也最可信')]
        body = "<Box style={{ height: 450, flexDirection: 'column', gap: 18, justifyContent: 'center' }}>"
        for k, (a, b, c) in enumerate(rows):
            w = 1050 - k * 160
            body += ("<Box style={{ width: " + str(w) + ", height: 112, marginLeft: " + str((1050 - w) // 2) + ", "
                     "background: '" + (PD if k == 2 else LIGHT) + "', borderRadius: 8, padding: 16, "
                     "flexDirection: 'row', alignItems: 'center', gap: 20 }}>"
                     "<Text style={{ width: 110, fontSize: 22, fontWeight: 'bold', color: '" + ('#FFFFFF' if k == 2 else P) + "' }}>" + J(a) + "</Text>"
                     "<Text style={{ width: 340, fontSize: 18, fontWeight: 'bold', color: '" + ('#FFFFFF' if k == 2 else INK) + "' }}>" + J(b) + "</Text>"
                     "<Text style={{ flex: 1, fontSize: 16, color: '" + ('#DBEAFE' if k == 2 else MU) + "', lineHeight: 1.4 }}>" + J(c) + "</Text>"
                     "</Box>")
        body += "</Box>"
    elif i == 1:
        cmds = [('编译', 'dotnet build src/VbeComplete.csproj -c Release'),
                ('纯逻辑', 'dist\\VbeComplete.exe --selftest'),
                ('合成探针', 'dist\\VbeComplete.exe --enginetest'),
                ('真机探针', 'dist\\VbeComplete.exe --declprobe / --fmtprobe / --hdrprobe'),
                ('发布', 'dotnet publish ... -r win-x64 -p:PublishSingleFile=true -o dist'),
                ('一键全跑', 'check-cs.bat（12 步，改 UI 加 --uitest）')]
        rows = []
        for k, (a, b) in enumerate(cmds):
            rows.append("<Box style={{ height: 56, flexDirection: 'row', alignItems: 'center', gap: 18, "
                        "background: '" + (SOFT if k % 2 else BG) + "', padding: '0 14px' }}>"
                        "<Text style={{ width: 130, fontSize: 18, fontWeight: 'bold', color: '" + PD + "' }}>" + J(a) + "</Text>"
                        "<Text style={{ flex: 1, fontSize: 17, color: '" + INK + "', fontFamily: 'Consolas, Arial' }}>" + J(b) + "</Text></Box>")
        body = ("<Box style={{ height: 450, flexDirection: 'column', gap: 16 }}>"
                "<Box style={{ flex: 1, border: '1px solid " + LINE + "', borderRadius: 8, padding: 14, flexDirection: 'column' }}>" +
                ''.join(rows) + "</Box>"
                "<Box style={{ height: 64, background: '" + PD + "', borderRadius: 8, padding: 16, flexDirection: 'column', gap: 6 }}>"
                "<Text style={{ fontSize: 18, fontWeight: 'bold', color: '#FFFFFF' }}>" + J('A/B 开关（Git Bash 写法）') + "</Text>"
                "<Text style={{ fontSize: 16, color: '#DBEAFE' }}>" +
                J('VBECOMPLETE_NO_PROC_HEADER_FIX=1 dist/VbeComplete.exe --hdrprobe  →  关掉修复，必须变红') + "</Text>"
                "</Box></Box>")
    else:
        steps = [('1', '跑 --selftest', '先看到 587 / 0'),
                 ('2', '读 Rules.WordBefore', '只有 21 行'),
                 ('3', '加一条断言', '在 SelfTest.cs 里跑通'),
                 ('4', '读 Fuzzy.GreedyPositions', '理解模糊命中'),
                 ('5', '跑 --enginetest', '看合成后端用例'),
                 ('6', '用 A/B 复现 bug', '亲手把断言变红'),
                 ('7', '再进 App.cs', '此时三条线都懂了')]
        cards = []
        for k, (n_, a, b) in enumerate(steps):
            cards.append("<Box style={{ flex: 1, minWidth: 0, height: 180, background: '" + (LIGHT if k % 2 == 0 else SOFT) + "', "
                         "borderRadius: 8, padding: 14, flexDirection: 'column', gap: 12 }}>"
                         "<Text style={{ fontSize: 40, fontWeight: 'bold', color: '" + P + "' }}>" + J(n_) + "</Text>"
                         "<Text style={{ width: '100%', fontSize: 17, fontWeight: 'bold', color: '" + INK + "', lineHeight: 1.35 }}>" + J(a) + "</Text>"
                         "<Text style={{ width: '100%', fontSize: 15, color: '" + MU + "', lineHeight: 1.35 }}>" + J(b) + "</Text></Box>")
        body = ("<Box style={{ height: 450, flexDirection: 'column', gap: 24, justifyContent: 'center' }}>"
                "<Box style={{ width: '100%', flexDirection: 'row', gap: 10 }}>" + ''.join(cards) + "</Box>"
                "<Box style={{ width: '100%', background: '" + SOFT + "', borderRadius: 8, padding: 20, flexDirection: 'column', gap: 10 }}>"
                "<Text style={{ fontSize: 20, fontWeight: 'bold', color: '" + INK + "' }}>" + J('别从 App.cs 开始') + "</Text>"
                "<Text style={{ fontSize: 17, color: '" + INK + "', lineHeight: 1.5 }}>" +
                J('它有 1472 行，同时用了钩子、COM、轮询和判据四层知识。'
                  '按上面的顺序走，每一段都有自检断言当标准答案，读不懂就去看那条断言。') + "</Text></Box></Box>")
    return (slide_open() + header(eyebrow, title) + body +
            strip('带走：' + clean(text, 60), accent=(i == 2)) + footer(idx) + "</Slide>")


def ending(idx):
    return ("<Slide style={{ width: '1280px', height: '720px', padding: '80px', background: '" + PD + "', "
            "display: 'flex', flexDirection: 'column', justifyContent: 'center', gap: 24, fontFamily: '" + FONT + "' }}>"
            "<Text style={{ fontSize: 22, color: '#93C5FD', letterSpacing: 2 }}>" + J('VbeComplete-CS · 逐文件精读') + "</Text>"
            "<Text style={{ fontSize: 56, fontWeight: 'bold', color: '#FFFFFF', lineHeight: 1.3 }}>" +
            J('先懂判据，再懂宿主；') + "</Text>"
            "<Text style={{ fontSize: 56, fontWeight: 'bold', color: '#FFFFFF', lineHeight: 1.3 }}>" +
            J('先看断言，再看实现。') + "</Text>"
            "<Box style={{ height: 1, background: 'rgba(255,255,255,0.35)', width: 900 }} />"
            "<Text style={{ fontSize: 20, color: '#DBEAFE', lineHeight: 1.6 }}>" +
            J('42 个源文件、590 个方法，全部按阅读顺序过了一遍。'
              '配套文档：项目学习指南.md；一键回归：check-cs.bat。') + "</Text>"
            "</Slide>")


def build(idx, pg):
    k = pg['kind']
    if k == 'cover': return cover()
    if k == 'catalog': return catalog()
    if k == 'section': return section_page(idx, pg)
    if k == 'roadmap': return roadmap(idx)
    if k == 'intro': return intro_page(idx, pg)
    if k == 'file': return file_page(idx, pg)
    if k == 'outro': return outro_page(idx, pg)
    if k == 'ending': return ending(idx)
    raise ValueError(k)


# ---- 输出 ------------------------------------------------------------------
lim = None
if '--limit' in sys.argv:
    lim = int(sys.argv[sys.argv.index('--limit') + 1])
os.makedirs(os.path.join(BASE, 'slides'), exist_ok=True)
for i, pg in enumerate(pages):
    if lim and i + 1 > lim:
        break
    dsl = build(i, pg)
    io.open(os.path.join(BASE, 'slides', '%02d.slide' % (i + 1)), 'w', encoding='utf-8').write(dsl)

# DESIGN.md：全局约定 + 逐页字段表
lines = [
 '# 设计说明（逐文件精读课件）', '',
 '## 全局视觉约定', '',
 '- 画布 1280×720；padding 上下左右各 60px；内容区 1160×600。',
 '- 配色：主色 #2563EB / 深主色 #1E3A8A / 辅色 #06B6D4 / 强调 #F59E0B / 底色 #FFFFFF / 正文 #1E293B。',
 '- 字体：Microsoft YaHei（中文）+ Arial（西文）；代码命令用 Consolas。',
 '- 字号：封面 76 / 章节 52 / 页面主标 34 / 卡片小标 17-22 / 正文与表格 16 / 页脚 14。',
 '- 视觉母题：标题下 1px 贯通细线；固定页脚（章节页/封面/目录/封底无页脚）。',
 '- 素材：全部为结构化图形（Box / SVG 式布局），无外部图片；教学查读型，不使用装饰性配图。',
 '- 版式循环：左浅色信息卡 + 右函数表 / 右深色信息卡 + 左函数表 / 双色卡 + 左函数表，逐页交替。',
 '', '## 逐页字段表', '',
 '| 页 | 类型 | 角色 | 节奏 | 版式 | 内容 | 视觉 | 禁止项 |',
 '|---|---|---|---|---|---|---|---|',
]
for i, pg in enumerate(pages):
    k = pg['kind']
    if k == 'file':
        t = '%s（%d/%d）' % (pg['file'], pg['part'], pg['parts'])
        lay = ['左信息卡+右函数表', '右深卡+左函数表', '双色卡+左函数表'][i % 3]
        vis = 'L3 母版眉标 + 函数表（12 行）'
        anti = '禁止裸表、禁止缩字号塞内容、禁止函数说明留空'
    elif k == 'section':
        t, lay, vis, anti = pg['title'], '章节大字（深底）', '巨型章节编号', '禁止加页脚 / 禁止章节内看点列表'
    elif k == 'cover':
        t, lay, vis, anti = '封面', '大字报式', '超大主标题', '禁止数据指标 / 禁止页脚'
    elif k == 'catalog':
        t, lay, vis, anti = '目录', '左右不对称杂志式', '大字目录 + 六章清单', '禁止统计块 / 禁止阅读顺序复述条'
    elif k == 'roadmap':
        t, lay, vis, anti = '5 站阅读路线', '流程链', '五节点流程 + 结论条', '禁止等宽卡片横排'
    elif k == 'intro':
        t, lay, vis, anti = INTRO[pg['idx']][0], ['对比式', '巨型数字+清单', '分层图+双线程卡'][pg['idx']], '结构化图形', '禁止抽象标题 / 禁止空容器'
    elif k == 'outro':
        t, lay, vis, anti = OUTRO[pg['idx']][0], ['金字塔', '命令清单', '七步流程'][pg['idx']], '结构化图形', '禁止把结尾内容塞进内容页'
    else:
        t, lay, vis, anti = '收束金句', '深底金句 + 落款', '大字两行', '禁止只写「谢谢」'
    role = 'hero' if k in ('cover', 'section', 'roadmap', 'ending') else 'supporting'
    rhythm = 'peak' if role == 'hero' else 'valley'
    lines.append('| P%02d | %s | %s | %s | %s | %s | %s | %s |' %
                 (i + 1, k, role, rhythm, lay, t, vis, anti))
io.open(os.path.join(BASE, 'DESIGN.md'), 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
json.dump({'total_pages': TOTAL_PAGES, 'pages': [
    {'n': i + 1, 'kind': p['kind'], 'file': p.get('file'), 'part': p.get('part'), 'parts': p.get('parts')}
    for i, p in enumerate(pages)]}, open(os.path.join(BASE, 'page_plan.json'), 'w', encoding='utf-8'),
    ensure_ascii=False, indent=1)
print('pages=%d files=%d methods=%d written=%d' %
      (TOTAL_PAGES, TOTAL_FILES, TOTAL_METHODS, lim or TOTAL_PAGES))
