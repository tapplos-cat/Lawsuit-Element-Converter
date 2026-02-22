#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
要素式起诉状转换工具 v7.0 (GUI版)
传统起诉状 → 要素式起诉状 (Word docx 输出，精确还原范本格式)

双击运行即可，自动安装依赖、弹出图形界面。
"""

import sys, os, subprocess, re, json, threading

# ============================================================
# 自动安装依赖
# ============================================================
def ensure_dependencies():
    required = {"python-docx": "docx"}
    missing = []
    for pkg_name, import_name in required.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pkg_name)
    if missing:
        print(f"正在安装依赖: {', '.join(missing)} ...")
        for pkg in missing:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
        print("依赖安装完成。")

ensure_dependencies()

from docx import Document as DocxDocument
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ============================================================
# 案由类型
# ============================================================
CASE_TYPES_ORDERED = [
    ("民事纠纷", [
        ("wyfw",   "物业服务合同纠纷"),
        ("mjjd",   "民间借贷纠纷"),
        ("jsgc",   "建设工程施工合同纠纷"),
        ("fwzl",   "房屋租赁合同纠纷"),
        ("maimai", "买卖合同纠纷"),
        ("fwmm",   "房屋买卖合同纠纷"),
        ("jrjk",   "金融借款合同纠纷"),
        ("xyk",    "信用卡纠纷"),
        ("rzzl",   "融资租赁合同纠纷"),
        ("ldzy",   "劳动争议纠纷"),
        ("lihun",  "离婚纠纷"),
        ("jtsg",   "机动车交通事故责任纠纷"),
        ("zqxjcs", "证券虚假陈述责任纠纷"),
        ("ccss",   "财产损失保险合同纠纷"),
        ("zrbx",   "责任保险合同纠纷"),
        ("bzbx",   "保证保险合同纠纷"),
        ("rsbx",   "人身保险合同纠纷"),
    ]),
    ("知识产权", [
        ("zhuzuoquan", "侵害著作权及邻接权纠纷"),
        ("shangbiao",  "侵害商标权纠纷"),
        ("fmzl",       "侵害发明专利权纠纷"),
        ("wgsj",       "侵害外观设计专利权纠纷"),
        ("zwxpz",      "侵害植物新品种权纠纷"),
        ("syms",       "侵害商业秘密纠纷"),
        ("jsht",       "技术合同纠纷"),
        ("bzdj",       "不正当竞争纠纷"),
        ("longduan",   "垄断纠纷"),
    ]),
    ("刑事自诉", [
        ("wuru", "侮辱案"), ("feibang", "诽谤案"),
        ("zhonghun", "重婚案"), ("jubuzx", "拒不执行判决、裁定案"),
    ]),
    ("海事纠纷", [
        ("cbpz", "船舶碰撞损害责任纠纷"),
        ("hsrs", "海上、通海水域人身损害责任纠纷"),
        ("hshyd", "海上、通海水域货运代理合同纠纷"),
        ("cylw", "船员劳务合同纠纷"),
    ]),
    ("公益诉讼", [
        ("hjwr", "环境污染民事公益诉讼"),
        ("stph", "生态破坏民事公益诉讼"),
        ("stsh", "生态环境损害赔偿诉讼"),
    ]),
    ("知产行政", [
        ("sbbhfs", "商标申请驳回复审纠纷"), ("sbcxfs", "商标撤销复审行政纠纷"),
        ("sbwx", "商标无效行政纠纷"), ("zlbhfs", "专利申请驳回复审行政纠纷"),
        ("zlwx", "专利无效行政纠纷"), ("ldxz", "垄断纠纷（行政）"),
    ]),
    ("行政诉讼", [
        ("xzcf", "行政处罚"), ("xzqzzx", "行政强制执行"), ("xzxk", "行政许可"),
        ("fwzs", "国有土地上房屋征收决定"), ("gsbx", "工伤保险资格或者待遇认定"),
        ("zfxxgk", "政府信息公开"), ("xzfy", "行政复议"), ("xzxy", "行政协议"),
        ("xzbc", "行政补偿"), ("xzpc", "行政赔偿"),
    ]),
]

CASE_TYPES = {}
for _cat, _items in CASE_TYPES_ORDERED:
    for _tid, _tname in _items:
        CASE_TYPES[_tid] = _tname


# ============================================================
# 自然人 / 法人 字段定义 (范本格式: 右列逐行)
# 每个字段: (模板文本, 对应info dict的key列表)
# ============================================================
def fill_person_fields(info):
    """填充自然人字段，info是包含各项详细信息的dict"""
    name = info.get('name', '')
    id_num = info.get('id_num', '')

    # 从身份证号自动提取性别和出生日期
    gender = info.get('gender', '')
    if not gender and id_num:
        gender = id_to_gender(id_num)
    gender_str = "男☑  女□" if gender == '男' else ("男□  女☑" if gender == '女' else "男□  女□")

    birth = info.get('birth', '')
    if not birth and id_num:
        birth = id_to_birth(id_num)
    if not birth:
        birth = "  年  月  日"

    nation = info.get('nation', '')
    work = info.get('work', '')
    job = info.get('job', '')
    phone = info.get('phone', '')
    addr = info.get('addr', '')
    habitual = info.get('habitual', '')
    id_type = info.get('id_type', '')
    if not id_type and id_num:
        id_type = '居民身份证'

    lines = [
        f"姓名：{name}",
        f"性别：{gender_str}",
        f"出生日期：{birth}    民族：{nation}",
        f"工作单位：{work}      职务：{job}      联系电话：{phone}",
        f"住所地（户籍所在地）：{addr}",
        f"经常居住地：{habitual}",
        f"证件类型：{id_type}",
        f"证件号码：{id_num}",
    ]
    return "\n".join(lines)


def fill_org_fields(info):
    """填充法人/非法人组织字段"""
    name = info.get('name', '')
    addr = info.get('addr', '')
    reg = info.get('reg_addr', '')
    legal_rep = info.get('legal_rep', '')
    job = info.get('job', '')
    phone = info.get('phone', '')
    credit = info.get('credit_code', '')
    org_type = info.get('org_type', '')
    ownership = info.get('ownership', '')

    # 类型勾选
    if not org_type:
        type_str = "有限责任公司□  股份有限公司□  上市公司□  其他企业法人□"
    else:
        type_str = org_type

    if not ownership:
        own_str = "国有□（控股□  参股□）  民营□  其他"
    else:
        own_str = ownership

    lines = [
        f"名称：{name}",
        f"住所地（主要办事机构所在地）：{addr}",
        f"注册地/登记地：{reg}",
        f"法定代表人/负责人：{legal_rep}      职务：{job}      联系电话：{phone}",
        f"统一社会信用代码：{credit}",
        f"类型：{type_str}",
        f"所有制性质：{own_str}",
    ]
    return "\n".join(lines)


def fill_agent_fields(info):
    """填充委托诉讼代理人字段"""
    has = info.get('has', '')
    has_str = "有☑  无□" if has else ("有□  无□" if not info.get('name') else "有☑  无□")
    name = info.get('name', '')
    job = info.get('job', '')
    firm = info.get('firm', '')
    phone = info.get('phone', '')
    auth = info.get('auth', '')
    if not auth:
        auth_str = "一般授权□    特别授权□"
    elif '特别' in auth:
        auth_str = "一般授权□    特别授权☑"
    else:
        auth_str = "一般授权☑    特别授权□"

    lines = [
        has_str,
        f"姓名：{name}      职务：{job}",
        f"单位：{firm}      联系电话：{phone}",
        f"代理权限：{auth_str}",
    ]
    return "\n".join(lines)


# ============================================================
# 模板定义
# ============================================================
def get_claim_fields(type_id):
    """返回诉讼请求的要素字段 (key, label)"""
    if type_id == "wyfw":
        return [
            ("claim_property_fee", "1. 物业费"),
            ("claim_penalty",      "2. 违约金"),
            ("claim_cost",         "3. 是否主张诉讼费用"),
            ("claim_other",        "4. 其他请求"),
            ("claim_total",        "5. 标的总额"),
        ]
    return [
        ("claim_full",  "诉讼请求"),
        ("claim_total", "标的总额"),
    ]

def get_facts_fields(type_id):
    """返回事实与理由的要素字段"""
    if type_id == "wyfw":
        return [
            ("facts_contract",    "1. 物业服务合同或前期物业服务合同签订情况（名称、编号、签订时间、地点等）"),
            ("facts_parties",     "2. 签订主体"),
            ("facts_property",    "3. 物业项目情况"),
            ("facts_fee_std",     "4. 约定的物业费标准"),
            ("facts_period",      "5. 约定的物业服务期限"),
            ("facts_pay_method",  "6. 约定的物业费支付方式"),
            ("facts_penalty_std", "7. 约定的逾期支付物业费违约金标准"),
            ("facts_unpaid",      "8. 被告欠付物业费数额及计算方式"),
            ("facts_penalty_amt", "9. 被告应付违约金数额及计算方式"),
            ("facts_collection",  "10. 催缴情况"),
            ("facts_other",       "11. 其他需要说明的内容（可另附页）"),
            ("facts_legal",       "12. 请求依据"),
            ("facts_evidence",    "13. 证据清单（可另附页）"),
        ]
    if type_id == "mjjd":
        return [
            ("facts_contract", "1. 合同签订情况"), ("facts_parties", "2. 签订主体"),
            ("facts_amount", "3. 借款金额"), ("facts_period", "4. 借款期限"),
            ("facts_rate", "5. 借款利率"), ("facts_time", "6. 借款提供时间"),
            ("facts_repay_method", "7. 还款方式"), ("facts_repay_status", "8. 还款情况"),
            ("facts_overdue", "9. 是否存在逾期还款"),
            ("facts_guarantee", "10. 担保情况"),
            ("facts_other", "11. 其他需要说明的内容"),
            ("facts_legal", "12. 请求依据"), ("facts_evidence", "13. 证据清单"),
        ]
    if type_id == "jsgc":
        return [
            ("facts_contract", "1. 施工合同签订情况"), ("facts_parties", "2. 签订主体"),
            ("facts_project", "3. 工程概况"), ("facts_price", "4. 工程价款约定"),
            ("facts_payment", "5. 付款方式和时间"), ("facts_period", "6. 工期约定"),
            ("facts_quality", "7. 质量标准"), ("facts_completion", "8. 竣工验收情况"),
            ("facts_settlement", "9. 工程结算情况"), ("facts_paid", "10. 工程款支付情况"),
            ("facts_sub", "11. 分包/转包情况"), ("facts_other", "12. 其他需要说明的内容"),
            ("facts_legal", "13. 请求依据"), ("facts_evidence", "14. 证据清单"),
        ]
    # 通用
    return [
        ("facts_basic", "1. 基本事实"), ("facts_contract", "2. 合同/法律关系情况"),
        ("facts_breach", "3. 违约/侵权/损害情况"), ("facts_amount", "4. 金额及计算"),
        ("facts_other", "5. 其他需要说明的内容"),
        ("facts_legal", "6. 请求依据"), ("facts_evidence", "7. 证据清单"),
    ]


# ============================================================
# 读取 docx（段落 + 表格）
# ============================================================
def read_docx(filepath):
    """读取docx文件，提取段落和表格中的所有文本"""
    doc = DocxDocument(filepath)
    lines = []

    # 读取段落
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            lines.append(t)

    # 读取表格（很多传统起诉状的当事人信息放在表格中）
    for table in doc.tables:
        for row in table.rows:
            row_texts = []
            for cell in row.cells:
                ct = cell.text.strip()
                if ct:
                    row_texts.append(ct)
            if row_texts:
                lines.append("\n".join(row_texts))

    return "\n".join(lines)


def id_to_birth(id_num):
    """从身份证号码中提取出生日期"""
    if not id_num:
        return ""
    id_num = id_num.strip()
    if len(id_num) == 18:
        y = id_num[6:10]
        m = id_num[10:12]
        d = id_num[12:14]
        return f"{y}年{int(m)}月{int(d)}日"
    elif len(id_num) == 15:
        y = "19" + id_num[6:8]
        m = id_num[8:10]
        d = id_num[10:12]
        return f"{y}年{int(m)}月{int(d)}日"
    return ""


def id_to_gender(id_num):
    """从身份证号码中提取性别"""
    if not id_num:
        return ""
    id_num = id_num.strip()
    if len(id_num) == 18:
        return "男" if int(id_num[16]) % 2 == 1 else "女"
    elif len(id_num) == 15:
        return "男" if int(id_num[14]) % 2 == 1 else "女"
    return ""


# ============================================================
# 智能提取
# ============================================================
def extract_elements(text, type_id):
    """从传统起诉状文本中提取要素信息，返回 dict"""
    data = {}

    # ============================================================
    # 辅助：从一段文本中提取主体详细信息
    # ============================================================
    def parse_person_detail(block):
        info = {}
        m = re.search(r'身份证号?[：:]\s*(\w{15,18}\w?)', block)
        if m:
            info['id_num'] = m.group(1).strip()
            info['id_type'] = '居民身份证'
        m = re.search(r'住所[：:]\s*(.+?)(?:\n|$)', block)
        if m:
            info['addr'] = m.group(1).strip()
        m = re.search(r'联系电话[：:]\s*(.+?)(?:\n|$)', block)
        if m:
            info['phone'] = m.group(1).strip()
        m = re.search(r'联系地址[：:]\s*(.+?)(?:\n|$)', block)
        if m:
            info['habitual'] = m.group(1).strip()
        m = re.search(r'邮编[：:]\s*(\d+)', block)
        if m:
            info['zip'] = m.group(1)
        return info

    def parse_org_detail(block):
        info = {}
        m = re.search(r'法定代表人[：:]\s*(.+?)(?:\n|$)', block)
        if m:
            info['legal_rep'] = m.group(1).strip()
        m = re.search(r'(?:住所|住址|地址)[：:]\s*(.+?)(?:\n|$)', block)
        if m:
            info['addr'] = m.group(1).strip()
        m = re.search(r'(?:统一(?:社会)?信用代码|信用代码)[：:]\s*(\w+)', block)
        if m:
            info['credit_code'] = m.group(1).strip()
        m = re.search(r'(?:联系方式|联系电话|电话)[：:]\s*([\d\-]+)', block)
        if m:
            info['phone'] = m.group(1).strip()
        return info

    # ============================================================
    # 按 "原告：""被告一：" 等标签行分段
    # ============================================================
    party_blocks = re.split(r'\n(?=(?:原告|被告[一二三四五六七八九十]?|第三人)[：:])', text)

    plaintiffs_info = []
    defendants_info = []
    thirds_info = []
    agent_info = {}

    for block in party_blocks:
        block = block.strip()
        if not block:
            continue

        if re.match(r'原告[：:]', block):
            nm = re.match(r'原告[：:]\s*(.+?)(?:\n|$)', block)
            if nm:
                pname = nm.group(1).strip()
                is_org = any(kw in pname for kw in ['公司','有限','集团','厂','所','中心','局','院'])
                info = parse_org_detail(block) if is_org else parse_person_detail(block)
                info['name'] = pname
                info['type'] = 'org' if is_org else 'person'
                plaintiffs_info.append(info)

        elif re.match(r'被告[一二三四五六七八九十]?[：:]', block):
            nm = re.match(r'被告[一二三四五六七八九十]?[：:]\s*(.+?)(?:\n|$)', block)
            if nm:
                dname = nm.group(1).strip()
                is_org = any(kw in dname for kw in ['公司','有限','集团','厂','所','中心','局','院'])
                info = parse_org_detail(block) if is_org else parse_person_detail(block)
                info['name'] = dname
                info['type'] = 'org' if is_org else 'person'
                defendants_info.append(info)

        elif re.match(r'第三人[：:]', block):
            nm = re.match(r'第三人[：:]\s*(.+?)(?:\n|$)', block)
            if nm:
                tname = nm.group(1).strip()
                is_org = any(kw in tname for kw in ['公司','有限','集团'])
                info = parse_org_detail(block) if is_org else parse_person_detail(block)
                info['name'] = tname
                info['type'] = 'org' if is_org else 'person'
                thirds_info.append(info)

    # 回退：从正文/具状人提取
    if not plaintiffs_info:
        m = re.search(r'具状人[：:]\s*(.+)', text)
        if m:
            for n in re.split(r'[、，,]', m.group(1).strip()):
                n = n.strip()
                if n:
                    is_org = any(kw in n for kw in ['公司','有限','集团','厂','所','中心','局','院'])
                    info = {'name': n, 'type': 'org' if is_org else 'person'}
                    # 补充从全文提取详细信息
                    if not is_org:
                        id_m = re.search(r'身份证号?[：:]\s*(\w{15,18}\w?)', text)
                        if id_m:
                            info['id_num'] = id_m.group(1)
                            info['id_type'] = '居民身份证'
                        addr_m = re.search(r'(?:^|\n)住所[：:]\s*(.+?)(?:\n|$)', text)
                        if addr_m and '被告' not in text[:text.find(addr_m.group(0))]:
                            info['addr'] = addr_m.group(1).strip()
                        phone_m = re.search(r'联系电话[：:]\s*(.+?)(?:\n|$)', text)
                        if phone_m:
                            info['phone'] = phone_m.group(1).strip()
                    plaintiffs_info.append(info)

    if not plaintiffs_info:
        m = re.search(r'原告[：:]\s*(.+?)(?:，|。|\n)', text)
        if m:
            pn = m.group(1).strip()
            is_org = any(kw in pn for kw in ['公司','有限','集团','厂','所','中心','局','院'])
            plaintiffs_info.append({'name': pn, 'type': 'org' if is_org else 'person'})

    # 回退：从事实与理由提取被告
    if not defendants_info:
        facts_sec = text
        fm = re.search(r'事实与理由[：:]?\s*([\s\S]*)', text)
        if fm:
            facts_sec = fm.group(1)
        for m in re.finditer(r'被告[一二三四五六七八九十]\s*(.+?)(?:（下称|自\d{4}|，|。)', facts_sec):
            name = m.group(1).strip()
            if name and 2 <= len(name) <= 40 and name not in [d.get('name','') for d in defendants_info]:
                is_org = any(kw in name for kw in ['公司','有限','集团','厂','所','中心','局','院'])
                defendants_info.append({'name': name, 'type': 'org' if is_org else 'person'})

    # 回退：从正文提取第三人
    if not thirds_info:
        for m in re.finditer(r'(?:第三人(?:为|系|是)|即第三人)\s*(.+?)(?:（|，|。)', text):
            name = m.group(1).strip()
            if name and 2 <= len(name) <= 30 and name not in [t.get('name','') for t in thirds_info]:
                is_org = any(kw in name for kw in ['公司','有限','集团'])
                thirds_info.append({'name': name, 'type': 'org' if is_org else 'person'})

    # ============================================================
    # 智能识别律师信息：如果原告的联系地址/联系电话含"律师""律所"，
    # 则这些信息属于委托诉讼代理人，不是原告的经常居住地
    # ============================================================
    LAWYER_KEYWORDS = ['律师', '律所', '律师事务所', '法律服务']

    def is_lawyer_info(text_str):
        return any(kw in text_str for kw in LAWYER_KEYWORDS)

    for pinfo in plaintiffs_info:
        # 检查联系地址是否含律师关键词
        contact_addr = pinfo.get('habitual', '')
        contact_phone = pinfo.get('phone', '')

        has_lawyer_addr = is_lawyer_info(contact_addr)
        has_lawyer_phone = is_lawyer_info(contact_phone)

        if has_lawyer_addr or has_lawyer_phone:
            # 将律师信息转移到agent_info
            if has_lawyer_addr:
                agent_info['firm'] = contact_addr
                pinfo.pop('habitual', None)  # 从原告中移除

            if has_lawyer_phone:
                # 提取律师姓名（如"13162490323（陈晓峰律师）"→ 姓名:陈晓峰）
                lawyer_name_m = re.search(r'[（(]([^）)]*?律师)[）)]', contact_phone)
                if lawyer_name_m:
                    raw = lawyer_name_m.group(1)
                    # "陈晓峰律师" → "陈晓峰"
                    agent_info['name'] = raw.replace('律师', '').strip()
                    agent_info['job'] = '律师'
                # 提取纯号码部分
                phone_num = re.match(r'([\d\-]+)', contact_phone)
                if phone_num:
                    agent_info['phone'] = phone_num.group(1)
                pinfo.pop('phone', None)  # 从原告中移除

            agent_info['has'] = True
            agent_info['auth'] = '特别授权'

    # 也检查原稿中是否有明确的委托代理人段落
    am = re.search(r'(?:委托(?:诉讼)?代理人|代理律师)[：:]\s*(.+?)(?:\n|$)', text)
    if am:
        agent_info['name'] = am.group(1).strip()
        agent_info['has'] = True
    # 联系地址可能单独出现
    if not agent_info.get('firm'):
        fm = re.search(r'联系地址[：:]\s*(.+?)(?:\n|$)', text)
        if fm and is_lawyer_info(fm.group(1)):
            agent_info['firm'] = fm.group(1).strip()

    # 保存
    data['plaintiffs_info'] = plaintiffs_info if plaintiffs_info else [{'name': '', 'type': 'person'}]
    data['defendants_info'] = defendants_info if defendants_info else [{'name': '', 'type': 'person'}]
    data['third_parties_info'] = thirds_info
    data['agent_info'] = agent_info
    data['plaintiffs'] = [p.get('name','') for p in plaintiffs_info]
    data['defendants'] = [{'name': d.get('name',''), 'type': d.get('type','person')} for d in defendants_info]
    data['third_parties'] = [{'name': t.get('name',''), 'type': t.get('type','person')} for t in thirds_info]

    # ---- 诉讼请求 ----
    m = re.search(r'(?:诉讼请求|请求贵院判令)[：:：]?\s*([\s\S]*?)(?=事实与理由)', text)
    if m:
        claims = m.group(1).strip()
        data['claim_full'] = claims
        amounts = re.findall(r'(?:人民币|共计|合计)\s*([\d,，.]+)\s*元', claims)
        if amounts:
            data['claim_total'] = amounts[-1].replace('，', ',') + '元'

    # ---- 事实与理由 ----
    m = re.search(r'事实与理由[：:]\s*([\s\S]*?)(?=此致)', text)
    if m:
        facts = m.group(1).strip()
        data['facts_basic'] = facts
        law_refs = []
        for lm in re.finditer(r'(《[^》]+》(?:第[^，。；\n]*?条)?)', facts):
            ref = lm.group(1)
            if ref not in law_refs:
                law_refs.append(ref)
        contract_refs = [r for r in law_refs if not any(kw in r for kw in ['民法典','合同法','物权法','侵权','公司法','劳动法'])]
        legal_refs = [r for r in law_refs if any(kw in r for kw in ['民法典','合同法','物权法','侵权','公司法','劳动法'])]
        parts = []
        if contract_refs:
            parts.append("合同约定：" + "、".join(contract_refs))
        if legal_refs:
            parts.append("法律规定：" + "、".join(legal_refs))
        if parts:
            data['facts_legal'] = "\n".join(parts)
        cm = re.search(r'《(.+?(?:合同|协议))》', facts)
        if cm:
            data['facts_contract'] = cm.group(1)
        pm = re.search(r'(?:系|位于|坐落)(.+?(?:\d+号\d+室|\d+室|\d+号))', text)
        if pm:
            data['facts_property'] = f"坐落位置：{pm.group(1).strip()}"
        period_m = re.search(r'自(\d{4}年\d{1,2}月\d{1,2}日)起', facts)
        if period_m:
            data['facts_period'] = period_m.group(1) + "起"
        pts = []
        if data.get('plaintiffs'):
            pts.append(f"业主/建设单位：{data['plaintiffs'][0]}")
        if data.get('defendants'):
            pts.append(f"物业服务人：{data['defendants'][0]['name']}")
        data['facts_parties'] = "\n".join(pts)
        fees = []
        for fm2 in re.finditer(r'(?:支出|支付|花费).+?([\d,，.]+)\s*元', facts):
            fees.append(fm2.group(1) + "元")
        if fees:
            data['facts_unpaid'] = "、".join(fees)

    m = re.search(r'此致\n(.+?法院)', text)
    if m:
        data['court'] = m.group(1).strip()
    m = re.search(r'日期[：:]\s*(.+)', text)
    if m:
        data['date'] = m.group(1).strip()

    return data


def auto_detect_type(text):
    keywords = {
        "wyfw": ["物业服务", "物业费", "物业管理"],
        "mjjd": ["民间借贷", "出借人", "借款人"],
        "jsgc": ["建设工程", "施工合同", "工程款"],
        "fwzl": ["房屋租赁", "租金", "租赁合同"],
        "ldzy": ["劳动争议", "劳动合同", "解除劳动"],
        "jtsg": ["交通事故", "机动车"],
        "lihun": ["离婚", "婚姻关系"],
        "maimai": ["买卖合同", "货款"],
        "fwmm": ["房屋买卖", "购房"],
        "jrjk": ["金融借款", "贷款合同"],
    }
    m = re.search(r'案由[：:]\s*(.+)', text)
    if m:
        cn = m.group(1).strip()
        for tid, tname in CASE_TYPES.items():
            if tname in cn or cn in tname:
                return tid
    for tid, kws in keywords.items():
        if any(kw in text for kw in kws):
            return tid
    return None


# ============================================================
# DOCX 生成 (精确还原范本格式)
# ============================================================
class DocxBuilder:
    """构建精确还原范本格式的docx"""

    def __init__(self):
        self.doc = DocxDocument()
        # 全局字体
        style = self.doc.styles['Normal']
        style.font.name = '宋体'
        style.font.size = Pt(10.5)
        style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        rpr = style.element.get_or_add_rPr()
        rpr.rFonts.set(qn('w:eastAsia'), '宋体')

    # ---- 底层工具 ----
    def _set_font(self, run, size=10.5, bold=False, font_name='宋体'):
        run.font.name = font_name
        run.font.size = Pt(size)
        run.font.bold = bold
        run.element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

    def _add_para(self, text, size=10.5, bold=False, align=None, font_name='宋体',
                  space_before=0, space_after=0, first_indent=None, right_indent=None):
        p = self.doc.add_paragraph()
        if align:
            p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(space_before)
        pf.space_after = Pt(space_after)
        if first_indent:
            pf.first_line_indent = Cm(first_indent)
        if right_indent:
            pf.right_indent = Cm(right_indent)
        run = p.add_run(text)
        self._set_font(run, size, bold, font_name)
        return p

    def _cell_borders(self, cell):
        """设置黑色单线边框"""
        tc = cell._element
        tcPr = tc.get_or_add_tcPr()
        for old in tcPr.findall(qn('w:tcBorders')):
            tcPr.remove(old)
        b = tcPr.makeelement(qn('w:tcBorders'), {})
        for edge in ['top','left','bottom','right']:
            el = b.makeelement(qn(f'w:{edge}'), {
                qn('w:val'): 'single', qn('w:sz'): '4',
                qn('w:space'): '0', qn('w:color'): '000000',
            })
            b.append(el)
        tcPr.insert(0, b)

    def _cell_shading(self, cell, color):
        tc = cell._element.get_or_add_tcPr()
        for old in tc.findall(qn('w:shd')):
            tc.remove(old)
        s = tc.makeelement(qn('w:shd'), {qn('w:fill'): color, qn('w:val'): 'clear'})
        tc.append(s)

    def _write_cell(self, cell, text, size=10.5, bold=False, align=None):
        """在单元格中写入文本（支持多行）"""
        cell.text = ''
        p = cell.paragraphs[0]
        if align:
            p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(1)
        pf.space_after = Pt(1)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        lines = text.split('\n')
        for i, line in enumerate(lines):
            if i > 0:
                r = p.add_run()
                r.add_break()
            r = p.add_run(line)
            self._set_font(r, size, bold)

    def _merge_vertical(self, table, col, row_start, row_end):
        """合并指定列的 row_start 到 row_end 行"""
        start_cell = table.cell(row_start, col)
        end_cell = table.cell(row_end, col)
        start_cell.merge(end_cell)

    # ---- 高层构建 ----
    def add_title(self, type_name):
        self._add_para('民事起诉状', size=22, bold=True,
                        align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2,
                        font_name='方正小标宋简体')
        self._add_para(f'（{type_name}）', size=14,
                        align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

    def add_notes(self):
        notes = [
            ("说明：", True),
            ("为了方便您更好地参加诉讼，保护您的合法权利，请填写本表。", False),
            ("1. 起诉时需向人民法院提交证明您身份的材料，如身份证复印件、营业执照复印件等。", False),
            ("2. 本表所列内容是您提起诉讼以及人民法院查明案件事实所需，请务必如实填写。", False),
            ('3. 本表有些内容可能与您的案件无关，您认为与案件无关的项目可以填"无"或不填；对于本表中勾选项可以在对应项打"√"；您认为另有重要内容需要列明的，可以另附页填写。', False),
            ("4. 本表word电子版填写时，相关栏目可复制粘贴或扩容，但不得改变要素内容、格式设置。", False),
        ]
        for text, bold in notes:
            self._add_para(text, size=9, bold=bold, space_before=1, space_after=1)
        self._add_para("★特别提示★", size=9, bold=True,
                        align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4)
        self._add_para("诉讼参加人应遵守诚信原则如实认真填写表格。", size=9,
                        align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
        self._add_para("如果诉讼参加人违反有关规定，虚假诉讼、恶意诉讼、滥用诉权，"
                        "人民法院将视违法情形依法追究责任。", size=9,
                        align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)


    def add_party_table(self, data):
        """
        构建当事人信息表。
        左列：角色标签（如"原告\n（自然人）"），跨对应行合并
        右列：该角色所有字段逐行展示，填入已提取的详细信息
        多个原告/被告/第三人 → 重复整块
        """
        blocks = []  # [(role_label, right_text), ...]

        # -- 原告 --
        for info in data.get('plaintiffs_info', [{'name':'', 'type':'person'}]):
            if info.get('type') == 'org':
                label = "原告\n（法人、非法人组织）"
                right = fill_org_fields(info)
            else:
                label = "原告\n（自然人）"
                right = fill_person_fields(info)
            blocks.append((label, right))

        # -- 委托诉讼代理人 --
        blocks.append(("委托诉讼代理人", fill_agent_fields(data.get('agent_info', {}))))

        # -- 被告 --
        d_list = data.get('defendants_info', [{'name':'', 'type':'person'}])
        if not d_list:
            d_list = [{'name':'', 'type':'person'}]
        for info in d_list:
            if info.get('type') == 'org':
                label = "被告\n（法人、非法人组织）"
                right = fill_org_fields(info)
            else:
                label = "被告\n（自然人）"
                right = fill_person_fields(info)
            blocks.append((label, right))

        # -- 第三人 --
        t_list = data.get('third_parties_info', [])
        for info in t_list:
            if info.get('type') == 'org':
                label = "第三人\n（法人、非法人组织）"
                right = fill_org_fields(info)
            else:
                label = "第三人\n（自然人）"
                right = fill_person_fields(info)
            blocks.append((label, right))

        if not t_list:
            blocks.append(("第三人\n（自然人）", fill_person_fields({})))
            blocks.append(("第三人\n（法人、非法人组织）", fill_org_fields({})))

        # 构建表格: 1(表头) + len(blocks) 行
        self.doc.add_paragraph()
        nrows = 1 + len(blocks)
        table = self.doc.add_table(rows=nrows, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True

        # 表头
        hdr = table.cell(0, 0)
        hdr.merge(table.cell(0, 1))
        self._write_cell(hdr, "当事人信息", size=11, bold=True,
                         align=WD_ALIGN_PARAGRAPH.CENTER)
        self._cell_borders(hdr)
        self._cell_shading(hdr, "D9E2F3")

        # 填充每个角色块
        for i, (label, right_text) in enumerate(blocks):
            row_idx = i + 1
            lc = table.cell(row_idx, 0)
            rc = table.cell(row_idx, 1)

            self._write_cell(lc, label, size=10.5)
            lc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            # 设置左列宽度
            lc.width = Cm(4.5)
            self._cell_borders(lc)

            self._write_cell(rc, right_text, size=10.5)
            rc.width = Cm(12)
            self._cell_borders(rc)

        # 设置列宽
        for row in table.rows:
            row.cells[0].width = Cm(4.5)
            row.cells[1].width = Cm(12)

    def add_claim_table(self, data, type_id):
        """诉讼请求表"""
        fields = get_claim_fields(type_id)
        intro = "（可完整表述诉讼请求；为方便、准确梳理要点，相关内容请在下方要素式表格中填写）"
        rows_data = [(intro, data.get('claim_full', ''))]
        for key, label in fields:
            rows_data.append((label, data.get(key, '')))
        self._add_standard_table("诉讼请求", rows_data)

    def add_jurisdiction_table(self, data):
        """约定管辖和诉前保全表"""
        rows_data = [
            ("1. 有无仲裁、法院管辖约定", "有□  无□\n合同条款及内容："),
            ("2. 是否已经诉前保全", "是□  否□\n保全法院：    保全时间：    保全案号：\n（如申请诉讼保全，请另行提交诉讼保全申请及相关材料）"),
        ]
        self._add_standard_table("约定管辖和诉前保全", rows_data)

    def add_facts_table(self, data, type_id):
        """事实与理由表"""
        fields = get_facts_fields(type_id)
        intro = "（可完整表述纠纷涉及的事实与理由；为方便、准确梳理要点，相关内容请在下方要素式表格中填写）"
        rows_data = [(intro, data.get('facts_basic', ''))]
        for key, label in fields:
            rows_data.append((label, data.get(key, '')))
        self._add_standard_table("事实与理由", rows_data)

    def add_mediation_table(self):
        """调解意愿表 - 默认勾选了解、否"""
        rows = [
            ("是否了解调解作为非诉讼纠纷解决方式，能及时、高效、低成本、不伤和气地解决纠纷",
             "了解☑    不了解□"),
            ("是否了解先行调解解决纠纷的好处",
             "1. 立案后选择先行调解的，可以很快启动调解程序。如不同意调解，法院将依程序开庭审理案件，但可能需要经过较长一段时间的排期等待，且审理、执行周期相对较长。\n了解☑  不了解□\n"
             "2. 选择先行调解，调解成功且自动履行的免交诉讼费用，申请司法确认的不交纳诉讼费用，要求出具调解书的减半交纳诉讼费用。\n了解☑  不了解□\n"
             "3. 首次调解不成功，但仍有继续调解意愿的，可以选择更换调解组织和调解员再进行调解。调解无法达成一致意见的，法院将依程序排期开庭。\n了解☑  不了解□\n"
             "4. 依照法律规定，调解具有保密性要求，调解过程不公开，调解协议未经当事人同意不得公开。\n了解☑  不了解□\n"
             "5. 调解达成的协议具有法律效力，可以依照法律规定申请司法确认，具有强制执行效力。\n了解☑  不了解□"),
            ("是否考虑先行调解", "是□    否☑    暂不确定，想要了解更多内容□"),
        ]
        self._add_standard_table("对纠纷解决方式的意愿", rows)

    def add_signature(self, data):
        """签名栏"""
        self.doc.add_paragraph()
        signer = ''
        if data.get('plaintiffs'):
            signer = data['plaintiffs'][0]
        date_str = data.get('date', '    年    月    日')
        self._add_para(f'具状人（签字、盖章）：{signer}',
                        align=WD_ALIGN_PARAGRAPH.RIGHT, right_indent=1, space_before=8)
        self._add_para(f'日期：{date_str}',
                        align=WD_ALIGN_PARAGRAPH.RIGHT, right_indent=1)

    def _add_standard_table(self, header, rows_data):
        """
        标准两列表格：蓝色表头 + (标签, 值) 行
        标签在左列灰底，值在右列
        """
        self.doc.add_paragraph()
        nrows = 1 + len(rows_data)
        table = self.doc.add_table(rows=nrows, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # 表头
        hdr = table.cell(0, 0)
        hdr.merge(table.cell(0, 1))
        self._write_cell(hdr, header, size=11, bold=True,
                         align=WD_ALIGN_PARAGRAPH.CENTER)
        self._cell_borders(hdr)
        self._cell_shading(hdr, "D9E2F3")

        for i, (label, value) in enumerate(rows_data):
            r = i + 1
            lc = table.cell(r, 0)
            rc = table.cell(r, 1)
            self._write_cell(lc, label, size=9.5)
            self._cell_borders(lc)
            self._cell_shading(lc, "F2F2F2")
            lc.width = Cm(6)
            self._write_cell(rc, value, size=10.5)
            self._cell_borders(rc)
            rc.width = Cm(10.5)

        for row in table.rows:
            row.cells[0].width = Cm(6)
            row.cells[1].width = Cm(10.5)

    def save(self, path):
        self.doc.save(path)


def generate_docx(data, output_path, type_id):
    type_name = CASE_TYPES.get(type_id, "通用")
    builder = DocxBuilder()
    builder.add_title(type_name)
    builder.add_notes()
    builder.add_party_table(data)
    builder.add_claim_table(data, type_id)
    builder.add_jurisdiction_table(data)
    builder.add_facts_table(data, type_id)
    builder.add_mediation_table()
    builder.add_signature(data)
    builder.save(output_path)
    return output_path


# ============================================================
# GUI
# ============================================================
class App:
    # 配色
    BG       = "#FAFAFA"
    CARD_BG  = "#FFFFFF"
    PRIMARY  = "#2B5797"
    PRIMARY_H= "#3A6DB5"
    TEXT     = "#333333"
    TEXT2    = "#888888"
    SUCCESS  = "#2E7D32"
    BORDER   = "#E0E0E0"

    def __init__(self, root):
        self.root = root
        self.root.title("要素式起诉状转换工具")
        self.root.geometry("760x620")
        self.root.configure(bg=self.BG)
        self.root.resizable(True, True)
        self.input_path = None
        self.original_text = ""
        self.build_ui()

    def build_ui(self):
        main = tk.Frame(self.root, bg=self.BG, padx=28, pady=20)
        main.pack(fill=tk.BOTH, expand=True)

        # ── 顶部标题 ──
        hdr = tk.Frame(main, bg=self.BG)
        hdr.pack(fill=tk.X, pady=(0, 20))
        tk.Label(hdr, text="要素式起诉状转换工具",
                 font=("SimHei", 18), bg=self.BG, fg=self.TEXT).pack(side=tk.LEFT)
        tk.Label(hdr, text="传统起诉状 → 要素式（.docx）",
                 font=("SimSun", 10), bg=self.BG, fg=self.TEXT2).pack(side=tk.LEFT, padx=(12,0), pady=(6,0))

        # ── 操作卡片 ──
        card = tk.Frame(main, bg=self.CARD_BG, bd=1, relief=tk.SOLID,
                        highlightbackground=self.BORDER, highlightthickness=1)
        card.pack(fill=tk.X, pady=(0, 16))
        inner = tk.Frame(card, bg=self.CARD_BG, padx=24, pady=20)
        inner.pack(fill=tk.X)

        # 第一行：选择文件
        row1 = tk.Frame(inner, bg=self.CARD_BG)
        row1.pack(fill=tk.X, pady=(0, 14))
        tk.Label(row1, text="起诉状文件", font=("SimHei", 10),
                 bg=self.CARD_BG, fg=self.TEXT, width=10, anchor='e').pack(side=tk.LEFT)
        self.btn_import = tk.Button(row1, text="选择文件",
            font=("SimSun", 10), bg=self.PRIMARY, fg="white",
            activebackground=self.PRIMARY_H, relief=tk.FLAT, padx=16, pady=4,
            cursor="hand2", command=self.import_file)
        self.btn_import.pack(side=tk.LEFT, padx=(12, 8))
        self.lbl_file = tk.Label(row1, text="支持 .docx 格式",
                                 font=("SimSun", 9), bg=self.CARD_BG, fg=self.TEXT2)
        self.lbl_file.pack(side=tk.LEFT)

        # 第二行：案由选择
        row2 = tk.Frame(inner, bg=self.CARD_BG)
        row2.pack(fill=tk.X, pady=(0, 14))
        tk.Label(row2, text="案由类型", font=("SimHei", 10),
                 bg=self.CARD_BG, fg=self.TEXT, width=10, anchor='e').pack(side=tk.LEFT)

        self.type_values = []
        self.type_map = {}
        for cat, items in CASE_TYPES_ORDERED:
            for tid, tname in items:
                d = f"【{cat}】{tname}"
                self.type_values.append(d)
                self.type_map[d] = tid

        self.combo_type = ttk.Combobox(row2, values=self.type_values,
                                        font=("SimSun", 10), width=36, state="readonly")
        self.combo_type.pack(side=tk.LEFT, padx=(12, 8))
        self.lbl_auto = tk.Label(row2, text="", font=("SimSun", 9),
                                  bg=self.CARD_BG, fg=self.SUCCESS)
        self.lbl_auto.pack(side=tk.LEFT)

        # 分隔线
        sep = tk.Frame(inner, bg=self.BORDER, height=1)
        sep.pack(fill=tk.X, pady=(4, 16))

        # 转换按钮
        self.btn_convert = tk.Button(inner, text="开始转换",
            font=("SimHei", 13), bg=self.PRIMARY, fg="white",
            activebackground=self.PRIMARY_H, relief=tk.FLAT, padx=48, pady=8,
            cursor="hand2", command=self.start_convert, state=tk.DISABLED)
        self.btn_convert.pack()

        # ── 预览/日志区 ──
        self.txt_log = scrolledtext.ScrolledText(
            main, height=14, font=("Consolas", 9),
            bg="#F5F5F5", fg="#555555", wrap=tk.WORD,
            state=tk.DISABLED, bd=1, relief=tk.SOLID,
            highlightbackground=self.BORDER, highlightthickness=1)
        self.txt_log.pack(fill=tk.BOTH, expand=True)
        self.log("就绪，请选择传统起诉状文件。")

    def log(self, msg):
        self.txt_log.configure(state=tk.NORMAL)
        self.txt_log.insert(tk.END, msg + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state=tk.DISABLED)

    def import_file(self):
        path = filedialog.askopenfilename(
            title="选择传统起诉状文件",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")])
        if not path:
            return
        self.input_path = path
        fname = os.path.basename(path)
        self.lbl_file.config(text=f"✓ {fname}", fg=self.SUCCESS)
        try:
            self.original_text = read_docx(path)
            self.log(f"已读取：{fname}（{len(self.original_text)} 字）")

            detected = auto_detect_type(self.original_text)
            if detected:
                for dn, tid in self.type_map.items():
                    if tid == detected:
                        self.combo_type.set(dn)
                        self.lbl_auto.config(text="已自动识别")
                        self.log(f"自动识别案由：{CASE_TYPES[detected]}")
                        break
            self.btn_convert.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            self.log(f"读取失败：{e}")

    def start_convert(self):
        if not self.original_text:
            messagebox.showwarning("提示", "请先选择起诉状文件。")
            return
        sel = self.combo_type.get()
        if not sel or sel not in self.type_map:
            messagebox.showwarning("提示", "请选择案由类型。")
            return
        type_id = self.type_map[sel]
        default_name = os.path.splitext(os.path.basename(self.input_path))[0]
        output_path = filedialog.asksaveasfilename(
            title="保存要素式起诉状",
            initialfile=f"{default_name}_要素式_{CASE_TYPES[type_id]}.docx",
            defaultextension=".docx",
            filetypes=[("Word文档", "*.docx")])
        if not output_path:
            return

        self.btn_convert.config(state=tk.DISABLED, text="正在转换…")
        self.root.update()

        def do_convert():
            try:
                self.log(f"\n正在转换：{CASE_TYPES[type_id]}")
                data = extract_elements(self.original_text, type_id)

                # 简要日志
                if data.get('plaintiffs'):
                    self.log(f"  原告：{', '.join(data['plaintiffs'])}")
                if data.get('defendants'):
                    self.log(f"  被告：{', '.join(d['name'] for d in data['defendants'])}")

                generate_docx(data, output_path, type_id)
                self.log(f"  ✓ 已保存：{output_path}")
                self.root.after(0, lambda: messagebox.showinfo("完成",
                    f"要素式起诉状已生成。\n\n{output_path}"))
            except Exception as e:
                import traceback
                self.log(f"  ✗ 失败：{e}")
                self.root.after(0, lambda: messagebox.showerror("转换失败", str(e)))
            finally:
                self.root.after(0, lambda: self.btn_convert.config(
                    state=tk.NORMAL, text="开始转换"))

        threading.Thread(target=do_convert, daemon=True).start()


# ============================================================
# 入口
# ============================================================
if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use('clam')
    except:
        pass
    # 统一 Combobox 样式
    style.configure('TCombobox', padding=4)
    App(root)
    root.mainloop()
