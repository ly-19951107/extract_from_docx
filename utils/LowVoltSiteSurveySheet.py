# -*- coding: utf-8 -*-
"""用于解析「低压现场勘察单」文档

一个文档会提取出
"""
from docx import Document
from docx.table import _Cell
from docx.opc.exceptions import PackageNotFoundError
import re
import pymysql
from uuid import uuid1
from config import db_config
from utils import initialize

SCHEME_ID = 'LVSSS'

# <<<<<配置区域
classes = {
    'low_volt_site_survey_sheet': '低压现场勘察单',
    "customer": "用户",
    "manager": "办理信息",
    "device": "设备信息"
}

data_properties = {
    "customer_number": {'domain': 'customer', "range": "string", "desc": "户号"},
    "application_number": {'domain': 'low_volt_site_survey_sheet', 'range': 'string', 'desc': '申请编号'},
    "name": {'domain': 'customer', "range": "string", "desc": "户名"},
    "contact_name": {'domain': 'customer', "range": "string", "desc": "联系人"},
    "contact_phone": {'domain': 'customer', "range": "string", "desc": "联系电话"},
    "elec_address": {'domain': 'customer', "range": "string", "desc": "客户地址"},
    "application_note": {'domain': 'customer', "range": "string", "desc": "申请备注"},

    "elec_type": {'domain': 'customer', "range": "string", "desc": "申请用电类别"},
    "elec_type_check": {'domain': 'customer', "range": "string", "desc": "核定情况"},
    "industry_class": {'domain': 'customer', "range": "string", "desc": "申请行业分类"},
    "industry_class_check": {'domain': 'customer', "range": "string", "desc": "核定情况"},
    "supply_volt": {'domain': 'customer', "range": "string", "desc": "申请供电电压"},
    "supply_volt_check": {'domain': 'customer', "range": "string", "desc": "核定供电电压"},
    "application_cap": {'domain': 'customer', "range": "string", "desc": "申请用电容量"},

    "cap_check": {'domain': 'manager', "range": "string", "desc": "核定用电容量："},
    "access_point_info": {'domain': 'manager', "range": "string", "desc": "接入点信息"},
    "receive_point_info": {'domain': 'manager', "range": "string", "desc": "受电点信息"},
    "meter_point_info": {'domain': 'manager', "range": "string", "desc": "计量点信息"},
    "other_ins": {'domain': 'manager', "range": "string", "desc": "其他"},

    "device_name": {'domain': 'device', "range": "string", "desc": "设备名称"},
    "device_type": {'domain': 'device', "range": "string", "desc": "型号"},
    'device_num': {'domain': 'device', 'range': 'string', 'desc': '数量'},
    "total_cap": {'domain': 'device', "range": "string", "desc": "总容量（千瓦）"},
    "remark": {'domain': 'device', "range": "string", "desc": "备注"},

    "assignee": {'domain': 'low_volt_site_survey_sheet', 'range': 'string', 'desc': '勘查人'},
    "accept_date": {'domain': 'low_volt_site_survey_sheet', 'range': 'string', 'desc': '勘查日期'}
}

object_properties = {
    0: {
        'domain': 'low_volt_site_survey_sheet',
        'range': 'customer',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述表格属于哪个客户',
    },
    1: {
        'domain': 'low_volt_site_survey_sheet',
        'range': 'manager',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述表格的办理信息',
    },
    2: {
        'domain': 'low_volt_site_survey_sheet',
        'range': 'device',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述设备信息',
    }
}


def iter_visible_cells(row):
    tr = row._tr
    for tc in tr.tc_lst:
        yield _Cell(tc, row._parent.table)


def read_file(file_path):
    """读取一个docx文件"""
    try:
        docx = Document(file_path)
    except PackageNotFoundError:
        print(f'路径不正确或目标为加密文档：{file_path}')
        return
    class_ = [data_properties[i]['domain'] for i in data_properties]   # 每个属性对应的实体
    pros = [i for i in data_properties.keys()]   # 所有属性即表字段名
    keys = [data_properties[i]['desc'] for i in data_properties]  # 所有属性中文名称即表格中给出的属性
    message = []
    table = docx.tables[0]  # 表格
    # 取前15行内容
    for r in range(1, 16):
        if r == 6 or r == 15:
            continue
        row = table.rows[r]
        row_cells = list(iter_visible_cells(row))
        if r in range(1, 6):
            for i in row_cells[:-1]:
                message.append(i.text.replace(' ', ''))
        elif r in range(7, 10):
            for i in row_cells[:-1]:
                message.append(i.text.replace(' ', ''))
            line = row_cells[-1].text.replace(' ', '')
            info = re.compile(r'(.*)：(.*)')
            v1 = info.match(line).group(1)
            v2 = info.match(line).group(2)
            message.extend([v1, v2])
        else:
            for i in row_cells:
                message.append(i.text.replace(' ', ''))

    values = []  # 保存所有值
    str = ''
    for s in range(len(message)):
        if message[s] in keys:
            if str != '':
                values.append(str)
                str = ''
        elif s == len(message) - 1:
            values.append(message[s])
        else:
            str += message[s]
    # 保存前14行内容
    entity_dict = {}
    num = len(values)
    for c in range(num):
        if class_[c] not in entity_dict:
            entity = Entity(class_[c], uuid1().hex)
            entity_dict[class_[c]] = entity
        else:
            entity = entity_dict[class_[c]]
        entity.add_pro(pros[c], values[c])

    # 取设备信息值
    for r in range(-5, -2):
        row = table.rows[r]
        row_cells = list(iter_visible_cells(row))
        for i in row_cells:
            values.append(i.text.replace(' ', ''))

    # 保存设备信息内容
    n1 = num
    n2 = len(class_) - 2
    n = n2 - n1
    count = 0
    for c in range(-5, - 2):
        entity = Entity(class_[num], uuid1().hex)
        for j in range(n1, n2):
            entity.add_pro(pros[j - count * n], values[j])
        if class_[num] not in entity_dict:
            entity_dict[class_[num]] = [entity]
        else:
            entity_dict[class_[num]].append(entity)
        n1 += n
        n2 += n
        count += 1

    # 取最后一行值
    line = []
    row = table.rows[-1]
    row_cells = list(iter_visible_cells(row))
    for i in row_cells:
        line.append(i.text.replace(' ', ''))
    for t in line:
        for k in keys:
            if k in t:
                line.remove(t)
    values.extend(line)
    # 保存最后一行内容
    for c in range(len(class_) - 2, len(class_)):
        if class_[c] not in entity_dict:
            entity = Entity(class_[c], uuid1().hex)
            entity_dict[class_[c]] = entity
        else:
            entity = entity_dict[class_[c]]
        entity.add_pro(pros[c], values[c - len(class_)])
    return entity_dict


def save(entity_dict):
    """将提取的结果存入对应的数据库"""
    conn = pymysql.connect(**db_config)
    cr = conn.cursor()
    # 存实体
    for class_ in entity_dict:
        tab = SCHEME_ID + '_' + class_
        if isinstance(entity_dict[class_], Entity):
            id_ = entity_dict[class_].id_
            pros = entity_dict[class_].pros
            sql = f'insert into `{tab}`(`id`,'
            values = []
            for pro in pros:
                sql += f'`{pro}`,'
                values.append(pros[pro])
            sql = sql[:-1]
            sql += f') values ("{id_}",'
            for v in values:
                sql += f'"{v}",'
            sql = sql[:-1] + ')'
            cr.execute(sql)
        else:
            for entity in entity_dict[class_]:
                id_ = entity.id_
                pros = entity.pros
                sql = f'insert into `{tab}`(`id`,'
                values = []
                for pro in pros:
                    sql += f'`{pro}`,'
                    values.append(pros[pro])
                sql = sql[:-1]
                sql += f') values ("{id_}",'
                for v in values:
                    sql += f'"{v}",'
                sql = sql[:-1] + ')'
                cr.execute(sql)
    conn.commit()

    # 存关系
    for i in object_properties:
        rel = object_properties[i]
        domain = rel['domain']
        range_ = rel['range']
        rel_tab = SCHEME_ID + '_' + domain + '_2_' + range_
        from_id = entity_dict[domain].id_
        if isinstance(entity_dict[range_], Entity):
            to_ids = [entity_dict[range_].id_]
        else:
            to_ids = [j.id_ for j in entity_dict[range_]]
        for to_id in to_ids:
            sql = f'''insert into `{rel_tab}` (`id`, `from_id`, `to_id`) values (
                "{uuid1().hex}", "{from_id}", "{to_id}"
            )
            '''
            cr.execute(sql)
    conn.commit()
    conn.close()


class Entity:
    """实例表示从模板中提取出来的一个实体"""
    def __init__(self, class_, id_):
        self.class_ = class_
        self.pros = {}
        self.id_ = id_

    def add_pro(self, key, value):
        if isinstance(key, str) and isinstance(value, str):
            if key in self.pros:
                if not self.pros[key]:
                    self.pros[key] = value
                else:
                    if value:
                        self.pros[key] += '/' + value
            else:
                self.pros.update({key: value})
        else:
            raise


if __name__ == '__main__':
    file_path = r'C:\Users\liyang\Desktop\extract\extract_from_docx\templates\低压现场勘察单.docx'
    initialize(SCHEME_ID, classes, data_properties, object_properties)
    save(read_file(file_path))
