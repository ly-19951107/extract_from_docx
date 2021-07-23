# -*- coding: utf-8 -*-
"""用于解析「低压批量用电清单」文档

一个文档会提取出
"""

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
import re
import pymysql
from uuid import uuid1
from config import db_config
from utils import initialize

SCHEME_ID = 'LVBEL'

# <<<<<配置区域
classes = {
    'low_volt_batch_elec_list': '低压批量用电清单',
    "customer": "用户",
    "manager": "办理信息"
}

data_properties = {
    "manager_unit": {'domain': 'manager', "range": "string", "desc": "经办单位"},
    "application_number": {'domain': 'low_volt_batch_elec_list', 'range': 'string', 'desc': '申请编号'},
    "elec_address": {'domain': 'manager', "range": "string", "desc": "用电地址"},

    # "household_no": {'domain': 'customer', "range": "string", "desc": "序号"},
    "customer_number": {'domain': 'customer', "range": "string", "desc": "室号"},
    "name": {'domain': 'customer', "range": "string", "desc": "户名"},
    'elec_cap': {'domain': 'customer', 'range': 'string', 'desc': '用电容量'},
    "ID_number": {'domain': 'customer', "range": "string", "desc": "身份证号"},
    "customer_mobile_phone": {'domain': 'customer', "range": "string", "desc": "移动电话"},

    "manager_name": {'domain': 'manager', "range": "string", "desc": "经办人"},
    "accept_date": {'domain': 'low_volt_batch_elec_list', 'range': 'string', 'desc': '受理日期'}
}

object_properties = {
    0: {
        'domain': 'low_volt_batch_elec_list',
        'range': 'customer',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述表格属于哪个客户',
    },
    1: {
        'domain': 'low_volt_batch_elec_list',
        'range': 'manager',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述表格的办理信息',
    }
}


def read_file(file_path):
    """读取一个docx文件"""
    try:
        docx = Document(file_path)
    except PackageNotFoundError:
        print(f'路径不正确或目标为加密文档：{file_path}')
        return
    class_ = [data_properties[i]['domain'] for i in data_properties]
    pros = [i for i in data_properties.keys()]
    keys = [data_properties[i]['desc'] for i in data_properties]
    entity_dict = {}
    table = docx.tables[0]
    message = []
    values = []
    cell_set = []
    # 取前两行内容并保存
    for row in range(0, 2):
        for cell in table.rows[row].cells:
            if cell not in cell_set:
                cell_set.append(cell)
                text = cell.text.replace(' ', '')
                message.append(text)
    str = ''
    for s in range(len(message)):
        if message[s] in keys:
            if str != '':
                values.append(str)
                str = ''
            continue
        elif s == 5:
            elec_address = re.compile(r'(.*)幢(.*)_单元')
            elec_address = elec_address.match(message[s]).group()
            values.append(elec_address)
        else:
            str += message[s]
    num = len(values)
    for c in range(num):
        if class_[c] not in entity_dict:
            entity = Entity(class_[c], uuid1().hex)
            entity_dict[class_[c]] = entity
        else:
            entity = entity_dict[class_[c]]
        entity.add_pro(pros[c], values[c])

    # 取中间内容并保存
    for row in range(3, len(table.rows) - 1):
        line = []
        for cell in table.rows[row].cells:
            if cell not in cell_set:
                cell_set.append(cell)
                text = cell.text.replace(' ', '')
                line.append(text)
        values.extend(line[1:])
    n1 = num
    n2 = len(class_) - 2
    n = n2 - n1
    count = 0
    for c in range(3, len(table.rows) - 1):
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

    # 取最后一行并保存
    for cell in table.rows[-1].cells:
        if cell not in cell_set:
            cell_set.append(cell)
            message.append(cell.text)
    info = re.compile(r'(.*)：(.*)')
    v = info.match(message[-1]).group(2)
    v1 = v.split(' ')[0]
    v2 = v.replace(v1, '').replace(' ', '')
    values.append(v1)
    values.append(v2)
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
    file_path = r'C:\Users\liyang\Desktop\extract_from_docx\templates\低压批量用电清单.docx'
    initialize(SCHEME_ID, classes, data_properties, object_properties)
    save(read_file(file_path))
