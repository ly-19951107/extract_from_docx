# -*- coding: utf-8 -*-
"""用于解析「低压批量用电登记表」文档

一个文档会提取出
"""
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
import re
import pymysql
from uuid import uuid1
from config import db_config
from utils import initialize

SCHEME_ID = 'LVBERF'

# <<<<<配置区域
classes = {
    'low_volt_batch_elec_regis_form': '低压批量用电登记表',
    "customer": "用户",
    "manager": "办理信息"
}
data_properties = {
    "name": {'domain': 'customer', "range": "string", "desc": "户名"},
    "customer_number": {'domain': 'customer', "range": "string", "desc": "户号"},
    "elec_address": {'domain': 'customer', "range": "string", "desc": "用电地址"},
    "elec_type": {'domain': 'customer', "range": "string", "desc": "用电类别"},
    "application_households_num": {'domain': 'customer', "range": "string", "desc": "申请户数"},
    "single_house_cap": {'domain': 'customer', "range": "string", "desc": "单户容量"},
    "total_cap": {'domain': 'customer', "range": "string", "desc": "总容量"},

    "manager_unit": {'domain': 'manager', "range": "string", "desc": "经办单位"},
    "unit_address": {'domain': 'manager', "range": "string", "desc": "单位地址"},
    "unit_contact_address": {'domain': 'manager', "range": "string", "desc": "通信地址"},
    "unit_postcode": {'domain': 'manager', "range": "string", "desc": "邮编"},
    "unit_E-mail": {'domain': 'manager', "range": "string", "desc": "电子邮箱"},
    "unit_fax": {'domain': 'manager', "range": "string", "desc": "传真"},
    "manager_name": {'domain': 'manager', "range": "string", "desc": "经办人"},
    "manager_ID_number": {'domain': 'manager', "range": "string", "desc": "身份证号"},
    "manager_fixed_tel": {'domain': 'manager', "range": "string", "desc": "固定电话"},
    "manager_mobile_phone": {'domain': 'manager', "range": "string", "desc": "移动电话"},

    "assignee": {'domain': 'low_volt_batch_elec_regis_form', 'range': 'string', 'desc': '受理人'},
    "application_number": {'domain': 'low_volt_batch_elec_regis_form', 'range': 'string', 'desc': '申请编号'},
    "accept_date": {'domain': 'low_volt_batch_elec_regis_form', 'range': 'string', 'desc': '受理日期'},
    "power_supply_company": {'domain': 'low_volt_batch_elec_regis_form', 'range': 'string', 'desc': '供电企业'}
}

object_properties = {
    0: {
        'domain': 'low_volt_batch_elec_regis_form',
        'range': 'customer',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述表格属于哪个客户',
    },
    1: {
        'domain': 'low_volt_batch_elec_regis_form',
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
    message = []
    cell_set = []
    table = docx.tables[0]
    for row in range(1, len(table.rows)):
        if row == 6 or row == 13:
            continue
        for cell in table.rows[row].cells:
            if cell not in cell_set:
                cell_set.append(cell)
                text = cell.text.replace(' ', '')
                message.append(text)
    values = []
    tmp = []
    str = ''
    for s in range(len(message)):
        if message[s] in keys:
            if message[s] not in tmp:
                tmp.append(message[s])
                if str != '':
                    values.append(str)
                    str = ''
                continue
            else:
                continue
        elif s == len(message) - 8:
            values.append(str)
        else:
            str += message[s]
    for k in keys[-4:]:
        for j in message[-6:]:
            if k in j:
                info = re.compile(r'(.*)：(.*)')
                v = info.match(j).group(2)
                values.append(v)
            else:
                continue
    entity_dict = {}
    for c in range(len(class_)):
        if class_[c] not in entity_dict:
            entity = Entity(class_[c], uuid1().hex)
            entity_dict[class_[c]] = entity
        else:
            entity = entity_dict[class_[c]]
        entity.add_pro(pros[c], values[c])
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
    file_path = r'C:\Users\liyang\Desktop\extract\extract_from_docx\templates\低压批量用电登记表.docx'
    initialize(SCHEME_ID, classes, data_properties, object_properties)
    save(read_file(file_path))
