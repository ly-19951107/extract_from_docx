# -*- coding: utf-8 -*-
"""用于解析「高压客户用电登记表」文档

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

SCHEME_ID = 'HVCERF'

# <<<<<配置区域
classes = {
    'high_volt_cus_elec_regis_form': '高压客户用电登记表',
    "customer": "用户",
    "manager": "办理信息",
    "elec_demand": "用电需求"
}
data_properties = {
    "name": {'domain': 'customer', "range": "string", "desc": "户名"},
    "customer_number": {'domain': 'customer', "range": "string", "desc": "户号"},
    'customer_ID_name': {'domain': 'customer', 'range': 'string', 'desc': '（证件名称）'},
    'customer_ID_number': {'domain': 'customer', 'range': 'string', 'desc': '（证件号码）'},
    'industry_class': {'domain': 'customer', 'range': 'string', 'desc': '行业'},
    'VIP_client': {'domain': 'customer', 'range': 'string', 'desc': '重要客户'},
    "elec_address": {'domain': 'customer', "range": "string", "desc": "用电地址"},
    "contact_address": {'domain': 'customer', "range": "string", "desc": "通信地址"},
    "postcode": {'domain': 'customer', "range": "string", "desc": "邮编"},
    "E-mail": {'domain': 'customer', "range": "string", "desc": "电子邮箱"},
    "legal_representative": {'domain': 'customer', "range": "string", "desc": "法人代表"},
    "ID_number": {'domain': 'customer', "range": "string", "desc": "身份证号"},
    "customer_fixed_tel": {'domain': 'customer', "range": "string", "desc": "固定电话"},
    "customer_mobile_phone": {'domain': 'customer', "range": "string", "desc": "移动电话"},

    "manager_name": {'domain': 'manager', "range": "string", "desc": "经办人"},
    "manager_ID_number": {'domain': 'manager', "range": "string", "desc": "身份证号"},
    "manager_fixed_tel": {'domain': 'manager', "range": "string", "desc": "固定电话"},
    "manager_mobile_phone": {'domain': 'manager', "range": "string", "desc": "移动电话"},

    "business_type": {'domain': 'elec_demand', "range": "string", "desc": "业务类型"},
    "elec_type": {'domain': 'elec_demand', "range": "string", "desc": "用电类别"},
    "power_cap": {'domain': 'elec_demand', "range": "string", "desc": "电源容量"},
    "self_power": {'domain': 'elec_demand', "range": "string", "desc": "自备电源"},
    "self_power_cap": {'domain': 'elec_demand', "range": "string", "desc": "容量"},
    "VAT_invoice": {'domain': 'elec_demand', "range": "string", "desc": "需要增值税发票"},
    "non_line_load": {'domain': 'elec_demand', "range": "string", "desc": "非线性负荷"},

    "assignee": {'domain': 'high_volt_cus_elec_regis_form', 'range': 'string', 'desc': '受理人'},
    "application_number": {'domain': 'high_volt_cus_elec_regis_form', 'range': 'string', 'desc': '申请编号'},
    "accept_date": {'domain': 'high_volt_cus_elec_regis_form', 'range': 'string', 'desc': '受理日期'},
    "power_supply_company": {'domain': 'high_volt_cus_elec_regis_form', 'range': 'string', 'desc': '供电企业'}
}

object_properties = {
    0: {
        'domain': 'high_volt_cus_elec_regis_form',
        'range': 'customer',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述表格属于哪个客户',
    },
    1: {
        'domain': 'high_volt_cus_elec_regis_form',
        'range': 'manager',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述表格的办理信息',
    },
    2: {
        'domain': 'customer',
        'range': 'elec_demand',
        'name': 'need',
        'ZH_name': '需要',
        'desc': '描述用户的用电需求',
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
    class_ = [data_properties[i]['domain'] for i in data_properties]
    pros = [i for i in data_properties.keys()]
    keys = [data_properties[i]['desc'] for i in data_properties]
    message = []  # 保存前15行内容
    mes1 = []  # 保存19-20行内容
    val = []  # 保存16-18行的值
    val1 = []  # 保存最后两行的值
    table = docx.tables[0]

    for r in range(len(table.rows)):
        if r in [10, 13, len(table.rows) - 3]:
            continue
        row = table.rows[r]
        row_cells = list(iter_visible_cells(row))
        # 取前15行内容
        if r in range(1, 16):
            for i in row_cells:
                message.append(i.text.replace(' ', '').replace('\n', ''))
        # 保存电源容量信息，即16-18行的值
        elif r in range(16, len(table.rows) - 5):
            dic = {}
            mes = []
            for i in row_cells[: -1]:
                mes.append(i.text.replace(' ', '').replace('\n', ''))
            line = row_cells[-1].text
            if line == '':
                v1 = v2 = v3 = v4 = ''
            else:
                info = re.compile(r'(.*)：(.*) (.*)：(.*)')
                v1 = info.match(line).group(1).replace(' ', '').replace('\n', '')
                v2 = info.match(line).group(2).replace(' ', '').replace('\n', '')
                v3 = info.match(line).group(3).replace(' ', '').replace('\n', '')
                v4 = info.match(line).group(4).replace(' ', '').replace('\n', '')
            mes.extend([v1, v2, v3, v4])
            for i in range(0, len(mes), 2):
                dic[mes[i]] = mes[i + 1]
            val.append(dic)
        # 保存19-20行内容
        elif r in range(len(table.rows) - 5, len(table.rows) - 3):
            if r == len(table.rows) - 5:
                for i in row_cells[: -1]:
                    mes1.append(i.text.replace(' ', '').replace('\n', ''))
                line = row_cells[-1].text
                info = re.compile(r'(.*)：(.*)')
                v1 = info.match(line).group(1).replace(' ', '').replace('\n', '')
                v2 = info.match(line).group(2).replace(' ', '').replace('\n', '')
                mes1.extend([v1, v2])
            else:
                for i in row_cells:
                    mes1.append(i.text.replace(' ', '').replace('\n', ''))
        # 保存最后两行的值
        else:
            for i in row_cells[1:]:
                info = re.compile(r'(.*)：(.*)')
                v = info.match(i.text).group(2).replace(' ', '').replace('\n', '')
                val1.append(v)

    # 保存前15行的值
    values = []  # 保存所有值信息
    str = ''
    # 保存前15行的值
    for s in range(len(message)):
        if message[s] in keys:
            if str != '':
                values.append(str)
                str = ''
            continue
        elif s == len(message) - 1:
            values.append(message[s])
            str = ''
        else:
            str += message[s]
    # 保存16-18行的值
    values.append(val)
    # 保存19-20行的值
    for s in range(len(mes1)):
        if mes1[s] in keys:
            if str != '':
                values.append(str)
                str = ''
            continue
        elif s == len(mes1) - 1:
            values.append(mes1[s])
        else:
            str += mes1[s]
    # 保存后两行的值
    values.extend(val1)

    # 保存实体信息
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
        elif isinstance(key, str) and isinstance(value, list):
            if key in self.pros:
                if not self.pros[key]:
                    self.pros[key] = value
                else:
                    if value:
                        self.pros[key].append(value)
            else:
                self.pros.update({key: value})
        else:
            raise


if __name__ == '__main__':
    file_path = r'C:\Users\liyang\Desktop\extract\extract_from_docx\templates\高压客户用电登记表.docx'
    initialize(SCHEME_ID, classes, data_properties, object_properties)
    save(read_file(file_path))
