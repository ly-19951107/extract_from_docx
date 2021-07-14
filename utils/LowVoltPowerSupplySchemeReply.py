# -*- coding: utf-8 -*-
"""用于解析「低压供电方案答复单」文档
"""
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
import pymysql
from uuid import uuid1
from collections import OrderedDict
from config import db_config
from utils import initialize

SCHEME_ID = 'LVPSSR'

classes = {
    'high_volt_power_supply_schema_reply': '高压供电方案答复单',
    'customer': "用户",
    'charge': '营业费用',
    'scheme': '供电方案'
}
data_properties = {
    'customer_id': {'domain': 'customer', 'desc': '户号'},
    'customer_name': {'domain': 'customer', 'desc': '户名'},
    'apply_id': {'domain': 'customer', 'desc': '申请编号'},
    'addr': {'domain': 'customer', 'desc': '用电地址'},
    'type': {'domain': 'customer', 'desc': '用电类别'},
    'industry_class': {'domain': 'customer', 'desc': '行业分类'},
    'volt': {'domain': 'customer', 'desc': '供电电压'},
    'cap': {'domain': 'customer', 'desc': '供电容量'},
    'contacts': {'domain': 'customer', 'desc': '联系人'},
    'contact_phone': {'domain': 'customer', 'desc': '联系电话'},

    'charge_name': {'domain': 'charge', 'desc': '费用名称'},
    'unit_price': {'domain': 'charge', 'desc': '单价'},
    'num': {'domain': 'charge', 'desc': '数量/容量'},
    'amount_receivable': {'domain': 'charge', 'desc': '应收金额'},
    'charge_basis': {'domain': 'charge', 'desc': '收费依据'},

    'sign_date': {'domain': 'high_volt_power_supply_schema_reply', 'desc': '签订日期'},

    'pow_src_id': {'domain': 'scheme', 'desc': '电源编号'},
    'pow_src_nature': {'domain': 'scheme', 'desc': '电源性质'},
    'pow_volt': {'domain': 'scheme', 'desc': '供电电压'},
    'pow_cap': {'domain': 'scheme', 'desc': '供电容量'},
    'pow_src_info': {'domain': 'scheme', 'desc': '电源点信息'},
    'm_group_num': {'domain': 'scheme', 'desc': '计量点组号'},
    'price_type': {'domain': 'scheme', 'desc': '电价类别'},
    'dldb': {'domain': 'scheme', 'desc': '定量定比'},
    'meter_precision': {'domain': 'scheme', 'desc': '电能表精度'},
    'meter_norm': {'domain': 'scheme', 'desc': '电能表规格及接线方式'},
    'cur_trans_precision': {'domain': 'scheme', 'desc': '电流互感器精度'},
    'cur_trans_info': {'domain': 'scheme', 'desc': '电流互感器变比'},
}
object_properties = {
    0: {
        'domain': 'customer',
        'range': 'charge',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述客户与收费方式之间的关系',
    },
    2: {
        'domain': 'customer',
        'range': 'scheme',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述客户与供电方案之间的关系',
    }
}


def read_file(file_path):
    """读取一个docx文件"""
    try:
        docx = Document(file_path)
    except PackageNotFoundError:
        print(f'路径不正确或目标为加密文档：{file_path}')
        return
    table = docx.tables[0]
    customer = Entity('customer', uuid1().hex)
    charge = Entity('charge', uuid1().hex)
    scheme = Entity('scheme', uuid1().hex)
    i = 0
    while i < len(table.rows):
        cells = table.rows[i].cells
        distinct_cells = OrderedDict()
        for cell in cells:
            if id(cell) in distinct_cells:
                continue
            else:
                distinct_cells[id(cell)] = cell.text.strip()
        if i == 1:
            customer.add_pro('customer_id', list(distinct_cells.values())[1])
            customer.add_pro('apply_id', list(distinct_cells.values())[3])
            i += 1
        elif i == 2:
            customer.add_pro('customer_name', list(distinct_cells.values())[1])
            i += 1
        elif i == 3:
            customer.add_pro('addr', list(distinct_cells.values())[1])
            i += 1
        elif i == 4:
            customer.add_pro('type', list(distinct_cells.values())[1])
            customer.add_pro('industry_class', list(distinct_cells.values())[3])
            i += 1
        elif i == 5:
            customer.add_pro('volt', list(distinct_cells.values())[1])
            customer.add_pro('cap', list(distinct_cells.values())[3])
            i += 1
        elif i == 6:
            customer.add_pro('contacts', list(distinct_cells.values())[1])
            customer.add_pro('contact_phone', list(distinct_cells.values())[3])
            i += 1
        elif i == 9:
            charge.add_pro('charge_name', list(distinct_cells.values())[0])
            charge.add_pro('unit_price', list(distinct_cells.values())[1])
            charge.add_pro('num', list(distinct_cells.values())[2])
            charge.add_pro('amount_receivable', list(distinct_cells.values())[3])
            charge.add_pro('charge_basis', list(distinct_cells.values())[4])
            i += 1
        elif i == 13:
            scheme.add_pro('pow_src_id', list(distinct_cells.values())[0])
            scheme.add_pro('pow_src_nature', list(distinct_cells.values())[1])
            scheme.add_pro('pow_volt', list(distinct_cells.values())[2])
            scheme.add_pro('pow_cap', list(distinct_cells.values())[3])
            scheme.add_pro('pow_src_info', list(distinct_cells.values())[4])
            i += 1
        elif i == 16:
            scheme.add_pro('m_group_num', list(distinct_cells.values())[0])
            scheme.add_pro('price_type', list(distinct_cells.values())[1])
            scheme.add_pro('dldb', list(distinct_cells.values())[2])
            scheme.add_pro('meter_precision', list(distinct_cells.values())[3])
            scheme.add_pro('meter_norm', list(distinct_cells.values())[4])
            scheme.add_pro('cur_trans_precision', list(distinct_cells.values())[5])
            scheme.add_pro('cur_trans_info', list(distinct_cells.values())[6])
            i += 1
        else:
            i += 1
    return customer, charge, scheme


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


def save(customer: Entity, charge: Entity, scheme: Entity):
    conn = pymysql.connect(**db_config)
    cr = conn.cursor()
    # 存客户
    tab = SCHEME_ID + '_' + customer.class_
    id_ = customer.id_
    sql = f'insert into `{tab}`(`id`,'
    values = []
    for pro in customer.pros:
        sql += f'`{pro}`,'
        values.append(customer.pros[pro])
    sql = sql[:-1] + f') values ("{id_}",'
    for v in values:
        sql += f'"{v}",'
    sql = sql[:-1] + ')'
    cr.execute(sql)
    # 存收费方式
    tab = SCHEME_ID + '_' + charge.class_
    id_ = charge.id_
    sql = f'insert into `{tab}`(`id`,'
    values = []
    for pro in charge.pros:
        sql += f'`{pro}`,'
        values.append(charge.pros[pro])
    sql = sql[:-1] + f') values ("{id_}",'
    for v in values:
        sql += f'"{v}",'
    sql = sql[:-1] + ')'
    cr.execute(sql)
    # 存供电方案
    tab = SCHEME_ID + '_' + scheme.class_
    id_ = scheme.id_
    sql = f'insert into `{tab}`(`id`,'
    values = []
    for pro in scheme.pros:
        sql += f'`{pro}`,'
        values.append(scheme.pros[pro])
    sql = sql[:-1] + f') values ("{id_}",'
    for v in values:
        sql += f'"{v}",'
    sql = sql[:-1] + ')'
    cr.execute(sql)
    conn.commit()
    # 关系
    tab = SCHEME_ID + '_' + customer.class_ + '_2_' + charge.class_
    sql = f'insert into {tab} (`id`, `from_id`, `to_id`) values (' \
          f'"{uuid1().hex}", "{customer.id_}", "{charge.id_}")'
    cr.execute(sql)
    tab = SCHEME_ID + '_' + customer.class_ + '_2_' + scheme.class_
    sql = f'insert into {tab} (`id`, `from_id`, `to_id`) values (' \
          f'"{uuid1().hex}", "{customer.id_}", "{scheme.id_}")'
    cr.execute(sql)
    conn.commit()


if __name__ == '__main__':
    file_path = '/Volumes/工作/2021年日常/7-北京业扩报装项目/文档解析/templates/低压供电方案答复单.docx'
    initialize(SCHEME_ID, classes, data_properties, object_properties)
    save(*read_file(file_path))
