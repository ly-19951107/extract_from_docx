# -*- coding: utf-8 -*-
"""用于解析「高压客户供电方案」文档
"""
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
import re
import pymysql
from uuid import uuid1
from config import db_config
from utils import initialize

SCHEME_ID = 'HVCPSS'

rules = [
    # 每个字典为一个针对特定段落的规则，其中
    # rule_no为一个数字，指当前规则的标识
    # location_rule对应一个用于定位段落的正则表达式；
    # keys为一个列表，里面按照顺序指定了该段落中各个下划线的值所对应的属性名
    # pros为一个列表，里面按照顺序指定了keys中每个属性名所属的实体属性
    # class为一个字符串，指当前提取的属性属于哪一个概念
    # class_ZH为一个字符串，指class的中文描述
    # match_once的值为一个布尔型，代表当前规则是否只在文件中匹配一次
    {
        "rule_no": 1,
        "location_rule": r'根据.*确定供电方案如下',
        "keys": ["客户名称", "用电设备总容量"],
        "pros": ["name", "total_cap"],
        "class": 'customer',
        "class_ZH": '用户',
        "match_once": True,
    },
    {
        "rule_no": 2,
        "location_rule": r'根据客户提供的用电设备技术参数.*千瓦',
        "keys": ["计算负荷", "供用电容量", "一级负荷", "二级负荷"],
        "pros": ["cal_load", "supply_cons_cap", "first_load", "second_load"],
        "class": 'power_supply_cap',
        "class_ZH": '供电容量',
        "match_once": True,
    },
    {
        "rule_no": 3,
        "location_rule": r"根据供电条件和客户用电需求.*电压等级。",
        "keys": ["主供电源电压等级", "备用电源电压等级"],
        "pros": ['main_volt', 'standby_volt'],
        "class": 'power_supply_mode',
        "class_ZH": '供电方式',
        "match_once": True,
    },
    {
        "rule_no": 4,
        "location_rule": r"主供电源.*母线的.*供电线路.*线路的型号与参数.*",
        "keys": ["主供电源变电所", "开关", "接线方式", "敷设方式", "线路的型号与参数", "供电容量"],
        "pros": ['power_source_no', 'main_or_standby', 'subs', 'switch', 'conn_mode', 'lay_mode',
                 'line_type_para', 'line_supply_cap'],
        "class": 'power_source',
        "class_ZH": '供电电源',
        "match_once": False,
    },
    {
        "rule_no": 5,
        "location_rule": r"备用电源.*母线的.*供电线路.*线路的型号与参数.*",
        "keys": ["备用电源变电所", "母线开关", "接线方式", "敷设方式", "线路的型号与参数", "供电容量"],
        "pros": ['power_source_no', 'main_or_standby', 'subs', 'switch', 'conn_mode', 'lay_mode',
                 'line_type_para', 'line_supply_cap'],
        "class": 'power_source',
        "class_ZH": '供电电源',
        "match_once": True,
    },
    {
        "rule_no": 6,
        "location_rule": r"用电人.*用电总容量.*千伏安",
        "keys": ["受电点数量", '用电总容量'],
        "pros": ['receive_point_num', 'total_cap'],
        "class": 'customer',
        "class_ZH": "用户",
        "match_once": True,
    },
    # {
    #     "rule_no": 7,
    #     "location_rule": r"受电点.*变压器.*千伏安变压器.*台",
    #     "keys": ["变压器类型", "容量", "数量"],
    #     "pros": ['trans_type', 'trans_num'],
    #     'class': 'receive_point',
    #     "class_ZH": "受电点",
    #     "match_once": True,
    # },
    {
        "rule_no": 8,
        "location_rule": r"用电人一、二级负荷.*千伏安(千瓦)。",
        "keys": ["客户自备保安容量", "应急保安容量"],
        "pros": ["security_cap", "emerge_security_cap"],
        'class': 'receive_point',
        "class_ZH": "受电点",
        "match_once": True,
    },
    {
        "rule_no": 9,
        "location_rule": r"高压部分配置：.*",
        "keys": ['进线柜数量 ', '计量柜数量', 'PT柜数量', "馈电柜数量", "电容柜数量", '其它数量'],
        "pros": ['in_line_cabinet_num', 'meter_cabinet_num', 'PT_cabinet_num',
                 'feed_cabinet_num', 'capacitor_cabinet_num', 'other_cabinet_num'],
        'class': 'receive_point',
        "class_ZH": "受电点",
        "match_once": True,
    },
    {
        "rule_no": 10,
        "location_rule": r"受电变压器电源侧.*",
        "keys": ["接线方式", "控制设备"],
        "pros": ['line_type', 'control_equip'],
        'class': 'receive_point',
        "class_ZH": "受电点",
        "match_once": True,
    },
    {
        "rule_no": 11,
        "location_rule": r'.*运行方式.*',
        "keys": ['运行方式'],
        'pros': ['run_mode'],
        'class': 'receive_point',
        "class_ZH": "受电点",
        "match_once": True,
    },
    {
        "rule_no": 12,
        "location_rule": r'.*产权及维护责任分界点划分.*',
        "keys": ['产权及维护责任分界点划分'],
        'pros': ['demarcation_division'],
        'class': 'receive_point',
        "class_ZH": "受电点",
        "match_once": True,
    },
    {
        "rule_no": 13,
        "location_rule": r'客户的用电类别分别.*',
        "keys": ['用电类别', '用电类别'],
        'pros': ['elec_type', 'elec_type'],
        "class": 'customer',
        "class_ZH": "用户",
        "match_once": True,
    },
    {
        "rule_no": 14,
        "location_rule": r'计量点.*用于计量用电.*电压互感.*',
        "keys": ['计量点编号', "用电量类别", "计量装置位置", "计量方式", "接线方式", "电能表规格", "精度",
                 "电压互感器规则", "精度", "电流互感器规格", "精度", "电量采集系统"],
        'pros': ['meter_no', 'point_elec_type', 'position', 'meter_type', 'meter_line_type',
                 'meter_specs', 'precision', 'volt_trans', 'volt_pre',
                 'cur_trans', 'cur_pre', 'acquisition'],
        "class": 'meter_point',
        "class_ZH": "计量点",
        "match_once": False,
    },
    {
        "rule_no": 15,
        "location_rule": r'.*根据客户的用电分类.*',
        "keys": ["收费方式", "电价类别", "电价类别", "电价类别", "电价类别"],
        "pros": ['method', ],
        "class": 'charge',
        "class_ZH": "收费",
        "match_once": True,
    },
    {
        "rule_no": 16,
        "location_rule": r"根据用电人用电性质应执行.*",
        "keys": ["功率因数考核标准", "配制下限"],
        'pros': ['power_factor', 'cap_inf'],
        "class": 'charge',
        "class_ZH": "收费",
        "match_once": True,
    },
    {
        "rule_no": 17,
        "location_rule": r".*根据相关规定.*双（多）电源客户应",
        "keys": ["高可靠性供电费"],
        'pros': ['HA_charge'],
        "class": 'charge',
        "class_ZH": "收费",
        "match_once": True,
    },
    {
        "rule_no": 18,
        "location_rule": r'.*根据相关规定，临时施工用电的.*',
        "keys": ['临时接电费'],
        'pros': ['tmp_charge'],
        "class": 'charge',
        "class_ZH": "收费",
        "match_once": True,
    },
    {
        "rule_no": 19,
        "location_rule": r'本方案有效期自.*',
        "keys": ['开始年', '开始月', '开始日', '结束年', '结束月', '结束日', '有效期'],
        'pros': ['term_start', 'term_start', 'term_start',
                 'term_end', 'term_end', 'term_end', 'validity_term'],
        'class': 'high_volt_cus_power_supply_schema',
        "class_ZH": "高压客户供电方案",
        "match_once": True,
    }
]

classes = {
    'high_volt_cus_power_supply_schema': '高压客户供电方案',
    "customer": "用户",
    "power_supply_cap": "供电容量",
    "power_supply_mode": "供电方式",
    "power_source": "供电电源",
    "receive_point": "受电点",
    "meter_point": "计量点",
    "charge": "收费",
}
data_properties = {
    'validity_term': {'domain': 'high_volt_cus_power_supply_schema', 'range': 'string', 'desc': '有效期'},
    'term_start': {'domain': 'high_volt_cus_power_supply_schema', 'range': 'string', 'desc': '开始有效时间'},
    'term_end': {'domain': 'high_volt_cus_power_supply_schema', 'range': 'string', 'desc': '结束有效时间'},

    "name": {"domain": "customer", "range": "string", "desc": "用户名称"},
    "type": {"domain": "customer", "range": "string", "desc": "用户类型"},
    "total_cap": {"domain": "customer", "range": "string", "desc": "用电总用量"},
    "elec_demand": {"domain": "customer", "range": "string", "desc": "用电需求"},
    "elec_type": {"domain": "customer", "range": "string", "desc": "用电类别"},
    "receive_point_num": {"domain": "customer", "range": "string", "desc": "受电点数量"},

    "cal_load": {"domain": "power_supply_cap", "range": "string", "desc": "计算负荷"},
    "supply_cons_cap": {"domain": "power_supply_cap", "range": "string", "desc": "供用电容量"},
    "first_load": {"domain": "power_supply_cap", "range": "string", "desc": "一级负荷"},
    "second_load": {"domain": "power_supply_cap", "range": "string", "desc": "二级负荷"},

    "power_num": {"domain": "power_supply_mode", "range": "string", "desc": "供电电源数量"},
    "main_volt": {"domain": "power_supply_mode", "range": "string", "desc": "主供电源电压等级"},
    "standby_volt": {"domain": "power_supply_mode", "range": "string", "desc": "备用电源电压等级"},

    "power_source_no": {"domain": "power_source", "range": "string", "desc": "电源编号"},
    "main_or_standby": {"domain": "power_source", "range": "string", "desc": "主供电源还是备用电源"},
    "volt": {"domain": "power_source", "range": "string", "desc": "电压等级"},
    "subs": {"domain": "power_source", "range": "string", "desc": "变电所"},
    "line": {"domain": "power_source", "range": "string", "desc": "母线"},
    "switch": {"domain": "power_source", "range": "string", "desc": "开关"},
    "conn_mode": {"domain": "power_source", "range": "string", "desc": "接线方式"},
    "lay_mode": {"domain": "power_source", "range": "string", "desc": "供电线路敷设方式"},
    "line_type_para": {"domain": "power_source", "range": "string", "desc": "线路型号与参数"},
    "line_supply_cap": {"domain": "power_source", "range": "string", "desc": "线路供电容量"},

    # "trans_type": {"domain": "receive_point", 'range': 'string', 'desc': '变压器类型'},
    # "trans_num": {"domain": "receive_point", 'range': 'string', 'desc': '变压器数量'},
    "security_cap": {"domain": "receive_point", 'range': 'string', 'desc': '客户自备保安容量'},
    "emerge_security_cap": {"domain": "receive_point", 'range': 'string', 'desc': '自备应急保安容量'},
    "in_line_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': '进线柜数量'},
    "meter_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': '计量柜数量'},
    "PT_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': 'PT柜数量'},
    "feed_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': '馈电柜数量'},
    "capacitor_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': '电容柜数量'},
    "other_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': '其他数量'},
    "line_type": {"domain": "receive_point", 'range': 'string', 'desc': '接线方式'},
    "control_equip": {"domain": "receive_point", 'range': 'string', 'desc': '控制设备'},
    "run_mode": {"domain": "receive_point", 'range': 'string', 'desc': '运行方式'},
    "demarcation_division": {"domain": "receive_point", 'range': 'string', 'desc': '分界点划分'},

    "meter_no": {"domain": "meter_point", 'range': 'string', 'desc': '计量点编号'},
    "point_elec_type": {"domain": "meter_point", 'range': 'string', 'desc': '用电类别'},
    "position": {"domain": "meter_point", 'range': 'string', 'desc': '计量装置位置'},
    "meter_type": {"domain": "meter_point", 'range': 'string', 'desc': '计量方式'},
    "meter_line_type": {"domain": "meter_point", 'range': 'string', 'desc': '接线方式'},
    "meter_specs": {"domain": "meter_point", 'range': 'string', 'desc': '电能表规格'},
    "precision": {"domain": "meter_point", 'range': 'string', 'desc': '精度'},
    "volt_trans": {"domain": "meter_point", 'range': 'string', 'desc': '电压互感器规格'},
    "volt_pre": {"domain": "meter_point", 'range': 'string', 'desc': '电压互感器精度'},
    "cur_trans": {"domain": "meter_point", 'range': 'string', 'desc': '电流互感器规格'},
    "cur_pre": {"domain": "meter_point", 'range': 'string', 'desc': '电流互感器精度'},
    "acquisition": {"domain": "meter_point", 'range': 'string', 'desc': '电量采集系统'},

    "method": {"domain": "charge", 'range': 'string', 'desc': '收费方式'},
    "power_factor": {"domain": "charge", 'range': 'string', 'desc': '功率因数考核标准'},
    "cap_inf": {"domain": "charge", 'range': 'string', 'desc': '总容量下界'},
    "HA_charge": {"domain": "charge", 'range': 'string', 'desc': '高可靠性供电费'},
    "tmp_charge": {"domain": "charge", 'range': 'string', 'desc': '临时接电费'},

}
object_properties = {
    0: {
        'domain': 'high_volt_cus_power_supply_schema',
        'range': 'customer',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述方案属于哪个客户',
    },
    1: {
        'domain': 'high_volt_cus_power_supply_schema',
        'range': 'power_supply_cap',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的供电容量之间的关系',
    },
    2: {
        'domain': 'high_volt_cus_power_supply_schema',
        'range': 'power_supply_mode',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的供电方式之间的关系',
    },
    3: {
        'domain': 'high_volt_cus_power_supply_schema',
        'range': 'power_source',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的供电电源之间的关系',
    },
    4: {
        'domain': 'high_volt_cus_power_supply_schema',
        'range': 'receive_point',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的受电点之间的关系',
    },
    5: {
        'domain': 'high_volt_cus_power_supply_schema',
        'range': 'meter_point',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的计量点之间的关系',
    },
    6: {
        'domain': 'high_volt_cus_power_supply_schema',
        'range': 'charge',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的收费方式之间的关系',
    }
}


def read_file(file_path):
    """读取一个docx文件"""
    try:
        docx = Document(file_path)
    except PackageNotFoundError:
        print(f'路径不正确或目标为加密文档：{file_path}')
        return
    paragraphs = docx.paragraphs
    entity_dict = {}
    for i in range(len(rules)):
        rule = rules[i]
        rule_no = rule['rule_no']
        # 根据rule_no的不同，要做不同的处理
        location_rule = re.compile(rule['location_rule'])
        match_once = rule['match_once']
        class_ = rule['class']
        pros = rule['pros']
        for p in paragraphs:
            if location_rule.match(p.text):
                if rule_no == 4:
                    values = handle_4(p)
                    entity = Entity(class_, uuid1().hex)
                    for j in range(len(pros)):
                        pro = pros[j]
                        value = values[j]
                        entity.add_pro(pro, value)
                    if class_ in entity_dict:
                        entity_dict[class_].append(entity)
                    else:
                        entity_dict[class_] = [entity]
                elif rule_no == 14:
                    values = handle_14(p)
                    entity = Entity(class_, uuid1().hex)
                    for j in range(len(pros)):
                        pro = pros[j]
                        value = values[j]
                        entity.add_pro(pro, value)
                    if class_ in entity_dict:
                        entity_dict[class_].append(entity)
                    else:
                        entity_dict[class_] = [entity]
                elif rule_no == 5:
                    values = handle_5(p)
                    entity = Entity(class_, uuid1().hex)
                    for j in range(len(pros)):
                        pro = pros[j]
                        value = values[j]
                        entity.add_pro(pro, value)
                    if class_ in entity_dict:
                        entity_dict[class_].append(entity)
                    else:
                        entity_dict[class_] = [entity]
                else:
                    values = cluster_underline(p.runs)
                    if class_ not in entity_dict:
                        entity = Entity(class_, uuid1().hex)
                        entity_dict[class_] = entity
                    else:
                        entity = entity_dict[class_]
                    for j in range(len(pros)):
                        pro = pros[j]
                        value = values[j]
                        entity.add_pro(pro, value)
                if match_once:
                    break
            else:
                continue
    return entity_dict


def cluster_underline(runs):
    """对一个段落的runs按照下划线进行聚合"""
    i = 0
    texts = []
    while i < len(runs):
        run = runs[i]
        if not run.underline:
            i += 1
            continue
        else:
            text = run.text.strip()
            if i == len(runs) - 1:
                texts.append(text)
                i += 1
                continue
            for j in range(i + 1, len(runs)):
                if not runs[j].underline:
                    i = j
                    break
                else:
                    text += runs[j].text.strip()
            texts.append(text)
    return texts


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


def handle_4(p):
    """处理规则4"""
    src_no = re.compile(r'主供电源(.)为')
    power_source_no = src_no.match(p.text).groups()[0]
    main_or_standby = '主供电源'
    values = cluster_underline(p.runs)
    return [power_source_no, main_or_standby] + values


def handle_5(p):
    """处理规则5"""
    main_or_standby = '备用电源'
    values = cluster_underline(p.runs)
    return ['', main_or_standby] + values


def handle_14(p):
    mer_no = re.compile(r'计量点(.)：用')
    meter_no = mer_no.match(p.text).groups()[0]
    values = cluster_underline(p.runs)
    return [meter_no] + values


if __name__ == '__main__':
    file_path = '/templates/高压客户供电方案.docx'
    initialize(SCHEME_ID, classes, data_properties, object_properties)
    save(read_file(file_path))
