# -*- coding: utf-8 -*-
"""用于解析「居民小区高压供电方案」文档

一个文档会提取出
"""
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
import re
import pymysql
from uuid import uuid1
from config import db_config
from utils import initialize

SCHEME_ID = 'CHVPSS'

# <<<<<配置区域
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
        "keys": ["客户名称"],
        "pros": ["name"],
        "class": 'customer',
        "class_ZH": '用户',
        "match_once": True,
    },
    {
        "rule_no": 2,
        "location_rule": r'根据客户提供的小区建设规划.*公建配套用房.*平方米',
        "keys": ["建筑面积", "建筑设计", "住宅面积", "住户数量", "商业用房面积", "公建配套用房面积"],
        "pros": ["building_area", "building_design", "residence_area", "householder_num", "commercial_building_area",
                 "public_building_area"],
        "class": 'community',
        "class_ZH": '小区',
        "match_once": True,
    },
    {
        "rule_no": 3,
        "location_rule": r'.*经计算用电负荷.*二级负荷.*千瓦',
        "keys": ["计算负荷", "供用电容量", "一级负荷", "二级负荷"],
        "pros": ["cal_load", "supply_cons_cap", "first_load", "second_load"],
        "class": 'power_supply_cap',
        "class_ZH": '供电容量',
        "match_once": True,
    },
    {
        "rule_no": 4,
        "location_rule": r"根据供电条件和小区用电需求.*电压等级。",
        "keys": ["供电电源类型", "主供电源电压等级", "备用电源电压等级"],
        "pros": ['power_source_type', 'main_volt', 'standby_volt'],
        "class": 'power_supply_mode',
        "class_ZH": '供电方式',
        "match_once": True,
    },
    {
        "rule_no": 5,
        "location_rule": r"主供电源.*母线的.*供电线路.*线路参数.*与公配线路.*",
        "keys": ["主供电源变电所", "开关", "接线方式", "敷设方式", "线路参数", "供电容量", "接点设备"],
        "pros": ['power_source_no', 'main_or_standby', 'subs', 'switch', 'conn_mode', 'lay_mode',
                 'line_para', 'line_supply_cap', 'contact_device'],
        "class": 'power_source',
        "class_ZH": '供电电源',
        "match_once": False,
    },
    {
        "rule_no": 6,
        "location_rule": r"备用电源.*母线的.*供电线路.*线路参数.*与公配线路.*",
        "keys": ["备用电源变电所", "母线开关", "接线方式", "敷设方式", "线路参数", "供电容量", "接点设备"],
        "pros": ['power_source_no', 'main_or_standby', 'subs', 'switch', 'conn_mode', 'lay_mode', 'line_para',
                 'line_supply_cap', 'contact_device'],
        "class": 'power_source',
        "class_ZH": '供电电源',
        "match_once": True,
    },
    {
        "rule_no": 7,
        "location_rule": r"该小区采用.*受电总容量.*千伏安",
        "keys": ["供电方式", "开闭所数量", "配电站数量", "受电总容量"],
        "pros": ['supply_mode', 'open_close_station_num', 'power_station_num', 'total_cap'],
        "class": 'community',
        "class_ZH": "小区",
        "match_once": True,
    },
    {
        "rule_no": 8,
        "location_rule": r"采用.*设进线柜.*台",
        "keys": ["主接线方式", "进线柜数量", "PT柜数量", "馈电柜数量", "联络柜数量", "其它数量"],
        "pros": ["main_line_type", "in_line_cabinet_num", "PT_cabinet_num", "feed_cabinet_num", "contact_cabinet_num",
                 "other_cabinet_num"],
        'class': 'receive_point',
        "class_ZH": "受电点",
        "match_once": True,
    },
    {
        "rule_no": 9,
        "location_rule": r"配电站.*高压配电装置.*",
        "keys": ["配电站编号", "变压器类型", "变压器数量", "单台变压器容量", "供电范围", "高压配电装置", "低压配电装置"],
        "pros": ["power_station_no", "trans_type", "trans_num", "single_trans_cap", "supply_range",
                 "high_vol_power_device", "low_vol_power_device"],
        'class': 'power_station',
        "class_ZH": "配电站",
        "match_once": False,
    },
    {
        "rule_no": 10,
        "location_rule": r"用电人一、二级负荷.*自备保安容量.*千伏安(千瓦)",
        "keys": ['客户自备保安容量', '客户自备保安容量'],
        "pros": ['security_cap', 'security_cap'],
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
        "location_rule": r'客户的用电类别分别.*',
        "keys": ['用电类别', '用电类别'],
        'pros': ['elec_type', 'elec_type'],
        "class": 'customer',
        "class_ZH": "用户",
        "match_once": True,
    },
    {
        "rule_no": 13,
        "location_rule": r'计量点.*用于计量用电.*电压互感.*',
        "keys": ['计量点编号', "用电量类别", "计量装置位置", "计量方式", "接线方式", "电能表规格", "精度",
                 "电压互感器规格", "精度", "电流互感器规格", "精度", "电量采集系统"],
        'pros': ['meter_point_no', 'point_elec_type', 'position', 'meter_type', 'meter_line_type',
                 'meter_specs', 'precision', 'volt_trans', 'volt_pre',
                 'cur_trans', 'cur_pre', 'acquisition'],
        "class": 'meter_point',
        "class_ZH": "计量点",
        "match_once": False,
    },
    {
        "rule_no": 14,
        "location_rule": r'.*根据客户的用电分类.*',
        "keys": ["收费方式", "电价类别", "电价类别", "电价类别"],
        "pros": ['method', 'elec_price_type', 'elec_price_type', 'elec_price_type'],
        "class": 'charge',
        "class_ZH": "收费",
        "match_once": True,
    },
    {
        "rule_no": 15,
        "location_rule": r'本方案有效期自.*',
        "keys": ['开始年', '开始月', '开始日', '结束年', '结束月', '结束日', '有效期'],
        'pros': ['term_start', 'term_start', 'term_start',
                 'term_end', 'term_end', 'term_end', 'validity_term'],
        'class': 'com_high_volt_power_supply_schema',
        "class_ZH": "居民小区高压供电方案",
        "match_once": True,
    }
]

classes = {
    'com_high_volt_power_supply_schema': '居民小区高压供电方案',
    "customer": "用户",
    "community": "小区",
    "power_supply_cap": "供电容量",
    "power_supply_mode": "供电方式",
    "power_source": "供电电源",
    "receive_point": "受电点",
    "power_station": "配电站",
    "meter_point": "计量点",
    "charge": "收费"
}
data_properties = {
    'validity_term': {'domain': 'com_high_volt_power_supply_schema', 'range': 'string', 'desc': '有效期'},
    'term_start': {'domain': 'com_high_volt_power_supply_schema', 'range': 'string', 'desc': '开始有效时间'},
    'term_end': {'domain': 'com_high_volt_power_supply_schema', 'range': 'string', 'desc': '结束有效时间'},

    "name": {"domain": "customer", "range": "string", "desc": "用户名称"},
    # "type": {"domain": "customer", "range": "string", "desc": "用户类型"},
    # "total_cap": {"domain": "customer", "range": "string", "desc": "用电总用量"},
    # "elec_demand": {"domain": "customer", "range": "string", "desc": "用电需求"},
    "elec_type": {"domain": "customer", "range": "string", "desc": "用电类别"},
    # "receive_point_num": {"domain": "customer", "range": "string", "desc": "受电点数量"},

    "building_area": {"domain": "community", "range": "string", "desc": "建筑面积"},
    "building_design": {"domain": "community", "range": "string", "desc": "建筑设计"},
    "residence_area": {"domain": "community", "range": "string", "desc": "住宅面积"},
    "householder_num": {"domain": "community", "range": "string", "desc": "住户数量"},
    "commercial_building_area": {"domain": "community", "range": "string", "desc": "商业用房面积"},
    "public_building_area": {"domain": "community", "range": "string", "desc": "公建配套用房面积"},
    "supply_mode": {"domain": "community", "range": "string", "desc": "供电方式"},
    "open_close_station_num": {"domain": "community", "range": "string", "desc": "开闭所数量"},
    "power_station_num": {"domain": "community", "range": "string", "desc": "配电站数量"},
    "total_cap": {"domain": "community", "range": "string", "desc": "受电总容量"},

    "cal_load": {"domain": "power_supply_cap", "range": "string", "desc": "计算负荷"},
    "supply_cons_cap": {"domain": "power_supply_cap", "range": "string", "desc": "供用电容量"},
    "first_load": {"domain": "power_supply_cap", "range": "string", "desc": "一级负荷"},
    "second_load": {"domain": "power_supply_cap", "range": "string", "desc": "二级负荷"},

    "power_source_type": {"domain": "power_supply_mode", "range": "string", "desc": "供电电源类型"},
    "main_volt": {"domain": "power_supply_mode", "range": "string", "desc": "主供电源电压等级"},
    "standby_volt": {"domain": "power_supply_mode", "range": "string", "desc": "备用电源电压等级"},

    "power_source_no": {"domain": "power_source", "range": "string", "desc": "主电源"},
    "main_or_standby": {"domain": "power_source", "range": "string", "desc": "主供电源还是备用电源"},
    "volt": {"domain": "power_source", "range": "string", "desc": "电压等级"},
    "subs": {"domain": "power_source", "range": "string", "desc": "变电所"},
    "line": {"domain": "power_source", "range": "string", "desc": "母线"},
    "switch": {"domain": "power_source", "range": "string", "desc": "开关"},
    "conn_mode": {"domain": "power_source", "range": "string", "desc": "接线方式"},
    "lay_mode": {"domain": "power_source", "range": "string", "desc": "供电线路敷设方式"},
    "line_para": {"domain": "power_source", "range": "string", "desc": "线路型号与参数"},
    "line_supply_cap": {"domain": "power_source", "range": "string", "desc": "线路供电容量"},
    "contact_device": {"domain": "power_source", "range": "string", "desc": "线路供电容量"},

    "main_line_type": {"domain": "receive_point", 'range': 'string', 'desc': '主接线方式'},
    "in_line_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': '进线柜数量'},
    "PT_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': 'PT柜数量'},
    "feed_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': '馈电柜数量'},
    "contact_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': '联络柜数量'},
    "other_cabinet_num": {"domain": "receive_point", 'range': 'string', 'desc': '其他数量'},
    "security_cap": {"domain": "receive_point", 'range': 'string', 'desc': '客户自备保安容量'},
    "run_mode": {"domain": "receive_point", 'range': 'string', 'desc': '运行方式'},

    "power_station_no": {"domain": "power_station", 'range': 'string', 'desc': '配电站编号'},
    "trans_type": {"domain": "power_station", 'range': 'string', 'desc': '变压器类型'},
    "trans_num": {"domain": "power_station", 'range': 'string', 'desc': '变压器数量'},
    "single_trans_cap": {"domain": "power_station", 'range': 'string', 'desc': '单台变压器容量'},
    "supply_range": {"domain": "power_station", 'range': 'string', 'desc': '供电范围'},
    "high_vol_power_device": {"domain": "power_station", 'range': 'string', 'desc': '高压配电装置'},
    "low_vol_power_device": {"domain": "power_station", 'range': 'string', 'desc': '高压配电装置'},

    "meter_point_no": {"domain": "meter_point", 'range': 'string', 'desc': '计量点编号'},
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
    "elec_price_type": {"domain": "charge", 'range': 'string', 'desc': '电价类别'}
}


object_properties = {
    0: {
        'domain': 'com_high_volt_power_supply_schema',
        'range': 'customer',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述方案属于哪个客户',
    },
    1: {
        'domain': 'customer',
        'range': 'community',
        'name': 'BelongsTo',
        'ZH_name': '属于',
        'desc': '描述客户与小区之间的关系',
    },
    2: {
        'domain': 'com_high_volt_power_supply_schema',
        'range': 'power_supply_cap',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的供电容量之间的关系',
    },
    3: {
        'domain': 'com_high_volt_power_supply_schema',
        'range': 'power_supply_mode',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的供电方式之间的关系',
    },
    4: {
        'domain': 'com_high_volt_power_supply_schema',
        'range': 'power_source',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的供电电源之间的关系',
    },
    5: {
        'domain': 'com_high_volt_power_supply_schema',
        'range': 'receive_point',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的受电点之间的关系',
    },
    6: {
        'domain': 'com_high_volt_power_supply_schema',
        'range': 'meter_point',
        'name': 'Untitled',
        'ZH_name': '',
        'desc': '描述方案与其记录的计量点之间的关系',
    },
    7: {
        'domain': 'com_high_volt_power_supply_schema',
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
                if rule_no == 5:
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
                elif rule_no == 13:
                    values = handle_13(p)
                    entity = Entity(class_, uuid1().hex)
                    for j in range(len(pros)):
                        pro = pros[j]
                        value = values[j]
                        entity.add_pro(pro, value)
                    if class_ in entity_dict:
                        entity_dict[class_].append(entity)
                    else:
                        entity_dict[class_] = [entity]
                elif rule_no == 6:
                    values = handle_6(p)
                    entity = Entity(class_, uuid1().hex)
                    for j in range(len(pros)):
                        pro = pros[j]
                        value = values[j]
                        entity.add_pro(pro, value)
                    if class_ in entity_dict:
                        entity_dict[class_].append(entity)
                    else:
                        entity_dict[class_] = [entity]
                elif rule_no == 9:
                    values = handle_9(p)
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
                        if rule_no == 3:
                            pro = pros[j]
                            value = values[j + 6]
                            entity.add_pro(pro, value)
                        else:
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


def handle_5(p):
    """处理规则5"""
    src_no = re.compile(r'主供电源(.)为')
    power_source_no = src_no.match(p.text).groups()[0]
    main_or_standby = '主供电源'
    values = cluster_underline(p.runs)
    return [power_source_no, main_or_standby] + values


def handle_6(p):
    """处理规则6"""
    main_or_standby = '备用电源'
    values = cluster_underline(p.runs)
    return ['', main_or_standby] + values


def handle_9(p):
    """处理规则9"""
    receive_point_no = re.compile(r'配电站(.)配置')
    receive_point_no = receive_point_no.match(p.text).groups()[0]
    values = cluster_underline(p.runs)
    return [receive_point_no] + values


def handle_13(p):
    mer_no = re.compile(r'计量点(.)：用')
    meter_no = mer_no.match(p.text).groups()[0]
    values = cluster_underline(p.runs)
    return [meter_no] + values


if __name__ == '__main__':
    file_path = r'C:\Users\liyang\Desktop\extract\extract_from_docx\templates\居民小区高压供电方案.docx'
    initialize(SCHEME_ID, classes, data_properties, object_properties)
    save(read_file(file_path))
