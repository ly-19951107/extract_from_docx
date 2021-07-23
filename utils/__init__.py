# -*- coding: utf-8 -*-
import pymysql
from config import db_config

conn = pymysql.connect(**db_config)
cr = conn.cursor()


def initialize(scheme_id: str, classes: dict, data_properties: dict, object_properties: dict):
    """根据本体模型初始化相关表"""
    for _class in classes:
        table_name = scheme_id + '_' + _class
        fields = ['id']
        comments = ['唯一标识']
        for pro in data_properties:
            if data_properties[pro]['domain'] == _class:
                fields.append(pro)
                comments.append(data_properties[pro]['desc'])
        sql = f"create table if not exists `{table_name}`("
        for i in range(len(fields)):
            sql += f"`{fields[i]}` varchar(255) comment '{comments[i]}',"
        sql = sql[:-1]
        sql += ')'
        cr.execute(sql)
    conn.commit()
    for i in object_properties:
        rel = object_properties[i]
        rel_tab = scheme_id + '_' + rel['domain'] + '_2_' + rel['range']
        sql = f"""create table if not exists `{rel_tab}`(
                `id` varchar(255) primary key,
                `from_id` varchar(255),
                `to_id` varchar(255),
                `rel_name` varchar (10) default '{rel["name"]}'
                )
            """
        cr.execute(sql)
    conn.commit()


def std_rel(scheme_id: str, class_std: dict):
    object_properties1 = {}
    class_std_id = {}
    obj_per_list = []
    num = 0
    for i in class_std:
        if i not in class_std_id.keys():
            class_std_id[i] = {}
        cr.execute("update bz_tab set BZ_name = replace(BZ_name,' ','')")
        cr.execute("update bz_1_info set BZ_first_title = replace(BZ_first_title,' ','')")
        cr.execute("update bz_2_info set BZ_second_title = replace(BZ_second_title,' ','')")
        for j in class_std[i]:
            j = j.replace(' ', '')
            object_properties1[num] = {'domain': i, 'name': 'reference', 'ZH_name': '参考', 'desc': '描述参考的哪个标准'}
            cr.execute(f"select id from bz_tab where BZ_name = '{j}'")
            res = cr.fetchone()
            if res is not None:
                object_properties1[num]['range'] = 'bz_tab'
                if 'bz_tab' not in class_std_id[i]:
                    class_std_id[i]['bz_tab'] = []
                class_std_id[i]['bz_tab'].append(res[0])
            else:
                cr.execute(f"select id from bz_1_info where BZ_first_title = '{j}'")
                res = cr.fetchone()
                if res is not None:
                    object_properties1[num]['range'] = 'bz_1_info'
                    if 'bz_1_info' not in class_std_id[i]:
                        class_std_id[i]['bz_1_info'] = []
                    class_std_id[i]['bz_1_info'].append(res[0])
                else:
                    cr.execute(f"select id from bz_2_info where BZ_second_title = '{j}'")
                    res = cr.fetchone()
                    if res is not None:
                        object_properties1[num]['range'] = 'bz_2_info'
                        if 'bz_2_info' not in class_std_id[i]:
                            class_std_id[i]['bz_2_info'] = []
                        class_std_id[i]['bz_2_info'].append(res[0])
                    else:
                        print(f'未找到对应标准{j}，请检查对应标准是否与数据库存储值一致')
                        continue
            num += 1
    for i in list(object_properties1.keys()):
        if object_properties1[i] not in obj_per_list:
            obj_per_list.append(object_properties1[i])
        else:
            del object_properties1[i]
    for i in object_properties1:
        rel = object_properties1[i]
        rel_tab = scheme_id + '_' + rel['domain'] + '_2_' + rel['range']
        sql = f"""create table if not exists `{rel_tab}`(
                `id` varchar(255) primary key,
                `from_id` varchar(255),
                `to_id` varchar(255),
                `rel_name` varchar (10) default '{rel["name"]}'
                )
            """
        cr.execute(sql)
    conn.commit()
    return object_properties1, class_std_id
