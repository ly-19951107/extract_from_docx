# -*- coding: utf-8 -*-
import pymysql
from config import db_config


def initialize(scheme_id: str, classes: dict, data_properties: dict, object_properties: dict):
    """根据本体模型初始化相关表"""
    conn = pymysql.connect(**db_config)
    cr = conn.cursor()
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
