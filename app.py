# -*- coding: utf-8 -*-
from flask import Flask, request, jsonify
import os

app = Flask(__name__)


@app.route('/extract_from_docx', methods=['POST'])
def main():
    """
    参数格式：
    {
        "template": "",  # 指定模板的名称
        "file_path": ""  # 指定文件（文件夹）的路径
    }
    :return:
    """
    args = request.json
    if not args:
        return jsonify({'state': 0, 'msg': '非法的参数（application/json is needed!）'})
    template = args['template']
    file_path = args['file_path']
    if template not in ['HVCPSS', 'CHVPSS', 'LVBERF', 'LVNRERF', 'LVRERF', 'CMEDL', 'HVCERF', 'HVPSSR', 'HVSSS',
                        'LVBEL', 'LVPSSR', 'LVSSS']:
        return jsonify({'state': 0, 'msg': f'不支持的模板名称：{template}'})
    if not os.path.exists(file_path):
        return jsonify({'state': 0, 'msg': f'非法的路径：{file_path}'})
    extract(template, file_path)
    return jsonify({'state': 1, 'msg': 'success'})


def extract(template_name, file_path):
    if os.path.isfile(file_path):
        _extract(template_name, file_path)
    else:
        files = os.listdir(file_path)
        for file in files:
            if file.endswith('.docx'):
                abs_file_path = os.path.join(file_path, file)
                _extract(template_name, abs_file_path)


def _extract(template_name, file_path):
    pass
