#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel测试用例转XMind脚本
作者: Claude
功能: 将Excel格式的测试用例转换为XMind思维导图格式（新版JSON格式）
"""

import pandas as pd
import os
import sys
import zipfile
import shutil
import json
import uuid
import time
from collections import defaultdict

# 设置标准输出编码为UTF-8
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')


def gen_id():
    """生成UUID"""
    return str(uuid.uuid4())


def get_priority_marker_id(priority):
    """根据优先级返回对应的标记ID"""
    priority_map = {
        'P0': 'priority-1',
        'P1': 'priority-2',
        'P2': 'priority-3',
        'P3': 'priority-4'
    }
    return priority_map.get(str(priority).upper(), None)


def build_topic_content(title, markers=None, children=None):
    """构建主题对象"""
    topic = {
        "id": gen_id(),
        "title": title
    }
    if markers:
        topic["markers"] = markers
    if children:
        topic["children"] = {"attached": children}
    return topic


def convert_excel_to_xmind(excel_path, output_path=None):
    """将Excel测试用例转换为XMind格式"""
    print("[1/5] 读取Excel文件: {}".format(excel_path))

    # 读取Excel文件
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print("[ERROR] 读取Excel文件失败: {}".format(e))
        return False

    print("[2/5] 解析数据结构，共 {} 条测试用例".format(len(df)))

    # 使用列索引来访问数据
    col_idx_level1 = 0
    col_idx_level2 = 1
    col_idx_level3 = 2
    col_idx_scenario = 3
    col_idx_precondition = 4
    col_idx_steps = 5
    col_idx_expected = 6
    col_idx_priority = 7

    # 按层级组织数据
    tree = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))

    for idx, row in df.iterrows():
        try:
            level1 = str(row.iloc[col_idx_level1]).strip() if pd.notna(row.iloc[col_idx_level1]) else '未分类'
            level2 = str(row.iloc[col_idx_level2]).strip() if pd.notna(row.iloc[col_idx_level2]) else '未分类'
            level3 = str(row.iloc[col_idx_level3]).strip() if pd.notna(row.iloc[col_idx_level3]) else '未分类'

            testcase = {
                'scenario': str(row.iloc[col_idx_scenario]).strip() if pd.notna(row.iloc[col_idx_scenario]) else '',
                'precondition': str(row.iloc[col_idx_precondition]).strip() if pd.notna(row.iloc[col_idx_precondition]) else '',
                'steps': str(row.iloc[col_idx_steps]).strip() if pd.notna(row.iloc[col_idx_steps]) else '',
                'expected': str(row.iloc[col_idx_expected]).strip() if pd.notna(row.iloc[col_idx_expected]) else '',
                'priority': str(row.iloc[col_idx_priority]).strip() if pd.notna(row.iloc[col_idx_priority]) else 'P3'
            }

            tree[level1][level2][level3].append(testcase)

        except Exception as e:
            print("[WARNING] 跳过第 {} 行: {}".format(idx + 1, e))
            continue

    print("[3/5] 生成XMind JSON结构...")

    # 构建attached主题列表
    attached_topics = []

    for level1, level2_dict in tree.items():
        # 构建二级模块
        level2_topics = []
        for level2, level3_dict in level2_dict.items():
            # 构建三级模块
            level3_topics = []
            for level3, testcases in level3_dict.items():
                # 构建测试用例
                tc_topics = []
                for i, tc in enumerate(testcases, 1):
                    # 构建测试用例的子节点
                    tc_children = []

                    # 添加前置条件
                    if tc['precondition']:
                        tc_children.append(build_topic_content("前置条件: {}".format(tc['precondition'])))

                    # 添加测试步骤，预期结果作为测试步骤的子节点
                    if tc['steps']:
                        steps_children = []
                        if tc['expected']:
                            steps_children.append(build_topic_content("预期结果: {}".format(tc['expected'])))
                        tc_children.append(build_topic_content("测试步骤: {}".format(tc['steps']), children=steps_children if steps_children else None))

                    # 构建测试用例主题
                    tc_title = "TC{}: {}".format(i, tc['scenario']) if tc['scenario'] else "TC{}".format(i)
                    tc_markers = None
                    marker_id = get_priority_marker_id(tc['priority'])
                    if marker_id:
                        tc_markers = [{"markerId": marker_id}]

                    tc_topic = build_topic_content(tc_title, tc_markers, tc_children if tc_children else None)
                    tc_topics.append(tc_topic)

                # 三级模块主题
                level3_topic = build_topic_content(level3, children=tc_topics)
                level3_topics.append(level3_topic)

            # 二级模块主题
            level2_topic = build_topic_content(level2, children=level3_topics)
            level2_topics.append(level2_topic)

        # 一级模块主题
        level1_topic = build_topic_content(level1, children=level2_topics)
        attached_topics.append(level1_topic)

    # 构建根主题
    root_topic = {
        "id": gen_id(),
        "class": "topic",
        "title": "测试用例",
        "structureClass": "org.xmind.ui.map.clockwise",
        "children": {
            "attached": attached_topics
        }
    }

    # 构建完整的sheet对象
    sheet_id = gen_id()
    content = [
        {
            "id": sheet_id,
            "revisionId": gen_id(),
            "class": "sheet",
            "title": "测试用例",
            "rootTopic": root_topic,
            "topicOverlapping": "overlap",
            "arrangeableLayerOrder": [root_topic["id"]],
            "zones": [],
            "extensions": [
                {
                    "provider": "org.xmind.ui.skeleton.structure.style",
                    "content": {
                        "centralTopic": "org.xmind.ui.map.clockwise",
                        "mainTopic": "org.xmind.ui.logic.right"
                    }
                }
            ],
            "theme": {
                "map": {
                    "id": gen_id(),
                    "properties": {
                        "svg:fill": "#ffffff"
                    }
                }
            }
        }
    ]

    print("[4/5] 创建XMind文件...")

    # 生成输出文件名
    if output_path is None:
        excel_name = os.path.splitext(os.path.basename(excel_path))[0]
        output_path = os.path.join(os.path.dirname(excel_path), "{}_转换.xmind".format(excel_name))

    if not output_path.endswith('.xmind'):
        output_path = output_path + '.xmind'

    # 创建XMind文件
    import tempfile
    temp_dir = tempfile.mkdtemp()

    try:
        # 创建manifest.json
        manifest = {
            "file-entries": [
                {"path": "content.json", "media-type": "application/json"},
                {"path": "metadata.json", "media-type": "application/json"},
                {"path": "manifest.json", "media-type": "application/json"}
            ]
        }
        with open(os.path.join(temp_dir, 'manifest.json'), 'w', encoding='utf-8') as f:
            json.dump(manifest, f, ensure_ascii=False, separators=(',', ':'))

        # 创建content.json
        with open(os.path.join(temp_dir, 'content.json'), 'w', encoding='utf-8') as f:
            json.dump(content, f, ensure_ascii=False, separators=(',', ':'))

        # 创建metadata.json
        timestamp = int(time.time() * 1000)
        metadata = {
            "CreateDate": {"__time__": timestamp}
        }
        with open(os.path.join(temp_dir, 'metadata.json'), 'w', encoding='utf-8') as f:
            json.dump(metadata, f, ensure_ascii=False, separators=(',', ':'))

        # 创建ZIP文件（不压缩content.json以保持兼容性）
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_STORED) as zf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zf.write(file_path, arcname)

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    print("[5/5] 保存完成")

    print("")
    print("[SUCCESS] 转换成功！")
    print("[OUTPUT] 输出文件: {}".format(output_path))
    print("[STATS] 统计信息:")
    print("   - 一级模块: {} 个".format(len(tree)))
    print("   - 测试用例: {} 条".format(len(df)))

    return True


def main():
    """主函数"""
    print("=" * 50)
    print("    Excel测试用例 -> XMind转换工具")
    print("=" * 50)
    print("")

    # 默认文件路径
    default_excel = r"D:\Project\test-case\用户登录测试用例.xlsx"

    # 获取输入文件
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    else:
        excel_path = default_excel

    # 检查文件是否存在
    if not os.path.exists(excel_path):
        print("[ERROR] 文件不存在: {}".format(excel_path))
        print("")
        print("用法: python excel_to_xmind.py [Excel文件路径] [输出路径]")
        return

    # 获取输出文件路径（可选）
    output_path = None
    if len(sys.argv) > 2:
        output_path = sys.argv[2]

    # 执行转换
    convert_excel_to_xmind(excel_path, output_path)


if __name__ == "__main__":
    main()
