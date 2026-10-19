#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
地址坐标转换工具 v2.8.66（模块化版）
修复：结构化地址要素列在「跳过已有结果/只重试异常行/断点续跑」路径下缺补写，
导致勾选后只有表头没有数据
入口：业务代码在 src/ 包内。直接运行本文件即可。
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from src.gui import main

if __name__ == "__main__":
    main()
