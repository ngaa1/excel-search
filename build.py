#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel搜索工具打包脚本
"""

import os
import subprocess
import sys


def run_command(command):
    """运行命令并输出结果"""
    print(f"执行命令: {command}")
    result = subprocess.run(
        command, 
        shell=True, 
        text=True, 
        encoding='utf-8', 
        stdout=subprocess.PIPE, 
        stderr=subprocess.STDOUT
    )
    print(result.stdout)
    if result.returncode != 0:
        print(f"命令执行失败，返回码: {result.returncode}")
        return False
    return True


def main():
    """主函数"""
    print("=== Excel搜索工具打包脚本 ===")
    
    # 安装依赖
    print("\n1. 安装依赖...")
    if not run_command("pip install -r requirements.txt"):
        print("安装依赖失败！")
        return
    
    # 安装PyInstaller
    print("\n2. 安装PyInstaller...")
    if not run_command("pip install pyinstaller"):
        print("安装PyInstaller失败！")
        return
    
    # 开始打包
    print("\n3. 开始打包...")
    if not run_command("pyinstaller --onefile --name excel-search --hidden-import=openpyxl.cell._writer main.py"):
        print("打包失败！")
        return
    
    # 完成
    print("\n4. 打包完成！")
    print("可执行文件位于 dist 目录中")
    print("\n使用方法:")
    print("  excel-search.exe <搜索目录> <搜索关键词> [--filename <文件名关键词>]")
    
    # 等待用户输入
    input("\n按回车键退出...")


if __name__ == "__main__":
    main()