#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import PyInstaller.__main__
import os
import shutil
from pathlib import Path

def package_program():
        if os.path.exists("dist"):
            shutil.rmtree("dist")
        if os.path.exists("build"):
            shutil.rmtree("build")
        
        args = [
            "通用接龙数据报表生成器.py",
            "--onedir",  # 改为目录模式，避免解压耗时
            "--windowed",
            "--name=DragonReportGenerator",
            "--add-data=存量业务配置.json;.",
            "--add-data=存量经理配置.json;.",
            "--add-data=渠道厅店配置.json;.",
            "--clean",
            "--noconfirm"
        ]
        
        print("Starting packaging...")
        try:
            PyInstaller.__main__.run(args)
            print("Packaging completed! Executable located at dist/DragonReportGenerator.exe")
            print("Please place config folder in same directory as executable")
        except Exception as e:
            print(f"Packaging failed: {e}")

if __name__ == "__main__":
    package_program()
