# -*- coding: utf-8 -*-
"""
visio skill 环境自检脚本。

用法:  python check_env.py

检查内容:
  1. pywin32 是否可用
  2. Visio COM 是否可连接(版本号) + 可用模具(stencil)清单
输出供 agent 决策使用: COM 是否可用、模具文件的真实路径。
"""
import glob
import os
import sys

# 管道/重定向环境下强制 UTF-8 输出, 避免中文乱码
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SERVER_EXEC_FAILURE = -2146959355  # 0x80080005, 应用正在启动时的典型错误

def get_or_create(progid, retries=4, wait=5):
    """附加到已运行实例; 没有则新建。返回 (app, started_by_us)。

    Visio 冷启动较慢, COM 激活可能返回 CO_E_SERVER_EXEC_FAILURE,
    此处自动重试。
    """
    import time
    import pythoncom
    import win32com.client
    try:
        return win32com.client.GetActiveObject(progid), False
    except pythoncom.com_error:
        pass
    last_err = None
    for attempt in range(retries):
        try:
            return win32com.client.Dispatch(progid), True
        except pythoncom.com_error as e:
            last_err = e
            if e.hresult == SERVER_EXEC_FAILURE and attempt < retries - 1:
                time.sleep(wait)
                continue
            raise
    raise last_err

def check_visio():
    print("== Visio ==")
    try:
        app, started = get_or_create("Visio.Application")
    except Exception as e:
        print("  COM: 不可用 (%s)" % e)
        return
    try:
        print("  COM: 可用, 版本 %s" % app.Version)
        office_root = app.Path  # 形如 C:\Program Files\Microsoft Office\root\Office16\
        content = os.path.join(office_root, "Visio Content")
        dirs = sorted(glob.glob(os.path.join(content, "*")))
        dirs = [d for d in dirs if os.path.isdir(d)]
        if not dirs:
            print("  模具目录未找到: %s" % content)
            return
        # 优先中文(2052), 其次英文(1033), 否则第一个
        pref = None
        for want in ("2052", "1033"):
            cand = [d for d in dirs if os.path.basename(d) == want]
            if cand:
                pref = cand[0]
                break
        pref = pref or dirs[0]
        print("  模具语言目录: %s (共 %d 个语言目录)" % (pref, len(dirs)))
        stencils = sorted(os.path.basename(p) for p in
                          glob.glob(os.path.join(pref, "*.vss*")))
        print("  模具数量: %d" % len(stencils))
        for s in stencils:
            print("    - %s" % s)
    finally:
        if started:
            try:
                app.Quit()
            except Exception:
                pass

def main():
    try:
        import win32com.client  # noqa: F401
        print("pywin32: 可用")
    except ImportError:
        print("pywin32: 未安装! 请先执行: python -m pip install pywin32")
        sys.exit(1)
    check_visio()
    print("== 自检完成 ==")

if __name__ == "__main__":
    main()
