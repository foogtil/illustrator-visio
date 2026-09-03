# -*- coding: utf-8 -*-
"""
illustrator skill 环境自检脚本。

用法:  python check_env.py

检查内容:
  1. pywin32 是否可用
  2. Illustrator COM 是否可连接(版本号)
  3. Illustrator 可用中文字体清单
输出供 agent 决策使用: COM 是否可用、字体真实名称。
"""
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

    Illustrator 冷启动较慢, COM 激活可能返回 CO_E_SERVER_EXEC_FAILURE,
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

def check_illustrator():
    print("== Illustrator ==")
    try:
        app, started = get_or_create("Illustrator.Application")
    except Exception as e:
        print("  COM: 不可用 (%s)" % e)
        return
    try:
        print("  COM: 可用, 版本 %s" % app.Version)
        try:
            fonts = app.TextFonts
            names = []
            for i in range(1, min(fonts.Count, 2000) + 1):
                try:
                    names.append(fonts.Item(i).Name)
                except Exception:
                    break
            cjk = [n for n in names if any(k in n for k in
                   ("YaHei", "SimSun", "SimHei", "KaiTi", "FangSong",
                    "PingFang", "SourceHan", "Source Han", "Noto"))]
            print("  字体总数(采样): %d; 中文字体: %s"
                  % (len(names), ", ".join(cjk[:12]) if cjk else "未检测到"))
            # 科研插图常用西文字体探测
            western = [n for n in names if any(k in n.lower() for k in
                       ("arial", "helvetica", "times"))]
            print("  西文字体(论文常用): %s"
                  % (", ".join(western[:12]) if western else "未检测到"))
        except Exception as e:
            print("  字体枚举失败(可忽略): %s" % e)
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
    check_illustrator()
    print("== 自检完成 ==")

if __name__ == "__main__":
    main()
