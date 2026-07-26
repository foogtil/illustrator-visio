# -*- coding: utf-8 -*-
"""
在 Illustrator 中执行 ExtendScript(.jsx) 文件。

用法:
  python run_jsx.py <script.jsx> [--quit]

行为:
  - 优先附加到正在运行的 Illustrator 实例, 否则启动新实例。
  - 以 UTF-8 读取 .jsx 文件, 通过 COM DoJavaScript 执行(支持中文)。
  - JSX 中最后的表达式值会作为返回值打印出来, 可用于回传结果。
  - JSX 抛出的异常会以非零退出码报告, 并给出 ExtendScript 错误信息。
  - --quit: 仅当 Illustrator 是本脚本启动的, 执行后退出 Illustrator。

注意: 脚本只负责"执行"。是否保存/关闭文档由 JSX 自己决定。
"""
import argparse
import sys

# 管道/重定向环境下强制 UTF-8 输出, 避免中文乱码
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("jsx", help="要执行的 .jsx 文件路径(UTF-8 编码)")
    parser.add_argument("--quit", action="store_true",
                        help="若 Illustrator 由本脚本启动, 执行后退出它")
    args = parser.parse_args()

    try:
        with open(args.jsx, "r", encoding="utf-8-sig") as f:  # utf-8-sig: 自动剥离 BOM
            code = f.read()
    except OSError as e:
        print("无法读取 JSX 文件: %s" % e, file=sys.stderr)
        sys.exit(2)

    import time
    import pythoncom
    import win32com.client

    SERVER_EXEC_FAILURE = -2146959355  # 0x80080005, Illustrator 冷启动中

    try:
        app = win32com.client.GetActiveObject("Illustrator.Application")
        started = False
    except pythoncom.com_error:
        app = None
        last_err = None
        for attempt in range(4):
            try:
                app = win32com.client.Dispatch("Illustrator.Application")
                break
            except pythoncom.com_error as e:
                last_err = e
                if e.hresult == SERVER_EXEC_FAILURE and attempt < 3:
                    time.sleep(5)
                    continue
                raise
        if app is None:
            raise last_err
        started = True

    rc = 0
    try:
        try:
            result = app.DoJavaScript(code)
        except pythoncom.com_error as e:
            # ExtendScript 异常通常体现在 COM 错误描述里
            msg = e.excepinfo[2] if e.excepinfo else str(e)
            print("JSX 执行失败: %s" % msg, file=sys.stderr)
            rc = 1
        else:
            if result is not None and str(result) != "":
                print(result)
    finally:
        if started and args.quit:
            try:
                app.Quit()
            except Exception:
                pass
    sys.exit(rc)

if __name__ == "__main__":
    main()
