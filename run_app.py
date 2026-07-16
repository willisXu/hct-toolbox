# -*- coding: utf-8 -*-
"""EXE 進入點：用 PyInstaller 凍結後，啟動 Streamlit 介面。

設計（與 MOMO 日結工具相同）：app.py / core / mappings 都「不」打包進 exe，
而是放在 exe 旁邊，維持可編輯（改範本、調規則不用重新打包）。
"""
import os
import sys


def _suppress_streamlit_email_prompt():
    """首次啟動時 Streamlit 會在終端機詢問 email 並等待輸入，凍結成 EXE 後會卡死。
    預先寫一個空 email 的 credentials.toml（若不存在）即可跳過此提問。"""
    try:
        cfg_dir = os.path.join(os.path.expanduser("~"), ".streamlit")
        cred = os.path.join(cfg_dir, "credentials.toml")
        if not os.path.exists(cred):
            os.makedirs(cfg_dir, exist_ok=True)
            with open(cred, "w", encoding="utf-8") as f:
                f.write('[general]\nemail = ""\n')
    except OSError:
        pass


def main():
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)          # exe 所在資料夾
    else:
        base = os.path.dirname(os.path.abspath(__file__))

    os.chdir(base)
    sys.path.insert(0, base)                             # 讓 app.py 能 import core

    _suppress_streamlit_email_prompt()

    app = os.path.join(base, "app.py")
    if not os.path.exists(app):
        print("找不到 app.py，請確認 exe 與 app.py / core / mappings 在同一資料夾")
        input("按 Enter 結束...")
        sys.exit(1)

    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit", "run", app,
        "--global.developmentMode=false",
        "--server.port=8502",
        "--browser.gatherUsageStats=false",
    ]
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
