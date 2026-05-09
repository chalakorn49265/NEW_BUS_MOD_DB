"""
文成客户经济性 — 唯一入口（不含仓库根目录 multipage）。

Streamlit 只在「本脚本所在目录」查找 pages/；此处无 pages/，侧栏不会出现其它项目页面。

启动（在仓库根目录执行）::

    python3 -m streamlit run wenzhou/run_dashboard.py
"""

from __future__ import annotations

from wencheng_client_dashboard import main

if __name__ == "__main__":
    main()
