"""
文成隧道路灯改造 — 客户经济性（独立 Streamlit 应用）。

与其它 multipage 应用分离：仅包含本看板，无 pages 侧栏其它页面。

用法（在仓库根目录）::

    streamlit run streamlit_wenzhou_client.py

依赖：requirements.txt（streamlit、pandas、plotly、numpy-financial 等）。
"""

from __future__ import annotations

from wenzhou.wencheng_client_dashboard import main

if __name__ == "__main__":
    main()
