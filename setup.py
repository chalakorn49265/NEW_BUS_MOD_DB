from setuptools import find_packages, setup

setup(
    name="emc-institutional-model",
    version="0.1.0",
    packages=find_packages(include=["emc_institutional_model*"]),
    python_requires=">=3.9",
    install_requires=[
        "numpy>=1.24",
        "pandas>=2.0",
        "pydantic>=2.5",
        "streamlit>=1.28",
        "plotly>=5.18",
        "pytest>=7.4",
        "numpy-financial>=1.0",
    ],
)
