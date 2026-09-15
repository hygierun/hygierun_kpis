from setuptools import setup, find_packages

setup(
    name="hygierun-kpi-tool",
    version="1.0.0",
    description="Générateur de rapports KPI Hygierun",
    author="Hygierun Analytics",
    packages=find_packages(),
    python_requires=">=3.8",
    install_requires=[
        "streamlit==1.63.0",
        "pandas>=1.4.0",
        "openpyxl>=3.1.0",
    ],
)
