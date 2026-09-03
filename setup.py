from setuptools import setup, find_packages

setup(
    name="hr-toolkit",
    version="0.8.1",
    description="人事 Excel 自动化工具箱",
    packages=find_packages(include=["hr_toolkit*"]),
    package_data={"hr_toolkit.templates": ["*.xlsx"]},
    install_requires=[
        "certifi>=2024.8.30",
        "openpyxl>=3.1,<4",
        "py7zr>=0.20,<2",
        "pypdf[image]>=4.0.0,<7",
        "xlrd>=2.0,<3",
        "unrar2-cffi>=0.4.1,<0.6; python_version >= '3.10'",
        "rapidocr_onnxruntime>=1.3.0",
        'pywin32>=306; platform_system == "Windows"',
    ],
    entry_points={
        "console_scripts": [
            "hr-toolkit=hr_toolkit.cli:main",
        ],
    },
)
