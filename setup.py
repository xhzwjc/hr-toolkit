from setuptools import setup

# 兼容旧版本 pip（如 pip < 21.3）不支持直接读取 pyproject.toml 进行 editable 安装的情况
if __name__ == "__main__":
    setup()
