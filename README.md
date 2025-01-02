# 概述

这是一个用于批量尝试密码去打开加密的 Word 文档（docx和doc）的 Python 脚本。

# 详细解释

[【代码】Python｜Windows 批量尝试密码去打开加密的 Word 文档（docx和doc）](https://shandianchengzi.blog.csdn.net/article/details/144888638)

# 使用方法

1. 安装依赖

```bash
pip install -r requirements.txt
```
2. 修改main.py中的配置

```python
    try_words = ["password", "1234", "qwerty"]  # 自定义密码尝试词
    file_path = "secret.docx"  # 加密的文件路径
    result_path = "result.txt"  # 结果文件路径
```
3. 运行main.py

# 注意事项

- 本脚本仅供学习交流使用，请勿用于非法用途。
- 解密doc文档的时候，需要安装Microsoft Office Word（好像有人说WPS之类的也行，我没试过），并在运行脚本之前打开，否则会报错“RPC 服务器不可用”。