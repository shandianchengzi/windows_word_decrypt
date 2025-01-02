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