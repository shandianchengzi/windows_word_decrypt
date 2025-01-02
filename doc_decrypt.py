'''
pip install comtypes
'''

from comtypes.client import CreateObject

def read_secret_word_file(filename, password):
    """
    使用指定的密码打开加密的 Word 文件。
    
    :param filename: 加密的 Word 文件路径
    :param password: 用于解密的密码
    :return: None
    """
    try:
        # 启动 Word 应用程序
        word = CreateObject('Word.Application')
        word.Visible = False  # 设置为不可见
        
        # 打开加密的 Word 文件
        doc = word.Documents.Open(filename, PasswordDocument=password)
        doc.Close()
    except Exception as e:
        raise