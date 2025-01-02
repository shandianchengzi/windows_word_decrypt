'''
pip install msoffcrypto-tool
'''

import msoffcrypto
from io import BytesIO

def read_secret_word_file(file_path, password):
    """
    解密加密的 Word 文件并返回解密的内容。
    
    :param file_path: 加密的 Word 文件路径
    :param password: 用于解密的密码
    :return: None
    """
    try:
        with open(file_path, 'rb') as encrypted_file:
            # 使用 msoffcrypto 解密
            office_file = msoffcrypto.OfficeFile(encrypted_file)
            office_file.load_key(password=password)  # 提供密码
            
            # 尝试解密文件
            decrypted_content = BytesIO()
            office_file.decrypt(decrypted_content)
            return decrypted_content.getvalue()
    except Exception as e:
        raise