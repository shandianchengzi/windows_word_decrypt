import os
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def generate_passwords(try_words):
    """
    根据给定的单词列表生成所有可能的密码组合。
    
    :param try_words: 用于生成密码的单词列表
    :return: 密码列表
    """
    try_passwd = []
    for i in range(1, 4):
        for j in range(0, len(try_words)):
            for k in range(0, len(try_words)):
                for l in range(0, len(try_words)):
                    passwd = try_words[j]
                    if i > 1:
                        passwd += try_words[k]
                    if i > 2:
                        passwd += try_words[l]
                    try_passwd.append(passwd)
    
    # 去重并返回
    return list(set(try_passwd))

def try_decrypt_file(file_path, try_words, result_path):
    """
    尝试使用不同的密码解密文件并记录结果。
    
    :param file_path: 加密的 Word 文件路径
    :param try_words: 密码候选词列表
    :param result_path: 结果输出文件路径
    """
    # 生成密码列表
    try_passwd = generate_passwords(try_words)

    # 清空结果文件
    with open(result_path, "w") as f:
        f.write("")

    if file_path.endswith(".doc"):
        from doc_decrypt import read_secret_word_file
    elif file_path.endswith(".docx"):
        from docx_decrypt import read_secret_word_file

    # 尝试每个密码
    for passwd in try_passwd:
        try:
            read_secret_word_file(file_path, passwd)
            # 如果解密成功，写入结果并退出
            logging.info(f"成功解密文件，密码是: {passwd}")
            with open(result_path, "a") as f:
                f.write(f"{passwd} 密码正确\n")
            break
        except Exception as e:
            # 如果解密失败，记录错误并继续尝试
            logging.error(f"解密失败，密码错误: {passwd}, 错误信息: {e}")
            with open(result_path, "a") as f:
                f.write(f"{passwd} 密码错误，错误信息: {e}\n")

# 使用示例
if __name__ == "__main__":
    try_words = ["password", "1234", "qwerty"]  # 自定义密码尝试词
    file_path = "secret.docx"  # 加密的文件路径
    result_path = "result.txt"  # 结果文件路径

    # 将文件路径转换为绝对路径
    file_path = os.path.abspath(file_path)
    logging.info(f"使用的文件路径是: {file_path}")

    if not os.path.exists(file_path):
        logging.error("文件不存在，请检查文件路径是否正确")
        exit(1)

    try_decrypt_file(file_path, try_words, result_path)