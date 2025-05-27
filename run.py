import re
import os
from docx import Document

from openai import OpenAI


'''文档分离模块'''
def sanitize_filename(filename):
    """清理文件名中的非法字符"""
    filename = re.sub(r'[\\/*?:"<>|]', '_', filename)
    filename = filename.strip()[:200]
    return filename if filename else 'untitled'

def parse_docx(file_path):
    """解析Word文档并提取标题结构及内容"""
    document = Document(file_path)
    entries = []
    current_path = []
    current_content = []
    
    for para in document.paragraphs:
        if para.style.name.startswith('Heading'):
            match = re.match(r'Heading\s*(\d+)', para.style.name)
            if not match:
                continue
            level = int(match.group(1))
            heading_text = para.text.strip()
            
            if current_path:
                entries.append((list(current_path), '\n'.join(current_content)))
                current_content = []
            
            del current_path[level-1:]
            current_path.append(heading_text)
        else:
            text = para.text.strip()
            if text:
                current_content.append(text)
    
    if current_path:
        entries.append((list(current_path), '\n'.join(current_content)))
    
    return entries

def save_result1_to_file(entries, filename):
    """保存结果一到指定路径"""
    with open(filename, 'w', encoding='utf-8') as f:
        f.write("【结果一：各级标题及内容】\n\n")
        for path, content in entries:
            indent = '    ' * (len(path)-1)
            f.write(f"{indent}▶ {' → '.join(path)}\n")
            f.write(f"{indent}{content}\n\n")

def save_result2_to_files(entries):
    """保存结果二到指定目录"""
    output_dir = r'D:\python\docx\specific'
    os.makedirs(output_dir, exist_ok=True)  # 确保目录存在
    
    for index, (path, content) in enumerate(entries):
        full_path = ' → '.join(path)
        base_name = sanitize_filename(full_path)
        
        if not base_name:
            base_name = f'untitled_{index}'
        
        # 拼接完整输出路径
        filename = os.path.join(output_dir, f"{base_name}.txt")
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)

'''全文及框架检测'''

def runall():
    # 读取论文文件内容
    with open('overall.txt', 'r', encoding='utf-8') as f:
        paper_content = f.read()

    client = OpenAI(api_key="sk-xxx", base_url="https://api.deepseek.com")

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": "你是一个论文的评判员，根据论文的好坏给出ABCD四个等级。你要查看全文，并且注重整体框架，给出检测结果和具体修改建议"},
            {"role": "user", "content": f"论文内容如下：{paper_content}"},
        ],
        stream=False
    )

    # 将AI回答写入文件
    with open('overall_AI.txt', 'w', encoding='utf-8') as f:
        f.write(response.choices[0].message.content)

    # 可选：同时在控制台打印结果
    print("AI分析结果已保存到overall_AI.txt")
    print(response.choices[0].message.content)


'''分别进行段落检测'''
def runspecific():
    # 配置DeepSeek客户端
    client = OpenAI(
        api_key="sk-xxx",
        base_url="https://api.deepseek.com"
    )

    # 定义处理目录和输出目录
    input_dir = r"D:\python\docx\specific"
    output_dir = os.path.join(input_dir, "AI评审结果")
    os.makedirs(output_dir, exist_ok=True)  # 自动创建输出目录

    def process_file(file_path, output_dir):
        """处理单个文件并保存结果"""
        try:
            # 读取文件内容
            with open(file_path, 'r', encoding='utf-8') as f:
                paper_content = f.read()

            # 获取AI评审结果
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": "你是一个论文的评判员，根据论文的好坏给出ABCD四个等级，我会给你论文的部分段落，请你给出段落的检测结果以及修改建议。"},
                    {"role": "user", "content": f"论文内容如下：{paper_content}"},
                ],
                stream=False
            )
            
            # 生成输出文件名
            filename = os.path.basename(file_path)
            base_name, ext = os.path.splitext(filename)
            output_path = os.path.join(output_dir, f"{base_name}_AI评审结果{ext}")

            # 保存结果
            with open(output_path, 'w', encoding='utf-8') as f:
                result = response.choices[0].message.content
                f.write(f"文件名：{filename}\n\n{result}")
            
            print(f"成功处理：{filename}")

        except Exception as e:
            print(f"处理文件 {filename} 时出错：{str(e)}")

    # 遍历目录处理所有文件
    for filename in os.listdir(input_dir):
        file_path = os.path.join(input_dir, filename)
        if os.path.isfile(file_path):
            process_file(file_path, output_dir)

    print("\n所有文件处理完成，结果已保存至：", output_dir)


if __name__ == "__main__":
    print("检测时间跟论文大小有关，请耐心等待")
    docx_file = r"D:\python\docx\docx\10086.docx"  # 修改为你的文档路径
    entries = parse_docx(docx_file)
    
    # 保存结果一
    save_result1_to_file(entries, r'D:\python\docx\overall.txt')
    
    # 保存结果二
    save_result2_to_files(entries)
    
    print("处理完成！结果一已保存到 D:\\python\\docx\\overall.txt")
    print("结果二文件已保存到 D:\\python\\docx\\specific 目录")

    #检测全文及框架
    runall()

    #检测段落
    runspecific()
