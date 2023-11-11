import os


def convert_hexo_mermaid_code_to_markdown_mermaid_code(file_path):
    content = ''
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    content = content.replace('{% mermaid %}', '```mermaid')
    content = content.replace('{% endmermaid %}', '```')
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)


def convert_markdown_mermaid_code_to_hexo_mermaid_code(file_path):
    content = ''
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    flag = False
    for i in range(0, len(content)):
        if i+10 < len(content) and content[i] == '`' and content[i + 1] == '`' and content[i + 2] == '`' and \
                content[i + 3] == 'm' and content[i + 4] == 'e' and content[i + 5] == 'r' and \
                content[i + 6] == 'm' and content[i + 7] == 'a' and content[i + 8] == 'i' and \
                content[i + 9] == 'd' and content[i + 10] == '\n' and not flag:
            content = content[:i] + '{% mermaid %}' + content[i + 3:]
            flag = True
            break
        if i+2 < len(content) and content[i] == '`' and content[i + 1] == '`' and content[i + 2] == '`' and flag:
            content = content[:i] + '{% endmermaid %}' + content[i + 3:]
            flag = False
            break
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)


if __name__ == '__main__':
    print("Please choose the mode of the converter: ")
    print("1. Convert hexo mermaid code to markdown mermaid code")
    print("2. Convert markdown mermaid code to hexo mermaid code")
    mode = input("Please input the mode number: ")
    if mode == '1':
        for root, dirs, files in os.walk('./source/_posts'):
            for file in files:
                if file.endswith('.md'):
                    print("Converting file: " + file)
                    convert_hexo_mermaid_code_to_markdown_mermaid_code(os.path.join(root, file))

    elif mode == '2':
        for root, dirs, files in os.walk('./source/_posts'):
            for file in files:
                if file.endswith('.md'):
                    print("Converting file: " + file)
                    convert_markdown_mermaid_code_to_hexo_mermaid_code(os.path.join(root, file))
