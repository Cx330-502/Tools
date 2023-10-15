import os
from urllib.parse import quote
import requests
from tqdm import tqdm


def init_print():
    print("Welcome to use Cx330_502's openalex_downloader program, the program running time may be long, please be "
          "patient!")
    print("    ___  _  _  ___  ___   ___       ___   ___  ___  ")
    print("   / __)( \\/ )(__ )(__ ) / _ \\     | __) / _ \\(__ \\ ")
    print("  ( (__  )  (  (_ \\ (_ \\( (_) )___ |__ \\( (_) )/ _/ ")
    print("   \\___)(_/\\_)(___/(___/ \\___/(___)(___/ \\___/(____)")
    print()


def input_model():
    print("Please input the model you want to use:")
    print("1. input and output both in current directory")
    print("2. './data/openalex_downloader/' for input and './output/openalex_downloader/' for output")
    print("3. Customizing the working directory")
    while True:
        model = input("Please input 1 or 2 or 3: ")
        model = int(model)
        if model == 1:
            input_root0 = "./"
            output_root0 = "./"
            break
        elif model == 2:
            input_root0 = "./data/openalex_downloader/"
            output_root0 = "./data/openalex_downloader/"
            break
        elif model == 3:
            input_root0 = input("Please input the input directory: ")
            output_root0 = input("Please input the output directory: ")
            break
    os.makedirs(input_root0, exist_ok=True)
    os.makedirs(output_root0, exist_ok=True)
    return input_root0, output_root0

def print_error(e, content):
    print(e, content)
    with open(os.path.join(output_root, "error.txt"), "a") as f3:
        f3.write(content + "\n")

def get_remote_file_size(url):
    response = requests.head(url)
    if response.status_code == 200:
        return int(response.headers['Content-Length'])
    else:
        return None


def download_file(local_filename):
    url = "https://openalex.s3.amazonaws.com/"
    temp_url = url + local_filename
    temp_url = quote(temp_url, safe='/:?=#&')
    local_filename = os.path.join(output_root, local_filename)

    remote_file_size = get_remote_file_size(temp_url)
    if remote_file_size is not None:
        if os.path.exists(local_filename) and os.path.getsize(local_filename) == remote_file_size:
            print(f"文件 {local_filename} 已存在且大小相同，跳过下载。")
            return

    response = requests.get(temp_url, stream=True)
    os.makedirs(os.path.dirname(local_filename), exist_ok=True)

    file_size = int(response.headers['Content-Length'])
    block_size = 1024
    with open(local_filename, 'wb') as file, tqdm(
            desc=local_filename,
            total=file_size,
            unit='iB',
            unit_scale=True,
            unit_divisor=1024,
    ) as bar:
        for data in response.iter_content(block_size):
            file.write(data)
            bar.update(len(data))


if __name__ == '__main__':
    init_print()
    input_root, output_root = input_model()
    with open(os.path.join(input_root, 'openalex.txt'), "r") as f:
        if os.path.exists(os.path.join(output_root, "error.txt")):
            os.remove(os.path.join(output_root, "error.txt"))
        for line in f.readlines():
            content = line.split(' ')[-1]
            content = content.replace('\n', '')
            print("Downloading:", content)
            if not content.startswith("data/"):
                continue
            else:
                try:
                    download_file(content)
                except Exception as e:
                    print_error(e, content)
                    continue
