import subprocess
import sys
from config import operation_steps
#######################
###  第一步：配置环境  ###
######################

# 第三方包列表
required_packages = [
    'selenium',
    'openpyxl',
    'requests',
    'pandas',
    'beautifulsoup4',  # bs4 实际包名
    'pillow'           # PIL 实际包名
]

# 使用清华源安装缺失包
def install_if_missing(package_list):
    for pkg in package_list:
        try:
            __import__(pkg if pkg != 'beautifulsoup4' else 'bs4')
        except ImportError:
            subprocess.check_call([
                sys.executable, '-m', 'pip', 'install', pkg,
                '-i', 'https://pypi.tuna.tsinghua.edu.cn/simple'
            ])
install_if_missing(required_packages)

if operation_steps[0] == 1:
    # 运行 'crawler.py' 模块
    print("\n>>> 正在运行模块: crawler.py")
    result_crawler = subprocess.run([sys.executable, 'crawler.py'])
    if result_crawler.returncode != 0:
        print(f"模块 crawler.py 执行失败，退出码 {result_crawler.returncode}")
        sys.exit(result_crawler.returncode)

if operation_steps[1] == 1:
    # 运行 'pipeline.py' 模块
    print("\n>>> 正在运行模块: pipeline.py")
    result_pipeline = subprocess.run([sys.executable, 'pipeline.py'])
    if result_pipeline.returncode != 0:
        print(f"模块 pipeline.py 执行失败，退出码 {result_pipeline.returncode}")
        sys.exit(result_pipeline.returncode)
if operation_steps[2] == 1:
    # 运行 'final_process.py' 模块
    print("\n>>> 正在运行模块: final_process.py")
    result_final_process = subprocess.run([sys.executable, 'final_process.py'])
    if result_final_process.returncode != 0:
        print(f"模块 final_process.py 执行失败，退出码 {result_final_process.returncode}")
        sys.exit(result_final_process.returncode)

print("\n" + "#"*50)
print("#"*20, "全部模块已执行完成", "#"*20)
print("#"*50)