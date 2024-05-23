import os
import time
import pandas as pd
import argparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 设置Chrome和Chromedriver的路径
chrome_path = '/usr/bin/google-chrome'
chromedriver_path = '/usr/local/bin/chromedriver'

# 打印路径以进行调试
print(f"Chrome Path: {chrome_path}")
print(f"Chromedriver Path: {chromedriver_path}")

# 检查路径和权限
if not os.path.isfile(chrome_path):
    print(f"Chrome not found at {chrome_path}")
if not os.access(chrome_path, os.X_OK):
    print(f"Chrome at {chrome_path} is not executable")

if not os.path.isfile(chromedriver_path):
    print(f"Chromedriver not found at {chromedriver_path}")
if not os.access(chromedriver_path, os.X_OK):
    print(f"Chromedriver at {chromedriver_path} is not executable")

# 配置下载目录
download_dir = '/tmp/selenium_downloads'
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

# Chrome下载配置
options = Options()
options.add_argument('--headless')  # 无头模式
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

# Function to download and parse the table
def download_and_parse_table(pdb_id):
    url = f'https://www.nakb.org/naparams.html?{pdb_id}_1#BACKBONE/'
    print(f"Accessing URL: {url}")  # 调试信息

    # 使用Service指定Chromedriver路径
    service = ChromeService(executable_path=chromedriver_path)
    try:
        driver = webdriver.Chrome(service=service, options=options)
        print(f"WebDriver initialized successfully for PDB ID: {pdb_id}")
    except Exception as e:
        print(f"Error initializing WebDriver for {pdb_id}: {e}")
        return None
    
    try:
        driver.get(url)
    
        # 使用WebDriverWait等待元素加载
        try:
            wait = WebDriverWait(driver, 20)  # 等待最长20秒
            # 尝试点击“CSV”按钮
            try:
                csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'buttons-csv')]")))
                csv_button.click()
                print(f"Clicked CSV button for {pdb_id}")
            except Exception as e:
                print(f"CSV button not found for {pdb_id}, trying Excel button: {e}")
                # 尝试点击“Excel”按钮
                excel_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'buttons-excel')]")))
                excel_button.click()
                print(f"Clicked Excel button for {pdb_id}")
        except Exception as e:
            print(f"Error finding or clicking download button for {pdb_id}: {e}")
            driver.quit()
            return None
        
        # 等待文件下载完成
        time.sleep(10)  # 可根据网络速度调整等待时间
        
        # 找到最新下载的文件
        files = sorted(os.listdir(download_dir), key=lambda x: os.path.getctime(os.path.join(download_dir, x)), reverse=True)
        if not files:
            print(f"No files found in download directory for {pdb_id}")
            driver.quit()
            return None

        latest_file = os.path.join(download_dir, files[0])
        driver.quit()
    
        # 读取CSV文件
        try:
            table = pd.read_csv(latest_file)
            return table
        except Exception as e:
            print(f"Error reading CSV file for {pdb_id}: {e}")
            return None
    finally:
        driver.quit()

# Function to extract DNA sequence from table
def extract_dna_sequence_from_table(table):
    # Extract the 'Residue Name' column
    residue_names = table['Residue Name']
    
    # Concatenate the residue names and remove all 'D' characters
    sequence = ''.join(residue_names).replace('D', '')
    return sequence

# 解析命令行参数
parser = argparse.ArgumentParser(description='Process PDB IDs from an Excel file and output sequences.')
parser.add_argument('-i', '--input_file', type=str, help='Input Excel file containing PDB IDs')
parser.add_argument('-o', '--output_file', type=str, help='Output Excel file with sequences')
args = parser.parse_args()

# 读取输入文件
input_df = pd.read_excel(args.input_file)

# 确保PDB列存在
if 'PDB' not in input_df.columns:
    print(f"Error: 'PDB' column not found in {args.input_file}")
    exit(1)

# 处理每个PDB ID并获取序列
results = []
errors = []

for pdb_id in input_df['PDB']:
    table = download_and_parse_table(pdb_id)
    if table is not None:
        dna_sequence = extract_dna_sequence_from_table(table)
        results.append(dna_sequence)
    else:
        results.append(None)
        errors.append(pdb_id)

# 添加序列列到原始数据框
input_df['sequence'] = results

# 保存输出文件
input_df.to_excel(args.output_file, index=False)

# 输出错误的PDB ID
if errors:
    print("The following PDB IDs encountered errors:")
    for error in errors:
        print(error)

print(f"Results saved to {args.output_file}")
