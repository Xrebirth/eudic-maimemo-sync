# ====================== 配置区（请根据实际情况修改） ======================
# 需要运行的Obsidian同步脚本路径（支持相对/绝对路径）
OBSIDIAN_SYNC_SCRIPT_PATH = r"C:\Program Files\DevelopmentTools\Automation\eudic-maimemo-sync\sync obsidian.py"
# =========================================================================

import os
import requests
import json
import subprocess
import sys
from collections import defaultdict
from dotenv import load_dotenv
from datetime import datetime, timezone, timedelta

# -------------------------- 使用plyer实现Windows通知（无兼容问题） --------------------------
def show_windows_notification(title, message, is_success=True):
    """
    使用plyer实现Windows系统通知（彻底解决WPARAM/LRESULT兼容问题）
    :param title: 通知标题
    :param message: 通知内容
    :param is_success: 是否为成功类通知（仅用于标识，不影响功能）
    """
    try:
        # 优先使用plyer库（跨平台、更稳定）
        from plyer import notification
        notification.notify(
            title=title,
            message=message,
            app_name="欧路-墨墨单词同步工具",  # 通知栏显示的应用名称
            timeout=8,                          # 通知显示时长（秒）
            ticker="单词同步提醒"               # 通知滚动提示（可选）
        )
    except ImportError:
        # 若缺少plyer，自动安装并重试
        print("\n[INFO] 正在安装plyer依赖库...")
        os.system(f"{sys.executable} -m pip install plyer -q")
        try:
            from plyer import notification
            notification.notify(
                title=title,
                message=message,
                app_name="欧路-墨墨单词同步工具",
                timeout=8
            )
        except Exception as e:
            # 兜底：控制台打印通知内容
            print(f"\n[通知] {title}：{message}")
            print(f"[WARNING] 通知弹窗失败：{e}")
    except Exception as e:
        # 捕获其他异常，保证程序不崩溃
        print(f"\n[ERROR] 通知展示失败：{e}")
        print(f"[通知] {title}：{message}")

# -------------------------- 交互式选择函数 --------------------------
def ask_run_obsidian_sync():
    """
    交互式选择是否运行配置的Obsidian同步脚本
    返回：True（运行）/ False（不运行）
    """
    print("\n" + "="*50)
    # 循环直到输入有效
    while True:
        choice = input(f"是否将同步的生词本添加到Obsidian中？(Y/N): ").strip().upper()
        if choice in ["Y", "YES"]:
            return True
        elif choice in ["N", "NO"]:
            return False
        else:
            print("⚠️ 输入无效！请输入 Y (是) 或 N (否)")

def run_obsidian_sync_script():
    """
    运行配置区指定的Obsidian同步脚本，修复编码问题，兼容中文输出
    """
    script_path = OBSIDIAN_SYNC_SCRIPT_PATH  # 从配置区读取路径
    # 检查脚本是否存在
    if not os.path.exists(script_path):
        error_msg = f"未找到脚本文件：{script_path}"
        print(f"[ERROR] {error_msg}")
        show_windows_notification(
            title="Obsidian同步失败",
            message=error_msg,
            is_success=False
        )
        return False
    
    try:
        print(f"[INFO] 开始运行 {script_path} 脚本...")
        # 修复编码问题：使用gbk解码（Windows默认编码），忽略无法解码的字符
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            encoding="gbk",  # 替换utf-8为gbk，兼容Windows中文输出
            errors="ignore"  # 忽略无法解码的字符，避免崩溃
        )
        
        # 打印脚本输出（清理空行）
        if result.stdout:
            stdout_clean = "\n".join([line for line in result.stdout.splitlines() if line.strip()])
            if stdout_clean:
                print(f"[INFO] {script_path} 输出：\n{stdout_clean}")
        if result.stderr:
            stderr_clean = "\n".join([line for line in result.stderr.splitlines() if line.strip()])
            if stderr_clean:
                print(f"[WARNING] {script_path} 运行警告：\n{stderr_clean}")
        
        # 仅以返回码判断是否成功（忽略通知库的警告）
        if result.returncode == 0:
            success_msg = f"{script_path} 运行完成！"
            print(f"[SUCCESS] {success_msg}")
            show_windows_notification(
                title="Obsidian同步成功",
                message="生词本已成功添加到Obsidian Canvas！",
                is_success=True
            )
            return True
        else:
            error_msg = f"{script_path} 运行失败（返回码：{result.returncode}）"
            print(f"[ERROR] {error_msg}")
            show_windows_notification(
                title="Obsidian同步失败",
                message=error_msg,
                is_success=False
            )
            return False
    except Exception as e:
        error_msg = f"运行 {script_path} 时发生异常：{str(e)}"
        print(f"[ERROR] {error_msg}")
        show_windows_notification(
            title="Obsidian同步失败",
            message=error_msg,
            is_success=False
        )
        return False

# -------------------------- 原有功能代码 --------------------------
def fetch_word_list():
    """获取欧路词典生词本"""
    load_dotenv()
    
    headers = {
        "Authorization": os.getenv("EUDIC_API_KEY"),
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    }
    
    url = "https://api.frdic.com/api/open/v1/studylist/words/{id}".format(id=os.getenv("EUDIC_CATEGORY_ID"))

    try:
        response = requests.get(url, headers=headers, params={"language": "en"})
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        print(f"[ERROR] 获取单词列表失败: {e}")
        return None

def generate_word_output(word_data):
    """生成按日期分组的单词字符串，并将UTC时间转换为中国时间"""
    if not word_data or 'data' not in word_data:
        return ""

    # 中国时区 (UTC+8)
    china_tz = timezone(timedelta(hours=8))
    
    grouped_words = defaultdict(list)
    for item in word_data['data']:
        # 解析UTC时间
        utc_time = datetime.fromisoformat(item["add_time"].replace('Z', '+00:00'))
        # 转换为中国时间
        china_time = utc_time.astimezone(china_tz)
        # 获取中国时区的日期
        date = china_time.strftime("%Y-%m-%d")
        
        grouped_words[date].append(item["word"])

    output_string = ""
    for date in sorted(grouped_words.keys()):
        output_string += f"#{date}\n"
        output_string += "\n".join(grouped_words[date])
        output_string += "\n"

    return output_string

def update_maimemo_notepad(content):
    """同步到墨墨背单词"""
    # 加载环境变量
    load_dotenv()
    
    # 获取 API 密钥和笔记本 ID
    api_key = os.getenv("MOMO_API_KEY")
    notepad_id = os.getenv("MOMO_NOTEPAD_ID")
    
    # 请求 URL
    url = f"https://open.maimemo.com/open/api/v1/notepads/{notepad_id}"
    
    # 请求头
    headers = {
        'Accept': 'application/json',
        'Authorization': f'Bearer {api_key}',
        'Content-Type': 'application/json'
    }
    
    # 请求数据
    payload = {
        "notepad": {
            "status": "UNPUBLISHED",
            "content": content,
            "title": "欧路生词",
            "brief": "欧路词典上查询过的单词",
            "tags": ["自用"]
        }
    }
    
    try:
        # 发送 POST 请求（注：墨墨更新生词本应该用 PUT 方法？需确认API文档）
        response = requests.post(url, json=payload, headers=headers)
        
        # 检查响应
        response.raise_for_status()
        
        return response.json()
    except requests.RequestException as e:
        print(f"[ERROR] 更新墨墨生词本失败: {e}")
        return None

def save_words_to_file(word_data, filename="words_data.txt"):
    """将单词列表保存到文件中"""
    try:
        with open(filename, "w", encoding="utf-8") as file:
            file.write(generate_word_output(word_data))
        return True
    except Exception as e:
        print(f"[ERROR] 保存单词列表到文件失败: {e}")
        return False

def main():
    start_time = datetime.now()
    print(f"[INFO] 开始同步 - {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 获取欧路单词
    print("[INFO] 正在获取欧路词典单词...")
    word_data = fetch_word_list()
    
    if word_data:
        # 保存单词列表到文件
        word_count = len(word_data.get('data', []))
        print(f"[INFO] 获取到 {word_count} 个单词，正在保存到本地文件...")
        save_words_to_file(word_data)
        
        # 生成输出并同步到墨墨
        output_string = generate_word_output(word_data)
        print("[INFO] 正在同步到墨墨背单词...")
        response = update_maimemo_notepad(output_string)
        
        if response and response.get('success'):
            print("[SUCCESS] 墨墨背单词同步完成!")
            # 同步成功通知
            show_windows_notification(
                title="单词同步成功",
                message=f"成功获取 {word_count} 个欧路单词，并同步到墨墨背单词！",
                is_success=True
            )
            
            # 询问是否运行Obsidian同步
            if ask_run_obsidian_sync():
                run_obsidian_sync_script()
            else:
                print("[INFO] 跳过Obsidian同步流程")
                show_windows_notification(
                    title="Obsidian同步已跳过",
                    message="你选择不将生词本添加到Obsidian中",
                    is_success=True
                )
            
        else:
            print("[ERROR] 墨墨背单词同步失败!")
            # 同步失败通知
            show_windows_notification(
                title="单词同步失败",
                message="更新墨墨生词本失败，请检查API密钥、生词本ID或网络！",
                is_success=False
            )
    else:
        print("[ERROR] 未获取到欧路词典单词，同步终止")
        # 未获取到单词通知
        show_windows_notification(
            title="单词同步失败",
            message="未获取到欧路词典单词，同步终止！请检查欧路API密钥或分类ID。",
            is_success=False
        )
    
    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds()
    print(f"\n[INFO] 同步结束 - {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"[INFO] 总耗时: {duration:.2f} 秒")

if __name__ == "__main__":
    main()