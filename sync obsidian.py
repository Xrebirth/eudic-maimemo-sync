import json
import datetime
import os
import random
import string
import sys
import requests
from collections import defaultdict
from dotenv import load_dotenv
from datetime import datetime, timezone, timedelta

# ====================== 核心配置区（请根据实际情况修改） ======================
# 欧路API配置（需在.env文件中配置 EUDIC_API_KEY、EUDIC_CATEGORY_ID）
# 墨墨API配置（需在.env文件中配置 MOMO_API_KEY、MOMO_NOTEPAD_ID）
# 文件路径配置
VOCAB_FILE_PATH = r"C:\Program Files\DevelopmentTools\Automation\eudic-maimemo-sync - 副本\words_data.txt"
WORD_CANVAS_PATH = r"D:\Note\资源箱\附件\生词本\单词本.canvas"
PHRASE_CANVAS_PATH = r"D:\Note\资源箱\附件\生词本\词组本.canvas"
# Obsidian Canvas布局配置
INIT_X = -1000    # 初始X坐标
INIT_Y = -1000    # 初始Y坐标
X_OFFSET = 450    # 横向间距
Y_OFFSET = 350    # 纵向间距
NODE_WIDTH = 400  # 节点宽度
NODE_HEIGHT = 300 # 节点高度
MAX_PER_ROW = 3   # 每行最多排3个
# =========================================================================

# -------------------------- 统一的Windows通知函数 --------------------------
def show_windows_notification(title, message, is_success=True):
    """
    使用plyer实现Windows系统通知（解决兼容问题）
    :param title: 通知标题
    :param message: 通知内容
    :param is_success: 是否为成功类通知（仅用于标识）
    """
    try:
        from plyer import notification
        notification.notify(
            title=title,
            message=message,
            app_name="欧路-墨墨-Obsidian同步工具",
            timeout=10,
            ticker="单词同步提醒"
        )
    except ImportError:
        # 自动安装依赖并重试
        print("\n[INFO] 正在安装plyer依赖库...")
        os.system(f"{sys.executable} -m pip install plyer -q")
        try:
            from plyer import notification
            notification.notify(
                title=title,
                message=message,
                app_name="欧路-墨墨-Obsidian同步工具",
                timeout=10
            )
        except Exception as e:
            print(f"\n[WARNING] 通知弹窗失败：{e}")
            print(f"【通知】{title}：{message}")
    except Exception as e:
        print(f"\n[ERROR] 通知展示失败：{e}")
        print(f"【通知】{title}：{message}")

# -------------------------- 欧路单词获取 & 墨墨同步 相关函数 --------------------------
def fetch_word_list():
    """获取欧路词典生词本（从API）"""
    load_dotenv()
    
    headers = {
        "Authorization": os.getenv("EUDIC_API_KEY"),
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    }
    
    url = f"https://api.frdic.com/api/open/v1/studylist/words/{os.getenv('EUDIC_CATEGORY_ID')}"

    try:
        response = requests.get(url, headers=headers, params={"language": "en"})
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        print(f"[ERROR] 获取欧路单词列表失败: {e}")
        show_windows_notification(
            title="欧路单词获取失败",
            message=f"获取欧路单词列表失败：{str(e)}",
            is_success=False
        )
        return None

def generate_word_output(word_data):
    """生成按日期分组的单词字符串（UTC转中国时间）"""
    if not word_data or 'data' not in word_data:
        return ""

    china_tz = timezone(timedelta(hours=8))
    grouped_words = defaultdict(list)
    
    for item in word_data['data']:
        # UTC时间转中国时间
        utc_time = datetime.fromisoformat(item["add_time"].replace('Z', '+00:00'))
        china_time = utc_time.astimezone(china_tz)
        date = china_time.strftime("%Y-%m-%d")
        grouped_words[date].append(item["word"])

    output_string = ""
    for date in sorted(grouped_words.keys()):
        output_string += f"#{date}\n"
        output_string += "\n".join(grouped_words[date])
        output_string += "\n"

    return output_string

def save_words_to_file(word_data):
    """将欧路单词列表保存到本地文件"""
    try:
        output_str = generate_word_output(word_data)
        with open(VOCAB_FILE_PATH, "w", encoding="utf-8") as file:
            file.write(output_str)
        print(f"[SUCCESS] 欧路单词已保存到 {VOCAB_FILE_PATH}")
        return True
    except Exception as e:
        error_msg = f"保存单词文件失败：{str(e)}"
        print(f"[ERROR] {error_msg}")
        show_windows_notification(
            title="单词文件保存失败",
            message=error_msg,
            is_success=False
        )
        return False

def update_maimemo_notepad(word_data):
    """同步欧路单词到墨墨背单词"""
    load_dotenv()
    api_key = os.getenv("MOMO_API_KEY")
    notepad_id = os.getenv("MOMO_NOTEPAD_ID")
    
    if not api_key or not notepad_id:
        error_msg = "墨墨API密钥或生词本ID未配置（请检查.env文件）"
        print(f"[ERROR] {error_msg}")
        show_windows_notification(
            title="墨墨同步失败",
            message=error_msg,
            is_success=False
        )
        return False

    url = f"https://open.maimemo.com/open/api/v1/notepads/{notepad_id}"
    headers = {
        'Accept': 'application/json',
        'Authorization': f'Bearer {api_key}',
        'Content-Type': 'application/json'
    }
    payload = {
        "notepad": {
            "status": "UNPUBLISHED",
            "content": generate_word_output(word_data),
            "title": "欧路生词",
            "brief": "欧路词典上查询过的单词",
            "tags": ["自用"]
        }
    }

    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        if response.json().get('success'):
            print("[SUCCESS] 墨墨背单词同步完成!")
            return True
        else:
            raise Exception("墨墨API返回非成功状态")
    except requests.RequestException as e:
        error_msg = f"墨墨同步失败：{str(e)}"
        print(f"[ERROR] {error_msg}")
        show_windows_notification(
            title="墨墨同步失败",
            message=error_msg,
            is_success=False
        )
        return False

# -------------------------- Obsidian Canvas同步 相关函数 --------------------------
def clean_text(text: str) -> str:
    """清理文本空格：单词去所有多余空格，词组仅去首尾+合并中间空格"""
    stripped = text.strip()
    cleaned = ' '.join(stripped.split())
    return cleaned

def classify_word_phrase(text: str) -> tuple:
    """分类单词/词组，返回 (类型, 清理后的文本)"""
    cleaned_text = clean_text(text)
    if not cleaned_text:
        return (None, None)
    if ' ' in cleaned_text:
        return ("phrase", cleaned_text)
    else:
        return ("word", cleaned_text)

def extract_and_classify_vocab() -> tuple:
    """提取生词本并分类：返回 (单词集合, 词组集合)"""
    words = set()
    phrases = set()
    if not os.path.exists(VOCAB_FILE_PATH):
        error_msg = f"生词本文件 {VOCAB_FILE_PATH} 不存在！"
        print(f"[ERROR] {error_msg}")
        show_windows_notification(
            title="Obsidian同步失败",
            message=error_msg,
            is_success=False
        )
        return (words, phrases)

    with open(VOCAB_FILE_PATH, "r", encoding="utf-8") as f:
        for line in f:
            stripped_line = line.strip()
            if not stripped_line or stripped_line.startswith("#"):
                continue
            
            item_type, cleaned_item = classify_word_phrase(stripped_line)
            if item_type == "word" and cleaned_item:
                words.add(cleaned_item)
            elif item_type == "phrase" and cleaned_item:
                phrases.add(cleaned_item)
    
    return (words, phrases)

def generate_unique_id():
    """生成唯一ID：微秒时间戳 + 4位随机字符"""
    time_part = datetime.now().strftime("%Y%m%d%H%M%S%f")
    random_part = ''.join(random.choices(string.ascii_letters + string.digits, k=4))
    return f"{time_part}{random_part}"

def load_canvas_data(file_path: str) -> dict:
    """加载Canvas文件（兼容不存在/格式错误）"""
    default_canvas = {
        "nodes": [],
        "edges": [],
        "metadata": {"version": "1.0-1.0", "frontmatter": {}}
    }
    if not os.path.exists(file_path):
        print(f"[INFO] Canvas文件不存在，将创建新文件：{file_path}")
        return default_canvas
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError:
        print(f"[WARNING] {os.path.basename(file_path)} 格式损坏，使用初始结构覆盖！")
        return default_canvas
    except Exception as e:
        print(f"[ERROR] 读取{os.path.basename(file_path)}失败：{str(e)}")
        return default_canvas

def get_existing_items(canvas_data: dict) -> tuple:
    """提取Canvas中已存在的单词/词组"""
    existing = set()
    item_nodes = []
    for node in canvas_data.get("nodes", []):
        if node.get("type") == "text":
            item = node["text"].split("\n")[0].strip()
            existing.add(item)
            item_nodes.append((item, node))
    return (existing, item_nodes)

def generate_grid_node(item: str, node_index: int) -> dict:
    """生成网格排列的Canvas节点"""
    node_id = generate_unique_id()
    # 替换为新的文本格式
    node_text = f"""{item}  
**释义** 
无 

---

**例句** 

无  

---

**笔记** 

无"""
    
    col = node_index % MAX_PER_ROW
    row = node_index // MAX_PER_ROW
    new_x = INIT_X + (col * X_OFFSET)
    new_y = INIT_Y + (row * Y_OFFSET)
    
    return {
        "id": node_id,
        "type": "text",
        "text": node_text,
        "styleAttributes": {},
        "x": new_x,
        "y": new_y,
        "width": NODE_WIDTH,
        "height": NODE_HEIGHT
    }


def reorder_nodes_to_grid(item_nodes: list) -> list:
    """将节点重新排列为规整的网格布局"""
    reordered_nodes = []
    for idx, (item, node) in enumerate(item_nodes):
        col = idx % MAX_PER_ROW
        row = idx // MAX_PER_ROW
        new_x = INIT_X + (col * X_OFFSET)
        new_y = INIT_Y + (row * Y_OFFSET)
        
        node["x"] = new_x
        node["y"] = new_y
        reordered_nodes.append(node)
    
    return reordered_nodes

def save_canvas_data(file_path: str, canvas_data: dict):
    """保存Canvas文件"""
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(canvas_data, f, ensure_ascii=False, indent=4)
    print(f"✅ {os.path.basename(file_path)} 已保存")

def process_vocab(canvas_path: str, current_items: set, item_type: str) -> tuple:
    """
    处理单词/词组并写入对应Canvas（新增+删除+重排）
    修复核心：确保reordered_nodes始终是节点字典列表，避免元组索引错误
    """
    canvas_data = load_canvas_data(canvas_path)
    existing_items, item_nodes = get_existing_items(canvas_data)
    
    # 1. 删除Canvas中不存在于欧路的项
    items_to_delete = existing_items - current_items
    delete_count = len(items_to_delete)
    
    if delete_count > 0:
        print(f"\n=== 待删除{('单词' if item_type=='word' else '词组')}：{sorted(items_to_delete)} ===")
        # 过滤掉需要删除的节点
        remaining_item_nodes = [
            (item, node) for item, node in item_nodes 
            if item not in items_to_delete
        ]
        # 重新排列剩余节点为网格布局（返回节点字典列表）
        reordered_nodes = reorder_nodes_to_grid(remaining_item_nodes)
        canvas_data["nodes"] = reordered_nodes
        
        for item in sorted(items_to_delete):
            print(f"❌ 已删除{('单词' if item_type=='word' else '词组')}：{item}")
    else:
        print(f"\n=== 无需要删除的{('单词' if item_type=='word' else '词组')} ===")
        remaining_item_nodes = item_nodes
        # 关键修复：即使无删除，也要调用reorder_nodes_to_grid，确保返回节点字典列表
        reordered_nodes = reorder_nodes_to_grid(remaining_item_nodes)
    
    # 2. 添加欧路中有但Canvas中没有的项
    items_to_add = current_items - existing_items
    add_count = len(items_to_add)
    
    if add_count > 0:
        print(f"\n=== 待新增{('单词' if item_type=='word' else '词组')}：{sorted(items_to_add)} ===")
        start_idx = len(reordered_nodes)
        
        # 临时存储新增节点，避免变量复用导致类型混淆
        new_nodes = []
        for idx, item in enumerate(sorted(items_to_add)):
            new_node = generate_grid_node(item, start_idx + idx)
            new_nodes.append(new_node)
            print(f"✅ 已添加{('单词' if item_type=='word' else '词组')}：{item} (ID：{new_node['id']})")
        
        # 合并现有节点和新增节点
        all_nodes = reordered_nodes + new_nodes
        # 构建(item_text, node)元组列表，用于重排
        all_item_nodes = [
            (node["text"].split("\n")[0].strip(), node) 
            for node in all_nodes
        ]
        # 重新排列所有节点（包含新增的），确保布局规整
        canvas_data["nodes"] = reorder_nodes_to_grid(all_item_nodes)
    else:
        print(f"\n=== 无新增{('单词' if item_type=='word' else '词组')} ===")
        canvas_data["nodes"] = reordered_nodes
    
    # 3. 保存文件
    if add_count > 0 or delete_count > 0:
        save_canvas_data(canvas_path, canvas_data)
    else:
        print(f"\n=== {('单词' if item_type=='word' else '词组')}无变更，无需保存 ===")
    
    return (add_count, delete_count)

def sync_to_obsidian():
    """Obsidian Canvas同步主逻辑"""
    print("\n=== 开始同步欧路生词到Obsidian Canvas ===")
    # 提取并分类单词/词组
    words_set, phrases_set = extract_and_classify_vocab()
    print(f"[INFO] 提取结果：单词{len(words_set)}个 → {sorted(words_set)}")
    print(f"[INFO]        词组{len(phrases_set)}个 → {sorted(phrases_set)}")

    # 处理单词
    new_word_count, del_word_count = process_vocab(WORD_CANVAS_PATH, words_set, "word")
    
    # 处理词组
    new_phrase_count, del_phrase_count = process_vocab(PHRASE_CANVAS_PATH, phrases_set, "phrase")

    # 统计结果
    total_new = new_word_count + new_phrase_count
    total_del = del_word_count + del_phrase_count
    notification_msg = (
        f"新增单词：{new_word_count}个 | 删除单词：{del_word_count}个\n"
        f"新增词组：{new_phrase_count}个 | 删除词组：{del_phrase_count}个\n"
        f"总计：新增{total_new}个 | 删除{total_del}个"
    )
    
    show_windows_notification(
        title="Obsidian Canvas同步完成",
        message=notification_msg,
        is_success=True
    )
    print(f"\n=== Obsidian同步完成 ===")
    print(f"📊 同步统计：{notification_msg}")

# -------------------------- 主流程函数 --------------------------
def main():
    start_time = datetime.now()
    print(f"[INFO] 开始全流程同步 - {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 步骤1：获取欧路单词
    print("\n=== 步骤1：获取欧路词典单词 ===")
    word_data = fetch_word_list()
    if not word_data:
        print("[ERROR] 欧路单词获取失败，同步终止")
        return
    
    word_count = len(word_data.get('data', []))
    print(f"[INFO] 成功获取 {word_count} 个欧路单词")

    # 步骤2：保存单词到本地文件
    print("\n=== 步骤2：保存单词到本地文件 ===")
    if not save_words_to_file(word_data):
        return

    # 步骤3：同步到墨墨背单词
    print("\n=== 步骤3：同步到墨墨背单词 ===")
    if not update_maimemo_notepad(word_data):
        # 墨墨同步失败仍继续Obsidian同步（可根据需要修改为return终止）
        print("[WARNING] 墨墨同步失败，继续执行Obsidian同步...")

    # 步骤4：同步到Obsidian Canvas
    print("\n=== 步骤4：同步到Obsidian Canvas ===")
    sync_to_obsidian()

    # 总耗时统计
    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds()
    print(f"\n[INFO] 全流程同步结束 - {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"[INFO] 总耗时: {duration:.2f} 秒")
    
    # 最终通知
    show_windows_notification(
        title="全流程同步完成",
        message=f"成功获取{word_count}个欧路单词，已保存到本地、同步到墨墨和Obsidian！总耗时{duration:.2f}秒",
        is_success=True
    )

if __name__ == "__main__":
    main()
