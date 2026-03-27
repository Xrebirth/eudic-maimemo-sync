import json
import datetime
import os
import random
import string
import sys

# ====================== 核心配置项（请修改路径） ======================
# 欧路生词本文件路径
VOCAB_FILE_PATH = r"C:\Program Files\DevelopmentTools\Automation\eudic-maimemo-sync\words_data.txt"
# 单词Canvas文件路径（单独存放单词）
WORD_CANVAS_PATH = r"D:\Note\收集箱\单词本.canvas"
# 词组Canvas文件路径（单独存放词组）
PHRASE_CANVAS_PATH = r"D:\Note\收集箱\词组本.canvas"
# =========================================================================

# 布局配置
INIT_X = -1000    # 初始X坐标
INIT_Y = -1000    # 初始Y坐标
X_OFFSET = 450    # 横向间距
Y_OFFSET = 350    # 纵向间距
NODE_WIDTH = 400  # 节点宽度
NODE_HEIGHT = 300 # 节点高度
MAX_PER_ROW = 3   # 每行最多排3个

def show_windows_notification(title, message):
    """
    使用plyer实现Windows系统通知（彻底解决WPARAM兼容问题）
    :param title: 通知标题
    :param message: 通知内容
    """
    try:
        # 优先使用plyer库（跨平台、更稳定）
        from plyer import notification
        notification.notify(
            title=title,
            message=message,
            app_name="欧路生词同步工具",  # 通知栏显示的应用名称
            timeout=10,                   # 通知显示时长（秒）
            ticker="欧路生词同步提醒"      # 通知滚动提示（可选）
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
                app_name="欧路生词同步工具",
                timeout=10
            )
        except Exception as e:
            # 终极兜底：打印通知内容到控制台
            print(f"\n【系统通知】{title}：{message}")
            print(f"[WARNING] 通知弹窗失败：{e}")
    except Exception as e:
        # 捕获其他异常，保证程序不崩溃
        print(f"\n通知显示失败：{e}")
        print(f"【系统通知】{title}：{message}")

def clean_text(text: str) -> str:
    """
    清理文本空格：
    1. 单词：去除所有多余空格（包括中间），仅保留纯单词
    2. 词组：仅去除首尾空格，中间空格保留（且合并为单个空格）
    """
    # 先去除首尾空格
    stripped = text.strip()
    # 合并中间多个空格为单个（处理"a   lot   of"→"a lot of"）
    cleaned = ' '.join(stripped.split())
    return cleaned

def classify_word_phrase(text: str) -> tuple:
    """
    分类单词/词组，返回 (类型, 清理后的文本)
    类型："word"（单词） / "phrase"（词组）
    判断规则：清理后文本包含空格 → 词组；否则 → 单词
    """
    cleaned_text = clean_text(text)
    # 空文本直接跳过
    if not cleaned_text:
        return (None, None)
    # 判断是否包含空格（多个单词分隔）
    if ' ' in cleaned_text:
        return ("phrase", cleaned_text)
    else:
        return ("word", cleaned_text)

def extract_and_classify_vocab(file_path: str) -> tuple:
    """
    提取生词本并分类：返回 (单词列表, 词组列表)
    低内存逐行读取，自动去重、清理空格、分类
    """
    words = []
    phrases = []
    if not os.path.exists(file_path):
        print(f"错误：生词本文件 {file_path} 不存在！")
        return (words, phrases)

    # 逐行读取，避免大文件占用内存
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            stripped_line = line.strip()
            # 跳过空行、标题行、日期行
            if not stripped_line or stripped_line.startswith("#"):
                continue
            
            # 分类并清理
            item_type, cleaned_item = classify_word_phrase(stripped_line)
            if item_type == "word" and cleaned_item not in words:
                words.append(cleaned_item)
            elif item_type == "phrase" and cleaned_item not in phrases:
                phrases.append(cleaned_item)
    
    return (words, phrases)

def generate_unique_id():
    """生成绝对唯一的ID：微秒级时间戳 + 4位随机字符"""
    time_part = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
    random_part = ''.join(random.choices(string.ascii_letters + string.digits, k=4))
    return f"{time_part}{random_part}"

def load_canvas_data(file_path: str) -> dict:
    """加载Canvas文件，兼容文件不存在/格式错误"""
    default_canvas = {
        "nodes": [],
        "edges": [],
        "metadata": {"version": "1.0-1.0", "frontmatter": {}}
    }
    if not os.path.exists(file_path):
        print(f"Canvas文件不存在，将创建新文件：{file_path}")
        return default_canvas
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError:
        print(f"错误：{os.path.basename(file_path)} 格式损坏，使用初始结构覆盖！")
        return default_canvas
    except Exception as e:
        print(f"读取{os.path.basename(file_path)}失败：{str(e)}")
        return default_canvas

def get_existing_items(canvas_data: dict) -> set:
    """提取Canvas中已存在的单词/词组（用于去重）"""
    existing = set()
    for node in canvas_data.get("nodes", []):
        if node.get("type") == "text":
            # 提取第一行作为标识（单词/词组）
            item = node["text"].split("\n")[0].strip()
            existing.add(item)
    return existing

def generate_grid_node(item: str, node_index: int, start_x: int, start_y: int) -> dict:
    """生成网格排列的节点（兼容单词/词组）"""
    node_id = generate_unique_id()
    # 固定文本格式，替换为清理后的单词/词组
    node_text = f"""{item}  
*形态1,形态2*  
**释义：** 无  
**例句：** 无  """
    
    # 计算行列坐标
    col = node_index % MAX_PER_ROW
    row = node_index // MAX_PER_ROW
    new_x = start_x + (col * X_OFFSET)
    new_y = start_y + (row * Y_OFFSET)
    
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

def calculate_start_coords(nodes: list) -> tuple:
    """计算新增节点的起始坐标（基于最后一个节点）"""
    if not nodes:
        return (INIT_X, INIT_Y)
    
    last_node = nodes[-1]
    last_col = (len(nodes)-1) % MAX_PER_ROW
    if last_col < MAX_PER_ROW - 1:
        # 同一行还有位置
        start_x = last_node["x"] + X_OFFSET
        start_y = last_node["y"]
    else:
        # 换行
        start_x = INIT_X
        start_y = last_node["y"] + Y_OFFSET
    return (start_x, start_y)

def save_canvas_data(file_path: str, canvas_data: dict):
    """保存Canvas文件"""
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(canvas_data, f, ensure_ascii=False, indent=4)
    print(f"✅ {os.path.basename(file_path)} 已保存")

def process_vocab(canvas_path: str, items: list, item_type: str) -> int:
    """
    处理单词/词组并写入对应Canvas，返回新增数量
    :param canvas_path: Canvas文件路径
    :param items: 单词/词组列表
    :param item_type: "word" / "phrase"
    :return: 新增的数量
    """
    if not items:
        print(f"\n=== 无新增{('单词' if item_type=='word' else '词组')} ===")
        return 0
    
    # 加载Canvas
    canvas_data = load_canvas_data(canvas_path)
    # 获取已存在的项（去重）
    existing_items = get_existing_items(canvas_data)
    # 过滤新增项
    items_to_add = [item for item in items if item not in existing_items]
    
    if not items_to_add:
        print(f"\n=== 所有{('单词' if item_type=='word' else '词组')}已存在，无需新增 ===")
        return 0
    
    print(f"\n=== 待新增{('单词' if item_type=='word' else '词组')}：{items_to_add} ===")
    nodes = canvas_data["nodes"]
    # 计算起始坐标
    start_x, start_y = calculate_start_coords(nodes)
    
    # 生成节点
    for idx, item in enumerate(items_to_add):
        new_node = generate_grid_node(item, idx, start_x, start_y)
        nodes.append(new_node)
        print(f"✅ 已添加{('单词' if item_type=='word' else '词组')}：{item} (ID：{new_node['id']})")
    
    # 保存文件
    save_canvas_data(canvas_path, canvas_data)
    return len(items_to_add)

def main():
    print("=== 开始提取并分类欧路生词 ===")
    # 1. 提取并分类单词/词组
    words, phrases = extract_and_classify_vocab(VOCAB_FILE_PATH)
    print(f"提取结果：单词{len(words)}个 → {words}")
    print(f"          词组{len(phrases)}个 → {phrases}")

    # 2. 处理单词（写入单词本），记录新增数量
    new_word_count = process_vocab(WORD_CANVAS_PATH, words, "word")
    
    # 3. 处理词组（写入词组本），记录新增数量
    new_phrase_count = process_vocab(PHRASE_CANVAS_PATH, phrases, "phrase")

    # 4. 显示Windows系统通知
    total_new = new_word_count + new_phrase_count
    notification_title = "欧路生词同步完成"
    if total_new == 0:
        notification_msg = "本次无新增单词/词组"
    else:
        notification_msg = f"新增单词：{new_word_count}个\n新增词组：{new_phrase_count}个\n总计：{total_new}个"
    
    show_windows_notification(notification_title, notification_msg)
    print(f"\n=== 所有任务完成 {notification_msg} ===")

if __name__ == "__main__":
    main()