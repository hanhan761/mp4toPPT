import os
import re
from datetime import datetime, timedelta

# ================== 数据配置区域 ==================
RAW_TEXT = """
第9周 (共4节)
1. 结构力学Ⅱ 已结束 2025.11.03 周一 10:10 - 11:00
2. 结构力学Ⅱ 已结束 2025.11.03 周一 11:10 - 12:00
3. 结构力学Ⅱ 已结束 2025.11.06 周四 16:10 - 17:00
4. 结构力学Ⅱ 已结束 2025.11.06 周四 17:10 - 18:00

第10周 (共4节)
1. 结构力学Ⅱ 已结束 2025.11.10 周一 10:10 - 11:00
2. 结构力学Ⅱ 已结束 2025.11.10 周一 11:10 - 12:00
3. 结构力学Ⅱ 已结束 2025.11.13 周四 16:10 - 17:00
4. 结构力学Ⅱ 已结束 2025.11.13 周四 17:10 - 18:00

第11周 (共4节)
1. 结构力学Ⅱ 已结束 2025.11.17 周一 10:10 - 11:00
2. 结构力学Ⅱ 已结束 2025.11.17 周一 11:10 - 12:00
3. 结构力学Ⅱ 已结束 2025.11.20 周四 16:10 - 17:00
4. 结构力学Ⅱ 已结束 2025.11.20 周四 17:10 - 18:00

第12周 (共4节)
1. 结构力学Ⅱ 已结束 2025.11.24 周一 10:10 - 11:00
2. 结构力学Ⅱ 已结束 2025.11.24 周一 11:10 - 12:00
3. 结构力学Ⅱ 已结束 2025.11.27 周四 16:10 - 17:00
4. 结构力学Ⅱ 已结束 2025.11.27 周四 17:10 - 18:00

第13周 (共4节)
1. 结构力学Ⅱ 已结束 2025.12.01 周一 10:10 - 11:00
2. 结构力学Ⅱ 已结束 2025.12.01 周一 11:10 - 12:00
3. 结构力学Ⅱ 已结束 2025.12.04 周四 16:10 - 17:00
4. 结构力学Ⅱ 已结束 2025.12.04 周四 17:10 - 18:00

第14周 (共4节)
1. 结构力学Ⅱ 未开始 2025.12.08 周一 10:10 - 11:00
2. 结构力学Ⅱ 未开始 2025.12.08 周一 11:10 - 12:00
3. 结构力学Ⅱ 未开始 2025.12.11 周四 16:10 - 17:00
4. 结构力学Ⅱ 未开始 2025.12.11 周四 17:10 - 18:00
"""

FILE_PREFIX = "2-13325"
EARLY_MINUTES = 10
# 设定阈值：800MB (1 MB = 1024 * 1024 bytes)
MIN_SIZE_MB = 800
MIN_SIZE_BYTES = MIN_SIZE_MB * 1024 * 1024

# ================== 核心逻辑 ==================

def get_file_size_mb(filepath):
    """获取文件大小并转换为MB"""
    try:
        size = os.path.getsize(filepath)
        return size / (1024 * 1024)
    except OSError:
        return 0

def main():
    pattern = re.compile(r'(\d{4}\.\d{2}\.\d{2}).*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})')
    matches = pattern.findall(RAW_TEXT)
    current_files = set(os.listdir('.'))
    
    missing_list = []
    suspicious_list = []
    
    now = datetime.now()
    count_checked = 0
    
    print(f"{'='*15} 开始检查 (阈值: {MIN_SIZE_MB}MB) {'='*15}")

    for date_str, start_str, end_str in matches:
        start_dt = datetime.strptime(f"{date_str} {start_str}", "%Y.%m.%d %H:%M")
        end_dt = datetime.strptime(f"{date_str} {end_str}", "%Y.%m.%d %H:%M")
        
        if start_dt > now:
            continue
            
        count_checked += 1
        rec_start = start_dt - timedelta(minutes=EARLY_MINUTES)
        rec_end = end_dt
        fmt = "%Y%m%d%H%M"
        filename = f"{FILE_PREFIX}-{rec_start.strftime(fmt)}-{rec_end.strftime(fmt)}.mp4"
        
        # 检查逻辑
        if filename not in current_files:
            # 1. 彻底缺失
            missing_list.append(filename)
            print(f"[缺失] {date_str} {start_str} | {filename}")
        else:
            # 2. 文件存在，检查大小
            size_mb = get_file_size_mb(filename)
            if size_mb <= MIN_SIZE_MB:
                # 大小存疑
                suspicious_list.append((filename, size_mb))
                print(f"[存疑] {date_str} {start_str} | {filename}")
                print(f"       -> 大小仅为: {size_mb:.2f} MB (小于 {MIN_SIZE_MB} MB)")
            else:
                # print(f"[正常] {filename} ({size_mb:.2f} MB)") # 正常文件不刷屏
                pass

    print(f"\n{'='*15} 检查总结 {'='*15}")
    print(f"应有课程: {count_checked} 节")
    print(f"完全缺失: {len(missing_list)} 个")
    print(f"大小存疑: {len(suspicious_list)} 个")

    # 生成报告文件
    if missing_list or suspicious_list:
        with open("check_report.txt", "w", encoding='utf-8') as f:
            f.write("=== 缺失文件列表 (需下载) ===\n")
            for item in missing_list:
                f.write(f"{item}\n")
            
            f.write("\n=== 存疑文件列表 (需检查/重下) ===\n")
            for item, size in suspicious_list:
                f.write(f"{item}  [当前大小: {size:.2f} MB]\n")
                
        print(f"\n详细报告已生成至: check_report.txt")
    else:
        print("\n完美！所有文件均存在且大小符合要求。")

if __name__ == "__main__":
    main()