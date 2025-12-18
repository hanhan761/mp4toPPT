import cv2
import os
import numpy as np
import tempfile
import shutil
from pptx import Presentation
from pptx.util import Inches
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# 创建线程锁用于保护打印输出
print_lock = threading.Lock()

def dhash(image, hash_size=8):
    """
    计算图像的差异哈希（dHash）
    用于快速比较两张图片的相似度，解决重复图片和渐进动画问题
    
    原理：将图像缩放到 (hash_size+1) x hash_size，然后比较相邻像素的亮度差异
    返回：64位整数哈希值
    """
    # 将图像缩放到 (hash_size+1) x hash_size
    resized = cv2.resize(image, (hash_size + 1, hash_size))
    # 计算水平方向的差异：如果右边像素比左边亮则为1，否则为0
    diff = resized[:, 1:] > resized[:, :-1]
    # 将布尔数组转换为整数哈希值
    return sum([2 ** i for i, v in enumerate(diff.flatten()) if v])

def hamming_distance(hash1, hash2):
    """
    计算两个哈希值的汉明距离（不同位的数量）
    用于判断两张图片的相似度
    """
    return bin(hash1 ^ hash2).count('1')

def extract_slides_no_dedup(video_path, output_ppt_path, 
                            motion_threshold=5,     # 判定画面是否在动的阈值(%)
                            min_duration=1.0,       # 画面需要静止多久才保存(秒)
                            hash_threshold=5,      # 哈希去重阈值：汉明距离小于此值认为是重复
                            use_morphology=True,   # 是否使用形态学操作消除鼠标干扰
                            ignore_bottom_ratio=0.0,  # 忽略底部区域比例（0.0-1.0，如0.1表示忽略底部10%）
                            ignore_right_ratio=0.0   # 忽略右侧区域比例（0.0-1.0，如0.1表示忽略右侧10%）
                            ):
    """
    优化版PPT提取函数，解决鼠标干扰、重复图片和渐进动画问题
    
    核心优化：
    1. 图像哈希去重：使用dHash在保存前与上一张图片对比，避免重复保存
    2. 形态学操作：通过开运算消除鼠标指针等小物体，只关注大块内容变化
    3. ROI机制：可选择性忽略底部（任务栏）和右侧（摄像头）区域
    4. 两阶段处理：先宽松提取候选图片，后处理阶段进行去重
    """
    
    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    with print_lock:
        print(f"临时文件夹 created: {temp_dir}")
    
    try:
        cap = cv2.VideoCapture(video_path)
        if not cap.isOpened():
            with print_lock:
                print(f"错误：无法打开视频 {video_path}")
            return

        fps = cap.get(cv2.CAP_PROP_FPS)
        if fps <= 0: fps = 25 # 兜底防止读取不到fps
        
        # 采样处理间隔：每0.2秒检测一次画面状态（平衡性能与精度）
        process_interval = 0.2 
        frame_step = max(1, int(fps * process_interval))
        
        # 计算需要连续静止多少次循环才算"稳定"
        frames_needed_for_stable = int(min_duration / process_interval)

        slide_count = 0
        slide_images = []
        
        # 核心变量
        last_processed_gray = None   # 上一检测帧（用来判断屏幕还在动吗？）
        stable_counter = 0           # 计数器：画面已经连续静止了多少个周期
        last_saved_hash = None       # 上一张已保存图片的哈希值（用于去重）
        
        with print_lock:
            print(f"开始处理视频: {os.path.basename(video_path)}")
            print(f"视频FPS: {fps:.2f}")
            print(f"策略: 动画/动作停止后，静止 {min_duration} 秒即保存（带哈希去重）")
            if use_morphology:
                print(f"✓ 已启用形态学操作（消除鼠标干扰）")
            if ignore_bottom_ratio > 0 or ignore_right_ratio > 0:
                print(f"✓ ROI过滤: 底部{ignore_bottom_ratio*100:.0f}%, 右侧{ignore_right_ratio*100:.0f}%")

        while True:
            # 1. 快进读取（跳过中间帧，只读关键时间点）
            for _ in range(frame_step):
                cap.grab() 
            
            ret, frame = cap.read()
            if not ret:
                break

            # 2. 图像预处理
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            
            # 【优化1：ROI机制】选择性忽略底部和右侧区域（任务栏、摄像头等）
            h, w = gray.shape
            roi_top = 0
            roi_bottom = int(h * (1 - ignore_bottom_ratio))
            roi_left = 0
            roi_right = int(w * (1 - ignore_right_ratio))
            gray_roi = gray[roi_top:roi_bottom, roi_left:roi_right]
            
            # 高斯模糊去噪，防止视频压缩噪点导致误判为"在动"
            gray_roi = cv2.GaussianBlur(gray_roi, (21, 21), 0)
            
            # 【优化2：形态学操作】消除鼠标指针、激光笔等小物体干扰
            # 开运算（先腐蚀后膨胀）可以去除小的噪点，只保留大的内容块
            if use_morphology:
                kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
                gray_roi = cv2.morphologyEx(gray_roi, cv2.MORPH_OPEN, kernel)
            
            # 将处理后的ROI区域放回原图（用于后续保存完整帧）
            gray_processed = gray.copy()
            gray_processed[roi_top:roi_bottom, roi_left:roi_right] = gray_roi

            # 初始化第一帧
            if last_processed_gray is None:
                last_processed_gray = gray_processed
                continue

            # 3. 计算"动静"：当前帧 vs 0.2秒前的帧
            # 使用处理后的图像计算差异，这样鼠标移动不会影响判断
            frame_delta = cv2.absdiff(last_processed_gray, gray_processed)
            thresh = cv2.threshold(frame_delta, 25, 255, cv2.THRESH_BINARY)[1]
            
            # 计算变化区域占比 (%)
            motion_score = (np.count_nonzero(thresh) / thresh.size) * 100
            
            # 更新对比基准（使用处理后的图像）
            last_processed_gray = gray_processed

            # 4. 状态机逻辑
            if motion_score < motion_threshold:
                # 画面变化极小 -> 判定为"静止"
                stable_counter += 1
            else:
                # 画面变化较大 -> 判定为"运动中"（翻页、鼠标移动、动画播放中）
                # 只要在动，计数器就清零
                stable_counter = 0

            # 5. 触发保存（带哈希去重）
            # 当且仅当"刚达到"静止阈值时保存一次
            # 比如要求静止5次，当计数器由4变5时保存，变成6时就不保存了，避免同一页静止时重复存几百张
            if stable_counter == frames_needed_for_stable:
                # 【优化3：哈希去重】在保存前计算当前帧的哈希值，与上一张已保存的图片对比
                # 这样可以避免重复保存相同的页面，以及渐进动画导致的相似图片
                curr_hash = dhash(gray_processed)
                is_duplicate = False
                
                if last_saved_hash is not None:
                    hamming_dist = hamming_distance(curr_hash, last_saved_hash)
                    if hamming_dist < hash_threshold:
                        # 汉明距离小于阈值，认为是重复图片，跳过保存
                        is_duplicate = True
                        with print_lock:
                            print(f"【{os.path.basename(video_path)}】检测到重复帧（汉明距离={hamming_dist}），跳过保存")
                
                if not is_duplicate:
                    # 保存图片
                    img_path = os.path.join(temp_dir, f"slide_{slide_count:04d}.jpg")
                    cv2.imwrite(img_path, frame)
                    slide_images.append(img_path)
                    
                    # 更新哈希值
                    last_saved_hash = curr_hash
                    slide_count += 1
                    with print_lock:
                        print(f"【{os.path.basename(video_path)}】捕获第 {slide_count} 张 (于视频 {cap.get(cv2.CAP_PROP_POS_MSEC)/1000:.1f}秒处)")

        cap.release()
        
        # 6. 【优化4：后处理去重】对提取的候选图片进行二次去重
        # 解决渐进动画问题：如果图片A是图片B的子集（内容更少），则删除A，保留B
        original_count = len(slide_images)
        if len(slide_images) > 1:
            filtered_images = [slide_images[0]]  # 保留第一张
            for i in range(1, len(slide_images)):
                prev_img = cv2.imread(filtered_images[-1], cv2.IMREAD_GRAYSCALE)
                curr_img = cv2.imread(slide_images[i], cv2.IMREAD_GRAYSCALE)
                
                if prev_img is None or curr_img is None:
                    filtered_images.append(slide_images[i])
                    continue
                
                # 计算两张图片的哈希值
                prev_hash = dhash(prev_img)
                curr_hash = dhash(curr_img)
                hamming_dist = hamming_distance(prev_hash, curr_hash)
                
                # 如果相似度很高（可能是渐进动画），比较内容多少
                if hamming_dist < hash_threshold * 2:  # 使用更宽松的阈值
                    # 计算非零像素数量（粗略估计内容多少）
                    prev_content = np.count_nonzero(prev_img > 10)  # 忽略纯黑背景
                    curr_content = np.count_nonzero(curr_img > 10)
                    
                    # 如果当前图片内容明显更多（超过5%），替换上一张
                    if curr_content > prev_content * 1.05:
                        filtered_images[-1] = slide_images[i]  # 替换上一张
                        with print_lock:
                            print(f"【{os.path.basename(video_path)}】检测到渐进动画，保留内容更全的图片")
                    # 否则跳过当前图片（内容更少或相同）
                    else:
                        with print_lock:
                            print(f"【{os.path.basename(video_path)}】跳过内容较少的渐进动画帧")
                else:
                    # 内容差异较大，认为是新页面，保留
                    filtered_images.append(slide_images[i])
            
            slide_images = filtered_images
            with print_lock:
                print(f"【{os.path.basename(video_path)}】后处理去重完成：{len(slide_images)} 张（原始 {original_count} 张）")
        
        # 7. 生成PPT
        if not slide_images:
            with print_lock:
                print(f"【{os.path.basename(video_path)}】未提取到任何图片，请检查阈值设置。")
            return

        with print_lock:
            print(f"【{os.path.basename(video_path)}】共提取 {len(slide_images)} 张图片，开始生成PPT...")
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        
        for i, img_path in enumerate(slide_images):
            blank_slide_layout = prs.slide_layouts[6] 
            slide = prs.slides.add_slide(blank_slide_layout)
            
            # 读取图片并计算自适应尺寸
            img = cv2.imread(img_path)
            if img is None: continue
            
            img_height, img_width = img.shape[:2]
            slide_width = prs.slide_width
            slide_height = prs.slide_height
            
            img_ratio = img_width / img_height
            slide_ratio = slide_width / slide_height
            
            if img_ratio > slide_ratio:
                width = slide_width
                height = int(slide_width / img_ratio)
                left = 0
                top = (slide_height - height) // 2
            else:
                height = slide_height
                width = int(slide_height * img_ratio)
                top = 0
                left = (slide_width - width) // 2
            
            slide.shapes.add_picture(img_path, left, top, width, height)
            if (i+1) % 10 == 0:
                with print_lock:
                    print(f"【{os.path.basename(video_path)}】已处理 {i+1}/{len(slide_images)} 张...")
        
        prs.save(output_ppt_path)
        with print_lock:
            print(f"【{os.path.basename(video_path)}】成功！PPT已保存至: {output_ppt_path}")
        
    finally:
        # 清理临时文件
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
            with print_lock:
                print(f"【{os.path.basename(video_path)}】临时文件清理完成")

# ================= 配置区域 =================
# 支持的视频格式
VIDEO_EXTENSIONS = ['.mp4', '.avi', '.mov', '.mkv', '.flv', '.wmv', '.m4v']

# 输出文件夹名称
OUTPUT_FOLDER = "output_ppts"

# 获取当前脚本所在目录
current_dir = os.path.dirname(os.path.abspath(__file__)) if __file__ else os.getcwd()

# 创建输出文件夹
output_dir = os.path.join(current_dir, OUTPUT_FOLDER)
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
    print(f"已创建输出文件夹: {output_dir}")

# 查找所有视频文件
video_files = []
for file in os.listdir(current_dir):
    file_lower = file.lower()
    if any(file_lower.endswith(ext) for ext in VIDEO_EXTENSIONS):
        video_path = os.path.join(current_dir, file)
        video_files.append(video_path)

if not video_files:
    print(f"错误：在目录 {current_dir} 中未找到任何视频文件")
    print(f"支持的格式: {', '.join(VIDEO_EXTENSIONS)}")
else:
    print(f"找到 {len(video_files)} 个视频文件，开始多线程批量处理...\n")
    
    # 参数说明：
    # min_duration: 
    #   建议设为 1.5 到 2.0。
    #   意思是：当PPT动画播放完，或者翻页动作结束后，画面必须静止这么久，才会被抓取。
    #   这能有效防止抓取到动画的一半。
    
    # motion_threshold:
    #   建议设为 2 到 5。
    #   用来忽略视频压缩带来的轻微噪点。
    
    # 处理单个视频的包装函数
    def process_video(video_file, idx, total):
        video_name = os.path.basename(video_file)
        video_name_no_ext = os.path.splitext(video_name)[0]
        output_ppt = os.path.join(output_dir, f"{video_name_no_ext}.pptx")
        
        with print_lock:
            print(f"\n{'='*60}")
            print(f"线程开始处理 [{idx}/{total}]: {video_name}")
            print(f"{'='*60}")
        
        try:
            extract_slides_no_dedup(
                video_file, 
                output_ppt, 
                motion_threshold=0.2,   # 敏感度：越小越敏感（2%的变化就算动）
                min_duration=1.5,      # 静止时长：静止1.5秒后保存
                hash_threshold=5,       # 哈希去重阈值：汉明距离小于5认为是重复
                use_morphology=True,    # 启用形态学操作消除鼠标干扰
                ignore_bottom_ratio=0.0,  # 忽略底部区域比例（0=不忽略，0.1=忽略底部10%）
                ignore_right_ratio=0.0   # 忽略右侧区域比例（0=不忽略，0.1=忽略右侧10%）
            )
            with print_lock:
                print(f"✓ 完成 [{idx}/{total}]: {video_name}")
            return True, video_name
        except Exception as e:
            with print_lock:
                print(f"✗ 失败 [{idx}/{total}]: {video_name} - 错误: {str(e)}")
            return False, video_name
    
    # 使用多线程处理（线程数设置为CPU核心数，但不超过视频文件数量）
    import multiprocessing
    max_workers = min(multiprocessing.cpu_count(), len(video_files))
    print(f"使用 {max_workers} 个线程并行处理...\n")
    
    # 创建任务列表
    tasks = []
    for idx, video_file in enumerate(video_files, 1):
        tasks.append((video_file, idx, len(video_files)))
    
    # 使用线程池执行
    completed = 0
    failed = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # 提交所有任务
        future_to_video = {
            executor.submit(process_video, video_file, idx, len(video_files)): (idx, video_file)
            for idx, video_file in enumerate(video_files, 1)
        }
        
        # 等待所有任务完成
        for future in as_completed(future_to_video):
            idx, video_file = future_to_video[future]
            try:
                success, video_name = future.result()
                if success:
                    completed += 1
                else:
                    failed += 1
            except Exception as e:
                with print_lock:
                    print(f"✗ 异常 [{idx}/{len(video_files)}]: {os.path.basename(video_file)} - {str(e)}")
                failed += 1
    
    print(f"\n{'='*60}")
    print(f"批量处理完成！")
    print(f"成功: {completed} 个，失败: {failed} 个")
    print(f"所有PPT已保存至: {output_dir}")
    print(f"{'='*60}")