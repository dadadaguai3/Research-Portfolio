#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
预算AI分析系统 - Kimi版本 (Demo)
使用Kimi的文件上传功能，直接上传Excel、PDF等文件进行AI分析
"""

import os
import sys
import time
import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional
import json

# ==============================================================================
# 🎯 配置区域
# ==============================================================================

# Kimi API密钥
# 建议方式：将Key保存在环境变量 KIMI_API_KEY 中，或在下方引号内填入你的Key
# ⚠️ 注意：上传GitHub前请务必确保此处为空或使用环境变量！
KIMI_API_KEY = os.getenv("KIMI_API_KEY", "your_api_key_here")

# 输出目录 (使用相对路径，方便演示)
OUTPUT_DIR = os.path.join(os.getcwd(), "analysis_results")

# 要分析的文件夹列表 (示例路径)
# 实际使用时请修改为包含财政数据的文件夹路径
FOLDERS_TO_ANALYZE = [
    r"./data/sample_province/city_a",
    r"./data/sample_province/city_b",
    # r"C:\Users\YourName\Data\RealData\CityC"
]

# 提取的财政指标
PARAMETERS = [
    "财政事务",
    "税收事务",
    "审计事务",
    "财政事务——信息化建设",
    "税收事务——信息化建设",
    "审计事务——信息化建设",
]

# Kimi模型配置
KIMI_MODEL = "kimi-k2-turbo-preview"

# 文件数量限制配置
FILE_LIMIT_CONFIG = {
    "MAX_FILES": 1000,  # Kimi单用户最多1000个文件
    "MAX_SIZE_MB": 100,  # 单文件最大100MB
    "MAX_TOTAL_SIZE_GB": 10,  # 总容量最大10GB
    "WARNING_THRESHOLD": 0.8,  # 报警阈值
    "ACTION_ON_EXCEED": "skip",  # "skip"跳过 或 "stop"中断
    "COMPRESS_LARGE_PDF": True,  # 是否压缩大PDF文件
    "PDF_COMPRESS_THRESHOLD": 5.0,
    "PDF_COMPRESS_QUALITY": "medium",
}

# 是否在分析完成后删除上传的文件（节省空间）
DELETE_UPLOADED_FILES_AFTER_ANALYSIS = True

# API速率限制配置（根据Tier1账号限制）
RATE_LIMIT_CONFIG = {
    "TPM_LIMIT": 2000000,
    "RPM_LIMIT": 200,
    "RETRY_DELAY": 30,
    "MAX_RETRIES": 3,
    "ENABLE_RETRY": True,
}

# ==============================================================================
# 🚀 运行程序
# ==============================================================================

try:
    from openai import OpenAI
    import pandas as pd
    # 注意：如果使用了压缩功能，可能还需要导入其他库
except ImportError as e:
    print(f"❌ 缺少依赖库: {e}")
    print("请运行以下命令安装依赖:")
    print("pip install openai pandas")
    sys.exit(1)


class KimiBudgetAnalyzer:
    """基于Kimi的预算分析器"""

    def __init__(self):
        self.api_key = KIMI_API_KEY
        if "your_api_key_here" in self.api_key or not self.api_key:
            print("❌ 错误：未配置有效 API Key。请在代码中配置或设置环境变量 KIMI_API_KEY。")
            sys.exit(1)

        self.output_dir = OUTPUT_DIR
        self.parameters = PARAMETERS
        self.model = KIMI_MODEL
        self.file_limit_config = FILE_LIMIT_CONFIG

        # 创建输出目录
        os.makedirs(self.output_dir, exist_ok=True)

        self.client = None
        self.uploaded_files = []
        self.total_file_count = 0
        self.total_size_bytes = 0
        self.rate_limit_config = RATE_LIMIT_CONFIG
        self.last_request_time = 0
        self.conversation_history = {}
        self.compressed_files = []

    def initialize_client(self):
        """初始化Kimi客户端"""
        try:
            self.client = OpenAI(
                api_key=self.api_key,
                base_url="https://api.moonshot.cn/v1",
            )
            print("✅ Kimi客户端初始化成功")
            return True
        except Exception as e:
            print(f"❌ Kimi客户端初始化失败: {e}")
            return False

    def get_current_file_stats(self) -> Dict[str, Any]:
        """获取当前文件统计信息"""
        return {
            "uploaded_count": len(self.uploaded_files),
            "total_size_mb": self.total_size_bytes / (1024 * 1024),
            "max_files": self.file_limit_config["MAX_FILES"],
            "max_size_mb": self.file_limit_config["MAX_SIZE_MB"],
            "max_total_gb": self.file_limit_config["MAX_TOTAL_SIZE_GB"],
            "remaining_files": self.file_limit_config["MAX_FILES"] - len(self.uploaded_files),
            "remaining_size_mb": (self.file_limit_config["MAX_TOTAL_SIZE_GB"] * 1024) - (
                        self.total_size_bytes / (1024 * 1024)),
        }

    def initialize_conversation(self, city_name: str):
        if city_name not in self.conversation_history:
            base_system_prompt = """你是Kimi，由 Moonshot AI 提供的人工智能助手。"""
            self.conversation_history[city_name] = [
                {"role": "system", "content": base_system_prompt}
            ]
            # print(f"✅ 初始化 {city_name} 的对话上下文") # 减少日志输出

    def get_conversation_messages(self, city_name: str, max_history: int = 20) -> List[Dict]:
        if city_name not in self.conversation_history:
            self.initialize_conversation(city_name)
        messages = self.conversation_history[city_name]
        if len(messages) > max_history + 1:
            system_msg = messages[0]
            recent_messages = messages[-max_history:]
            self.conversation_history[city_name] = [system_msg] + recent_messages
            messages = self.conversation_history[city_name]
        return messages.copy()

    def add_to_conversation(self, city_name: str, message: Dict[str, str]):
        if city_name not in self.conversation_history:
            self.initialize_conversation(city_name)
        self.conversation_history[city_name].append(message)
        if len(self.conversation_history[city_name]) > 30:
            system_msg = self.conversation_history[city_name][0]
            recent_messages = self.conversation_history[city_name][-29:]
            self.conversation_history[city_name] = [system_msg] + recent_messages

    def check_file_limits(self, file_count: int, file_size_mb: float) -> tuple[bool, str]:
        stats = self.get_current_file_stats()
        if file_size_mb > self.file_limit_config["MAX_SIZE_MB"]:
            return False, f"文件大小 {file_size_mb:.1f}MB 超过限制"
        if stats["uploaded_count"] + file_count > self.file_limit_config["MAX_FILES"]:
            return False, f"文件数量将达到限制"
        new_total_size = stats["total_size_mb"] + file_size_mb
        if new_total_size > self.file_limit_config["MAX_TOTAL_SIZE_GB"] * 1024:
            return False, f"总容量将达到限制"
        return True, "检查通过"

    def handle_limit_exceeded(self, reason: str, file_info: Dict[str, Any] = None) -> bool:
        action = self.file_limit_config["ACTION_ON_EXCEED"]
        print(f"🚨 限制警告: {reason} -> 执行: {action}")
        return True if action == "skip" else False

    def handle_rate_limit(self, retry_count: int = 0) -> bool:
        if not self.rate_limit_config["ENABLE_RETRY"] or retry_count >= self.rate_limit_config["MAX_RETRIES"]:
            return False
        delay = self.rate_limit_config["RETRY_DELAY"]
        print(f"⏰ 触发速率限制，等待 {delay} 秒后重试...")
        time.sleep(delay)
        return True

    def upload_file(self, file_path: str) -> Optional[str]:
        try:
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            file_name = os.path.basename(file_path)

            # PDF压缩逻辑占位 (保留结构，简化依赖检查)
            is_compressed = False
            # ... (此处保留原有压缩逻辑结构，为代码简洁略去具体实现细节)

            can_upload, reason = self.check_file_limits(1, file_size_mb)
            if not can_upload:
                if not self.handle_limit_exceeded(reason): return None
                return "skipped"

            print(f"📤 上传中: {file_name}...")
            file_object = self.client.files.create(file=Path(file_path), purpose="file-extract")

            self.uploaded_files.append({
                "id": file_object.id, "name": file_name, "size_mb": file_size_mb,
                "id_compressed": is_compressed
            })
            self.total_file_count += 1
            self.total_size_bytes += os.path.getsize(file_path)
            return file_object.id
        except Exception as e:
            print(f"❌ 上传失败 {os.path.basename(file_path)}: {e}")
            return None

    def upload_files_batch(self, file_paths: List[str]) -> List[str]:
        uploaded_ids = []
        for path in file_paths:
            fid = self.upload_file(path)
            if fid and fid != "skipped": uploaded_ids.append(fid)
        return uploaded_ids

    def analyze_with_kimi(self, file_ids: List[str], city_name: str, year: str, description: str = "") -> Optional[
        Dict[str, Any]]:
        retry_count = 0
        while retry_count <= self.rate_limit_config["MAX_RETRIES"]:
            try:
                print(f"🤖 AI分析中: {city_name} {year} ({len(file_ids)} files)")
                messages = self.get_conversation_messages(city_name)

                # 优化: 仅在有文件ID时尝试获取内容
                valid_files = 0
                for file_id in file_ids:
                    try:
                        content = self.client.files.content(file_id=file_id).text
                        messages.append({"role": "system", "content": content})
                        valid_files += 1
                    except:
                        pass

                if valid_files == 0: return None

                system_prompt = f"""你是财政专家。请从文件中提取{city_name}{year}年的决算数：
{chr(10).join(self.parameters)}
请以JSON格式输出，key为指标名，value为数值(万元)，未找到填"未找到"。"""

                messages.append({"role": "system", "content": system_prompt})
                messages.append({"role": "user", "content": f"分析{city_name}{year}年数据并提取指标。"})

                completion = self.client.chat.completions.create(
                    model=self.model, messages=messages, temperature=0.1,
                    response_format={"type": "json_object"}
                )

                self.last_request_time = time.time()
                ai_result = completion.choices[0].message.content.strip()

                self.add_to_conversation(city_name, messages[-1])  # User
                self.add_to_conversation(city_name, {"role": "assistant", "content": ai_result})

                return {"ai_result": ai_result, "valid_files": valid_files}

            except Exception as e:
                if "rate limit" in str(e).lower():
                    if self.handle_rate_limit(retry_count):
                        retry_count += 1
                        continue
                print(f"❌ 分析出错: {e}")
                return None
        return None

    def parse_ai_result(self, ai_data: Dict[str, Any], city_name: str, year: str) -> Optional[Dict[str, Any]]:
        if not ai_data: return None
        result = {"年份": year, "城市": city_name, "状态": "成功"}
        try:
            data = json.loads(ai_data["ai_result"])
            for param in self.parameters:
                result[param] = data.get(param, "未找到")
        except:
            result["状态"] = "解析失败"
            for param in self.parameters: result[param] = "解析失败"
        return result

    def cleanup_uploaded_files(self):
        if not DELETE_UPLOADED_FILES_AFTER_ANALYSIS: return
        print("🧹 清理云端文件...")
        for f in self.uploaded_files:
            try:
                self.client.files.delete(file_id=f["id"])
            except:
                pass

    def analyze_folder(self, folder_path: str) -> Optional[List[Dict[str, Any]]]:
        folder_name = os.path.basename(folder_path)
        if not os.path.exists(folder_path):
            print(f"ℹ️ 路径不存在(演示模式): {folder_path}")
            return None

        # 简化的文件夹扫描逻辑
        results = []
        # 此处省略了复杂的递归扫描，实际运行时请确保目录结构正确
        # ...
        return results

    def save_results(self, results: List[Dict[str, Any]], folder_name: str = None) -> Optional[str]:
        if not results: return None
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{folder_name or 'Analysis'}_{timestamp}.xlsx"
        path = os.path.join(self.output_dir, filename)
        try:
            pd.DataFrame(results).to_excel(path, index=False)
            print(f"💾 保存成功: {path}")
            return path
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            return None


def main():
    print("🎯 预算AI分析系统 - Kimi Demo")

    analyzer = KimiBudgetAnalyzer()
    if not analyzer.initialize_client(): return

    all_results = []
    for folder in FOLDERS_TO_ANALYZE:
        res = analyzer.analyze_folder(folder)
        if res: all_results.extend(res)

    analyzer.cleanup_uploaded_files()
    print("Done.")


if __name__ == "__main__":
    main()