#!/usr/bin/env python3
"""发票自动处理工具 — 主入口"""
import argparse
import logging
import os
import sys
import yaml

def load_config(config_path="config.yaml"):
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def setup_logging(config):
    level = getattr(logging, config.get("logging", {}).get("level", "INFO").upper(), logging.INFO)
    log_file = config.get("logging", {}).get("log_file", "invoice_processor.log")
    handlers = [logging.StreamHandler(sys.stdout)]
    if log_file:
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    logging.basicConfig(level=level, format="%(asctime)s [%(levelname)s] %(message)s", handlers=handlers)

def main():
    parser = argparse.ArgumentParser(description="发票自动处理工具")
    parser.add_argument("--once", action="store_true", help="处理完已有文件后退出（不持续监控）")
    parser.add_argument("--config", default="config.yaml", help="配置文件路径")
    args = parser.parse_args()

    config = load_config(args.config)
    setup_logging(config)

    # 自动创建目录
    for key in ["watch_dir", "output_dir", "archive_dir"]:
        os.makedirs(config["paths"][key], exist_ok=True)

    from invoice_processor.watcher import process_existing_files, start_watching

    if args.once:
        process_existing_files(config)
    else:
        start_watching(config)

if __name__ == "__main__":
    main()
