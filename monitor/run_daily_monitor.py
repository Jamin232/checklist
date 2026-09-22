#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_daily_monitor.py — 物流运营监控看板「每日刷新」编排脚本
=========================================================
职责:
  1. 自动定位 D:/素芸/跟踪表/ 下最新的《产品跟踪表》Excel（按修改时间取最新）。
  2. 调用 build_monitor_json.py 生成 data.json（原始行透传，前端负责计算）。
  3. 校验输出（记录数 > 阈值），并把当天 Excel 同步快照到 source/current.xlsx。
  4. 用 PortableGit 把站点文件推送到 GitHub Pages（monitor 子目录）。

本地用法（由 Windows 计划任务每日 17:00 触发）:
    python run_daily_monitor.py
    python run_daily_monitor.py --input "D:/素芸/跟踪表/0905 产品跟踪表.xlsx"
    python run_daily_monitor.py --dir "D:/素芸/跟踪表"

退出码: 0=成功, 1=未找到 Excel, 2=build 失败, 3=校验失败, 4=推送失败
"""

import argparse
import datetime
import glob
import json
import os
import subprocess
import sys

# ---- 路径配置（按需修改）----
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SEARCH_DIR = r"D:/素芸/跟踪表"
BUILD_SCRIPT = os.path.join(SCRIPT_DIR, "build_monitor_json.py")
DATA_JSON = os.path.join(SCRIPT_DIR, "data.json")
LAST_RUN = os.path.join(SCRIPT_DIR, "last_run.json")
DAILY_LOG = os.path.join(SCRIPT_DIR, "daily_log.txt")

# 用户机器未安装系统 Git，唯一可用的是 WorkBuddy 自带 PortableGit
GIT_EXE = r"C:\Users\yanmi\.workbuddy\binaries\PortableGit\versions\1.2.0\cmd\git.exe"
if not os.path.isfile(GIT_EXE):
    GIT_EXE = "git"  # 回退：希望系统 PATH 里有 git

# 候选 Excel 文件名匹配（命中任意一个即视为跟踪表）
NAME_HINTS = ["产品跟踪表", "跟踪表"]
# data.json 校验：记录数低于此值视为解析异常
MIN_ROWS = 100


def find_latest_excel(search_dir):
    if not os.path.isdir(search_dir):
        return None
    candidates = []
    for path in glob.glob(os.path.join(search_dir, "*.xlsx")):
        name = os.path.basename(path)
        if name.startswith("~$"):
            continue
        if any(h in name for h in NAME_HINTS):
            candidates.append(path)
    if not candidates:
        candidates = [p for p in glob.glob(os.path.join(search_dir, "*.xlsx"))
                     if not os.path.basename(p).startswith("~$")]
    if not candidates:
        return None
    return max(candidates, key=lambda p: os.path.getmtime(p))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", help="指定跟踪表 Excel（覆盖自动发现）")
    ap.add_argument("--dir", default=DEFAULT_SEARCH_DIR, help="搜索目录，默认 D:/素芸/跟踪表")
    ap.add_argument("--python", default=sys.executable, help="用于运行 build 的 python 解释器")
    args = ap.parse_args()

    started = datetime.datetime.now()
    log_lines = ["\n=== %s 开始每日刷新（monitor）===" % started.strftime("%Y-%m-%d %H:%M:%S")]

    # 1) 定位 Excel
    excel = args.input or find_latest_excel(args.dir)
    if not excel or not os.path.exists(excel):
        msg = "[FAIL] 未在 %s 找到跟踪表 Excel，也未通过 --input 指定。" % args.dir
        log_lines.append(msg)
        print(msg)
        _write_run(False, None, log_lines)
        sys.exit(1)
    log_lines.append("[INFO] 使用数据源: %s" % excel)

    # 2) 调用 build
    cmd = [args.python, BUILD_SCRIPT, "--input", excel, "--output", DATA_JSON]
    _sync_source_snapshot(excel, log_lines)
    log_lines.append("[INFO] 命令: %s" % " ".join(cmd))
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    for line in (proc.stdout or "").splitlines():
        log_lines.append("[build] " + line)
    if proc.returncode != 0:
        msg = "[FAIL] build_monitor_json.py 退出码 %d" % proc.returncode
        log_lines.append(msg)
        if proc.stderr:
            log_lines.append(proc.stderr.strip()[:2000])
        print(msg)
        _write_run(False, os.path.basename(excel), log_lines)
        sys.exit(2)

    # 3) 校验
    try:
        with open(DATA_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
        rows = data.get("rows", [])
        if len(rows) < MIN_ROWS:
            raise ValueError("记录数异常: %d（预期 >=%d），疑似解析失败。" % (len(rows), MIN_ROWS))
        meta = data.get("meta", {})
    except Exception as e:
        msg = "[FAIL] 校验失败: %s" % e
        log_lines.append(msg)
        print(msg)
        _write_run(False, os.path.basename(excel), log_lines)
        sys.exit(3)

    log_lines.append(
        "[OK] 记录 %d | 数据基准日 %s | 来源 %s"
        % (len(rows), meta.get("dataDate", "未知"), meta.get("sourceFile", "未知"))
    )
    log_lines.append("=== 完成 ===")
    ok = _git_push(log_lines)
    print("\n".join(log_lines[-8:]))
    _write_run(ok, os.path.basename(excel), log_lines,
               {"rows": len(rows), "dataDate": meta.get("dataDate", "")})
    sys.exit(0 if ok else 4)


def _sync_source_snapshot(excel, log_lines):
    """把当天使用的 Excel 复制到工作区 source/current.xlsx，供云端重部署自动化使用。"""
    try:
        src_dir = os.path.join(SCRIPT_DIR, "source")
        os.makedirs(src_dir, exist_ok=True)
        dst = os.path.join(src_dir, "current.xlsx")
        if os.path.abspath(excel) != os.path.abspath(dst):
            import shutil
            shutil.copy2(excel, dst)
            log_lines.append("[INFO] 已同步 Excel 快照 -> source/current.xlsx")
    except Exception as e:
        log_lines.append("[WARN] 同步 Excel 快照失败（不影响本地 data.json）: %s" % e)


def _git_push(log_lines):
    """推送站点静态文件到 GitHub Pages（若存在 origin 远程）。

    采用普通 commit + push（非 amend），保留历史；data.json 每日变化，
    git 会自动增量压缩，体积可控。仅推送站点必需文件（不含 xlsx / 脚本）。
    若未检测到 git 仓库或未配置 origin，则静默跳过（不影响本地 data.json 生成）。
    """
    try:
        git_cmd = [GIT_EXE]
        if subprocess.run(git_cmd + ["rev-parse", "--is-inside-work-tree"],
                          capture_output=True).returncode != 0:
            log_lines.append("[INFO] 未检测到 git 仓库，跳过 GitHub 推送（参见 README）。")
            return True
        if subprocess.run(git_cmd + ["remote", "get-url", "origin"],
                          capture_output=True).returncode != 0:
            log_lines.append("[INFO] 未配置 git origin 远程，跳过推送。")
            return True
        files = ["index.html", "data.json", "chayan.js", "daily.js", "chayan.css", ".nojekyll"]
        # 仅 add 存在的文件
        files = [f for f in files if os.path.isfile(os.path.join(SCRIPT_DIR, f))]
        subprocess.run(git_cmd + ["add"] + files, capture_output=True, text=True)
        if not subprocess.run(git_cmd + ["status", "--porcelain"],
                              capture_output=True, text=True).stdout.strip():
            log_lines.append("[INFO] 站点文件无变化，无需推送。")
            return True
        msg = "monitor %s" % datetime.datetime.now().strftime("%Y-%m-%d")
        c = subprocess.run(git_cmd + ["commit", "-m", msg], capture_output=True, text=True)
        if c.returncode != 0:
            log_lines.append("[WARN] git commit 失败: %s" % c.stderr.strip()[:200])
            return False
        p = subprocess.run(git_cmd + ["push", "origin", "HEAD"], capture_output=True, text=True)
        if p.returncode == 0:
            log_lines.append("[OK] 已推送站点更新到 GitHub Pages。")
            return True
        else:
            log_lines.append("[WARN] git push 失败（检查 SSH/凭证）: %s" % p.stderr.strip()[:200])
            return False
    except Exception as e:
        log_lines.append("[WARN] git 推送异常: %s" % e)
        return False


def _write_run(success, source, log_lines, summary=None):
    rec = {
        "success": success,
        "source": source,
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": summary,
    }
    try:
        with open(LAST_RUN, "w", encoding="utf-8") as f:
            json.dump(rec, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    try:
        with open(DAILY_LOG, "a", encoding="utf-8") as f:
            f.write("\n".join(log_lines) + "\n")
    except Exception:
        pass


if __name__ == "__main__":
    main()
