import subprocess
import os
import time
import sys
import shlex

results = []  # 保存每一步执行结果

def run_command(step, total, description, cmd, input_data=None):
    """运行命令并实时输出日志；自动识别 .py/.exe/.bat 等类型并处理"""
    print(f"\n[步骤 {step}/{total}] {description}")
    cmd_list = cmd if isinstance(cmd, list) else [cmd]
    if len(cmd_list) >= 2 and os.path.basename(cmd_list[0]).lower() in (
        os.path.basename(sys.executable).lower(), "python", "python.exe"
    ):
        target = cmd_list[1]
    else:
        target = cmd_list[0]

    target = os.path.abspath(target)
    exe_dir = os.path.dirname(target) or os.getcwd()

    # 文件存在检查
    if not os.path.exists(target):
        print(f"[WARN] 文件不存在: {target}，跳过此步骤")
        results.append((step, description, "❌ 缺少文件"))
        return

    lower = target.lower()
    actual_cmd = None
    use_shell = False

    if lower.endswith(".py"):
        extra_args = cmd_list[2:] if (len(cmd_list) >= 2 and os.path.basename(cmd_list[0]).lower() in (os.path.basename(sys.executable).lower(), "python", "python.exe")) else cmd_list[1:]
        actual_cmd = [sys.executable, target] + extra_args
    elif lower.endswith(".bat"):
        args = cmd_list[1:] if (len(cmd_list) >= 2 and os.path.basename(cmd_list[0]).lower() in (os.path.basename(sys.executable).lower(), "python", "python.exe")) else cmd_list[1:]
        cmd_str = " ".join([shlex.quote(target)] + [shlex.quote(a) for a in args])
        actual_cmd = cmd_str
        use_shell = True
    else:
        actual_cmd = cmd_list
        use_shell = False

    start_time = time.time()
    try:
        if os.path.basename(target).lower() == "rsas2check.exe":
            os.makedirs(os.path.join(exe_dir, "combined_reports"), exist_ok=True)
        if use_shell:
            proc = subprocess.Popen(
                actual_cmd,
                cwd=exe_dir,
                stdin=subprocess.PIPE if input_data else None,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                shell=True,
                bufsize=1
            )
        else:
            proc = subprocess.Popen(
                actual_cmd,
                cwd=exe_dir,
                stdin=subprocess.PIPE if input_data else None,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                shell=False,
                bufsize=1
            )

        if input_data:
            try:
                proc.stdin.write(input_data)
                proc.stdin.flush()
                proc.stdin.close()
            except Exception:
                pass
        for line in proc.stdout:
            sys.stdout.write(line)
            sys.stdout.flush()

        proc.wait()
        if proc.returncode == 0:
            results.append((step, description, "✅ 成功"))
        else:
            results.append((step, description, f"❌ 执行失败 (code={proc.returncode})"))

    except FileNotFoundError as e:
        print(f"[ERROR] 可执行文件未找到: {e}")
        results.append((step, description, f"❌ 缺少文件 {e}"))
    except Exception as e:
        print(f"[ERROR] 执行出错: {e}")
        results.append((step, description, f"❌ 出错 {e}"))
    finally:
        cost = round(time.time() - start_time, 2)
        print(f"[INFO] 步骤 {step} 耗时 {cost} 秒")


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    total_steps = 7

    steps = [
        (1, "运行 move.py", [sys.executable, os.path.join(base_dir, "文件", "move.py")]),
        # 下面即便指向 .py，run_command 会自动用 python 去执行
        (2, "运行 nessus.bat", [os.path.join(base_dir, "整理结果", "nessus", "nessus.py")]),
        (3, "运行 AwvsReport.exe", [os.path.join(base_dir, "整理结果", "awvs", "AwvsReport.py")]),
        (4, "运行 awvs.py", [sys.executable, os.path.join(base_dir, "整理结果", "awvs", "awvs.py")]),
        (5, "运行 rsas2check.exe (请手动回车选择默认模式)", [os.path.join(base_dir, "整理结果", "RSAS", "rsas2check.py")]),
        (6, "运行 rsas.py", [sys.executable, os.path.join(base_dir, "整理结果", "RSAS", "rsas.py")]),
        (7, "运行 nmap.bat", [os.path.join(base_dir, "整理结果", "nmap", "nmap.py")])
    ]

    for step in steps:
        if len(step) == 3:
            run_command(step[0], total_steps, step[1], step[2])
        elif len(step) == 4:
            run_command(step[0], total_steps, step[1], step[2], input_data=step[3])

    # 执行结果总结
    print("\n========== 执行结果总结 ==========")
    for step, desc, status in results:
        print(f"[步骤 {step}] {desc} -> {status}")
    print("================================")

    print("\n[ALL DONE] 所有任务已执行完毕 ✅")


if __name__ == "__main__":
    main()
