# -*- coding: utf-8 -*-
"""
open_wait.py — 打开文件并等待用户关闭（用于分类结果人工核对后继续上传）
=========================================================================
用法: python open_wait.py <文件路径>

行为:
  1. 用系统默认程序(Excel)打开该文件
  2. 轮询 Excel 锁文件 ~$<name> 的出现与消失来判断"文件已关闭"
  3. 用户关闭文件后即退出，流程继续（如上传）

兜底: 若 15 秒内未检测到锁文件（如只读打开 / 无锁模式），
      降级为"在 Excel 关闭文件后按 Enter 继续"，避免卡死。
"""
import sys
import os
import time


def main():
    if len(sys.argv) < 2:
        print("用法: open_wait.py <文件路径>")
        sys.exit(2)
    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"[错误] 文件不存在: {path}")
        sys.exit(1)

    d, fn = os.path.split(path)
    lock = os.path.join(d, '~$' + fn)

    print(f"已在 Excel 打开分类结果：\n  {path}")
    print("请核对 / 修改 / 保存后【关闭该文件】以继续上传……")
    os.startfile(path)

    # —— 等待 Excel 创建锁文件（~$xxx）——
    appeared = False
    for _ in range(15):
        if os.path.exists(lock):
            appeared = True
            break
        time.sleep(1)

    if not appeared:
        print("（未检测到 Excel 锁文件，可能是只读打开）")
        print("请在 Excel 中关闭该文件后，按 Enter 继续……")
        input()
        return

    # —— 等待锁文件消失（用户关闭文件，Excel 释放）——
    print("（已检测到文件被 Excel 打开，等待你关闭……）")
    waited = 0
    while os.path.exists(lock):
        time.sleep(1)
        waited += 1
        if waited >= 3600:  # 兜底：最长等 1 小时，超时改手动确认
            print("（等待超过 1 小时，改为手动确认）")
            print("请确认已在 Excel 中关闭该文件，按 Enter 继续……")
            input()
            return
    print("检测到文件已关闭，继续。")


if __name__ == '__main__':
    main()
