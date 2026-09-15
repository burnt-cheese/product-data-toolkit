#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
校准脚本（只读，不提交任何商品）
用环境变量 SITE_USER / SITE_PASS 登录站点，进入批量上传页，
抓取真实选择器与 DOM 结构，保存 HTML + 正文 + 截图，供回填 upload.py。
绝不点击“提交/上传”，不实际导入数据。
"""
import json
import os
import re
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent
CAL_DIR = BASE / "calibrate"
CAL_DIR.mkdir(exist_ok=True)
UPLOAD_URL = "https://admin-cn.example.com/goods/import-products"

USER = os.environ.get("SITE_USER")
PWD = os.environ.get("SITE_PASS")
if not (USER and PWD):
    sys.exit("请先设置环境变量 SITE_USER / SITE_PASS")


def main():
    report = {"url": UPLOAD_URL, "steps": []}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, channel="chrome")
        ctx = browser.new_context(accept_downloads=True)
        page = ctx.new_page()
        try:
            print(">> 打开:", UPLOAD_URL)
            page.goto(UPLOAD_URL, wait_until="networkidle")
            report["steps"].append("goto upload url")

            # 登录（站点登录表单字段已确认）
            if page.query_selector("#login_password"):
                print(">> 检测到登录页，自动登录")
                page.fill("#login_mobilePhone", USER)
                page.fill("#login_password", PWD)
                btn = page.locator("form button[type=submit]")
                if btn.count() == 0:
                    btn = page.locator("button").last
                btn.first.click()
                report["steps"].append("login clicked")
                # 等登录真正完成（离开 login 页）再继续
                try:
                    page.wait_for_function("() => !location.href.includes('login')", timeout=20000)
                except Exception:
                    pass
                page.wait_for_timeout(2000)
                # 登录后重新跳转导入页（带鉴权）
                print(">> 重新跳转导入页（带鉴权）")
                page.goto(UPLOAD_URL, wait_until="networkidle")
                page.wait_for_timeout(3000)

            final_url = page.url
            report["final_url"] = final_url
            print(">> 最终 URL:", final_url)

            # 抓取文件输入
            file_inputs = []
            for el in page.query_selector_all("input[type=file]"):
                file_inputs.append({
                    "id": el.get_attribute("id"),
                    "name": el.get_attribute("name"),
                    "accept": el.get_attribute("accept"),
                    "multiple": el.get_attribute("multiple"),
                })
            report["file_inputs"] = file_inputs
            report["on_upload_page"] = bool(file_inputs)
            print(">> 上传页文件输入存在:", bool(file_inputs))

            # 抓取可能的提交按钮
            submits = []
            for el in page.query_selector_all("button, input[type=submit], a.btn, .btn"):
                txt = (el.inner_text() or "").strip()
                if txt and any(k in txt for k in ["提交", "上传", "导入", "确定", "开始", "下一步", "确认"]):
                    submits.append({"text": re.sub(r"\s+", "", txt),
                                    "id": el.get_attribute("id"),
                                    "class": el.get_attribute("class")})
            report["submit_candidates"] = submits

            # 编码/格式说明文字
            body = page.inner_text("body") or ""
            (CAL_DIR / "body.txt").write_text(body, encoding="utf-8")
            kw = re.compile(r"utf|编码|csv|格式|gbk|ansi|乱码|必填|模板|示例|失败|成功|导入|下载", re.I)
            notes = [ln.strip() for ln in body.splitlines() if kw.search(ln)]
            report["format_notes"] = notes[:50]
            report["body_preview"] = body[:1200]

            # 保存 HTML + 截图
            html_path = CAL_DIR / "upload_page.html"
            html_path.write_text(page.content(), encoding="utf-8")
            shot = CAL_DIR / "upload_page.png"
            page.screenshot(path=str(shot), full_page=True)
            report["artifacts"] = {"html": str(html_path), "body": str(CAL_DIR / "body.txt"),
                                   "screenshot": str(shot)}

            # 保存会话（cookies + localStorage），便于复用
            cookies = ctx.cookies()
            ls = page.evaluate("() => JSON.stringify(localStorage)")
            sess = {"cookies": cookies, "localStorage": json.loads(ls)}
            (BASE / "session.json").write_text(json.dumps(sess, ensure_ascii=False), encoding="utf-8")
            report["session_saved"] = True
            report["cookie_names"] = [c["name"] for c in cookies]

            print("\n==== 校准报告 ====")
            print(json.dumps(report, ensure_ascii=False, indent=2))
        except Exception as e:
            print("校准异常:", e)
            import traceback
            traceback.print_exc()
            try:
                (CAL_DIR / "error_page.html").write_text(page.content(), encoding="utf-8")
                page.screenshot(path=str(CAL_DIR / "error_page.png"), full_page=True)
            except Exception:
                pass
            raise
        finally:
            browser.close()


if __name__ == "__main__":
    main()
