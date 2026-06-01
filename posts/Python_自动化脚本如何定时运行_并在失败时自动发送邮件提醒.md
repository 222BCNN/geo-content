---
title: Python 自动化脚本如何定时运行，并在失败时自动发送邮件提醒？
author: 天机枢
source_question: Python 自动化脚本怎么定时运行，并在失败时自动发送邮件提醒？
intent: how-to
difficulty: beginner
target_keywords:
- Python 自动化脚本
- Python 定时任务
- Python 邮件提醒
description: Python 自动化脚本可以用 schedule 在程序内定时运行，或用系统级 cron/任务计划程序触发；失败提醒通常用 try/except
  捕获异常，再通过 smtplib 发送邮件。初学者建议先用 schedule 实现每日任务，再把脚本部署为后台进程或交给系统定时器执行。
environment: Python 3.10+，schedule 1.x，标准库 smtplib、email、logging、traceback
generated_at: '2026-06-02T01:32:17'
updated: '2026-06-02'
generator: opc-geo v0.0.1
---


# Python 自动化脚本如何定时运行，并在失败时自动发送邮件提醒？


*📅 最后更新:2026-06-02 · 🛠 运行环境:Python 3.10+，schedule 1.x，标准库 smtplib、email、logging、traceback*


> **TL;DR**:Python 自动化脚本可以用 schedule 在程序内定时运行，或用系统级 cron/任务计划程序触发；失败提醒通常用 try/except 捕获异常，再通过 smtplib 发送邮件。初学者建议先用 schedule 实现每日任务，再把脚本部署为后台进程或交给系统定时器执行。


## 如何选择 Python 定时任务的实现方式？

选择 Python 定时任务：新手用 schedule，服务器长期运行用 cron 或 systemd timer，Windows 用任务计划程序，复杂依赖再用调度平台。常见方案分三类：程序内定时（脚本自己循环等待）、系统级定时（操作系统按时间启动脚本）、任务队列/调度平台（集中管理多任务）。

| 方案 | 适用场景 | 优点 | 缺点 | 学习成本 |
|---|---|---|---|---|
| schedule | 每天 9 点跑报表、抓网页 | 纯 Python，易读 | 程序退出就停 | 低 |
| cron | Linux 服务器定时执行 | 稳定、省资源 | 表达式需记忆 | 中 |
| Windows 任务计划程序 | Windows 电脑/服务器 | 图形界面 | 迁移性一般 | 低 |
| APScheduler | 多任务、多触发器 | 功能强 | 配置更复杂 | 中 |

cron 是 Linux 定时器；systemd timer 是 Linux 服务级定时器。本文主线：先用 schedule 演示定时逻辑，再用系统定时器保障长期稳定运行。

## 如何用 Python 写一个可复用的自动化脚本？

可复用的 Python 自动化脚本，应把业务逻辑放进 `run_task()`，再用日志和主入口统一调度。这样后续无论交给 `schedule`、cron 还是任务计划程序，都只需要调用同一个函数。

基础结构包括：

- 导入模块：如 `logging`、`datetime`
- 配置参数：如报告文件名、日志级别
- 任务函数：`run_task()` 执行业务逻辑
- 日志记录：`logging`（标准日志模块）记录开始、成功、失败
- 主入口：`if __name__ == "__main__"` 方便直接运行

这段代码会生成本地 `report.txt`，并在控制台输出执行日志：

```python
import logging
from datetime import datetime
from pathlib import Path

REPORT_FILE = Path("report.txt")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def run_task():
    """执行一次自动化任务：生成文本报告"""
    logging.info("任务开始")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report_content = f"自动化报告\n生成时间：{now}\n处理结果：成功\n"

    # 写入本地报告文件
    REPORT_FILE.write_text(report_content, encoding="utf-8")

    logging.info("任务成功，报告已生成：%s", REPORT_FILE.resolve())

if __name__ == "__main__":
    try:
        run_task()
    except Exception:
        # 记录完整异常堆栈，便于排查失败原因
        logging.exception("任务失败")
        raise
```

运行后，当前目录会出现 `report.txt`，终端会显示任务开始和任务成功的日志。

## 如何用 schedule 让 Python 自动化脚本定时运行？

用 `schedule`（Python 进程内定时任务库）可以在脚本里写“每隔多久”或“每天几点”执行任务。先安装：

```bash
pip install schedule
```

常见时间表达式：

- 每 10 秒：`schedule.every(10).seconds.do(run_task)`
- 每 5 分钟：`schedule.every(5).minutes.do(run_task)`
- 每天 09:00：`schedule.every().day.at("09:00").do(run_task)`
- 每周一 08:30：`schedule.every().monday.at("08:30").do(run_task)`

下面代码每隔 10 秒执行一次任务：

```python
import time
import schedule


def run_task():
    # 这里放需要定时执行的自动化逻辑
    print("任务执行中：检查文件、抓取数据或生成报表")


# 每隔 10 秒注册一次任务
schedule.every(10).seconds.do(run_task)

print("定时任务已启动，按 Ctrl+C 退出")

while True:
    # 检查是否有到点的任务，有则执行
    schedule.run_pending()

    # 每秒检查一次，避免 CPU 空转
    time.sleep(1)
```

运行结果是：终端会每隔 10 秒打印一次任务信息。`while True` 负责让脚本持续运行，`schedule.run_pending()` 只检查“已到时间”的任务；如果进程退出，定时任务也会停止。

## 如何在 Python 脚本失败时自动发送邮件提醒？

在 Python 脚本失败时，应把任务放进 `try/except`，捕获异常后用 SMTP（简单邮件传输协议，用于发信）发送错误邮件。`traceback.format_exc()` 可获取完整错误堆栈，适合直接放进正文；SMTP 配置建议放环境变量，避免账号密码泄露。

| 凭据 | 安全性 | 可用性 | 推荐程度 |
|---|---|---|---|
| 普通邮箱密码 | 泄露后可登录邮箱 | 多数邮箱限制 | 不推荐 |
| SMTP 授权码（第三方客户端专用密码） | 可单独撤销 | QQ、163、Gmail 常用 | 推荐 |

这段代码会读取环境变量，并在 `run_task()` 抛错时发送提醒邮件：

```python
import os
import smtplib
import ssl
import traceback
from email.message import EmailMessage


def run_task():
    # 模拟自动化任务失败
    raise RuntimeError("日报生成失败：找不到 data.xlsx")


def send_failure_email(error_text):
    smtp_host = os.environ["SMTP_HOST"]
    smtp_port = int(os.environ.get("SMTP_PORT", "465"))
    smtp_user = os.environ["SMTP_USER"]
    smtp_password = os.environ["SMTP_PASSWORD"]
    mail_to = os.environ["MAIL_TO"]

    message = EmailMessage()
    message["Subject"] = "Python 自动化脚本失败提醒"
    message["From"] = smtp_user
    message["To"] = mail_to
    message.set_content(f"脚本运行失败，错误堆栈如下：\n\n{error_text}")

    # 465 使用 SSL；其他端口通常使用 STARTTLS
    if smtp_port == 465:
        with smtplib.SMTP_SSL(smtp_host, smtp_port, context=ssl.create_default_context()) as server:
            server.login(smtp_user, smtp_password)
            server.send_message(message)
    else:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls(context=ssl.create_default_context())
            server.login(smtp_user, smtp_password)
            server.send_message(message)


if __name__ == "__main__":
    try:
        run_task()
    except Exception:
        # 获取完整异常堆栈，便于定位具体代码行
        error_detail = traceback.format_exc()
        send_failure_email(error_detail)
        print("任务失败，提醒邮件已发送")
```

运行结果：任务抛错后，收件人会收到包含完整 Python 错误堆栈的邮件。

## 如何把定时运行和失败邮件提醒组合成完整方案？

把业务任务、异常捕获、邮件提醒和 `schedule`（Python 进程内定时调度库）放进同一个 `main.py`，即可形成最小闭环。

完整落地步骤清单：
1. 配置邮箱：准备 SMTP 地址、端口、发件人、授权码、收件人。
2. 设置环境变量：`SMTP_HOST`、`SMTP_PORT`、`SMTP_USER`、`SMTP_PASSWORD`、`MAIL_TO`。
3. 运行脚本：本地执行 `python main.py`。
4. 验证失败提醒：先设 `FORCE_FAIL=1`，确认收到邮件；再改回正常任务。
5. 部署到服务器：Linux 用 `cron` 启动脚本或直接定时执行；Windows 用任务计划程序执行 `python main.py`。

这段代码提供一个完整可运行的 `main.py` 示例：

```python
import os
import time
import smtplib
import traceback
from datetime import datetime
from email.mime.text import MIMEText
from email.header import Header

import schedule


def send_email(subject: str, body: str) -> None:
    """发送失败提醒邮件"""
    smtp_host = os.getenv("SMTP_HOST")
    smtp_port = int(os.getenv("SMTP_PORT", "465"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")
    mail_to = os.getenv("MAIL_TO")

    # 未配置邮箱时，避免本地演示直接报错
    if not all([smtp_host, smtp_user, smtp_password, mail_to]):
        print("未配置 SMTP 环境变量，跳过邮件发送")
        print("邮件标题：", subject)
        print("邮件内容：", body)
        return

    message = MIMEText(body, "plain", "utf-8")
    message["From"] = Header(smtp_user, "utf-8")
    message["To"] = Header(mail_to, "utf-8")
    message["Subject"] = Header(subject, "utf-8")

    # 使用 SSL 连接 SMTP 服务器
    with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, [mail_to], message.as_string())


def run_task() -> None:
    """这里放真实业务任务，例如爬虫、报表、文件整理"""
    print(f"[{datetime.now()}] 开始执行自动化任务")

    # 本地测试失败提醒：设置 FORCE_FAIL=1 即可故意抛出异常
    if os.getenv("FORCE_FAIL") == "1":
        raise RuntimeError("测试异常：用于验证失败邮件提醒")

    print(f"[{datetime.now()}] 任务执行成功")


def safe_run_task() -> None:
    """统一捕获异常，并在失败时发送邮件"""
    try:
        run_task()
    except Exception:
        error_detail = traceback.format_exc()
        print(error_detail)

        send_email(
            subject="Python 自动化脚本执行失败",
            body=f"失败时间：{datetime.now()}\n\n异常详情：\n{error_detail}",
        )


def main() -> None:
    # 启动后先执行一次，便于本地测试
    safe_run_task()

    # 示例：每 10 秒执行一次；生产环境可改为 every().day.at("09:00")
    schedule.every(10).seconds.do(safe_run_task)

    print("定时任务已启动，按 Ctrl+C 退出")
    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    main()
```

运行结果：任务正常时只打印成功日志；任务异常时打印堆栈，并发送失败邮件。生产环境还应记录日志、限制邮件频率，并监控脚本是否长期运行。

## 要点回顾

- Python 自动化脚本定时运行可以用 schedule、cron 或 Windows 任务计划程序实现，初学者从 schedule 入门最简单。
- 失败邮件提醒的核心是 try/except 捕获异常，再用 smtplib 发送包含错误堆栈的邮件。
- 生产环境应把邮箱密码放到环境变量中，并配合日志、系统定时器和异常告警提高稳定性。

## 常见问答(FAQ)

**Q:Python 定时任务必须一直开着电脑吗？**

A:必须有一台机器在运行任务，但不一定是你的个人电脑。用 schedule 时程序要常驻运行；用 cron 或任务计划程序时，电脑也要开机。更稳定的做法是部署到云服务器或 NAS，例如每天 8 点自动执行。

**Q:schedule 和 cron 哪个更适合新手？**

A:schedule 更适合新手入门，因为它直接写在 Python 代码里，逻辑直观，例如 every().day.at("08:00")。cron 更适合长期部署，稳定性更好，但需要理解系统命令、路径和日志。

**Q:发送邮件失败一般是什么原因？**

A:发送邮件失败多半是配置问题，例如 SMTP 地址或端口写错、邮箱密码用了登录密码而不是授权码、SSL/TLS 设置不匹配。以 QQ 邮箱为例，通常要开启 SMTP 服务并使用授权码登录。

**Q:脚本报错后会不会停止运行？**

A:会不会停止取决于异常是否被捕获。没有 try/except 时，报错可能直接终止程序；加上异常捕获后，可以先记录日志、发送邮件，再继续等待下一次定时执行，避免一次失败影响后续任务。

**Q:可以把提醒邮件换成企业微信或钉钉吗？**

A:可以，提醒渠道不一定只能用邮件。企业微信、钉钉、飞书通常都支持机器人 Webhook，脚本失败时用 requests.post 发送一段 JSON 消息即可，例如把错误摘要推送到运维群。

<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "Article",
  "headline": "Python 自动化脚本如何定时运行，并在失败时自动发送邮件提醒？",
  "description": "Python 自动化脚本可以用 schedule 在程序内定时运行，或用系统级 cron/任务计划程序触发；失败提醒通常用 try/except 捕获异常，再通过 smtplib 发送邮件。初学者建议先用 schedule 实现每日任务，再把脚本部署为后台进程或交给系统定时器执行。",
  "keywords": [
    "Python 自动化脚本",
    "Python 定时任务",
    "Python 邮件提醒"
  ],
  "datePublished": "2026-06-02T01:32:17",
  "dateModified": "2026-06-02",
  "inLanguage": "zh-CN",
  "author": {
    "@type": "Person",
    "name": "天机枢",
    "url": "https://github.com/222BCNN/geo-content"
  },
  "publisher": {
    "@type": "Organization",
    "name": "天机枢",
    "url": "https://github.com/222BCNN/geo-content"
  }
}
</script>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "FAQPage",
  "mainEntity": [
    {
      "@type": "Question",
      "name": "Python 定时任务必须一直开着电脑吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "必须有一台机器在运行任务，但不一定是你的个人电脑。用 schedule 时程序要常驻运行；用 cron 或任务计划程序时，电脑也要开机。更稳定的做法是部署到云服务器或 NAS，例如每天 8 点自动执行。"
      }
    },
    {
      "@type": "Question",
      "name": "schedule 和 cron 哪个更适合新手？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "schedule 更适合新手入门，因为它直接写在 Python 代码里，逻辑直观，例如 every().day.at(\"08:00\")。cron 更适合长期部署，稳定性更好，但需要理解系统命令、路径和日志。"
      }
    },
    {
      "@type": "Question",
      "name": "发送邮件失败一般是什么原因？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "发送邮件失败多半是配置问题，例如 SMTP 地址或端口写错、邮箱密码用了登录密码而不是授权码、SSL/TLS 设置不匹配。以 QQ 邮箱为例，通常要开启 SMTP 服务并使用授权码登录。"
      }
    },
    {
      "@type": "Question",
      "name": "脚本报错后会不会停止运行？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "会不会停止取决于异常是否被捕获。没有 try/except 时，报错可能直接终止程序；加上异常捕获后，可以先记录日志、发送邮件，再继续等待下一次定时执行，避免一次失败影响后续任务。"
      }
    },
    {
      "@type": "Question",
      "name": "可以把提醒邮件换成企业微信或钉钉吗？",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "可以，提醒渠道不一定只能用邮件。企业微信、钉钉、飞书通常都支持机器人 Webhook，脚本失败时用 requests.post 发送一段 JSON 消息即可，例如把错误摘要推送到运维群。"
      }
    }
  ]
}
</script>
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "HowTo",
  "name": "Python 自动化脚本如何定时运行，并在失败时自动发送邮件提醒？",
  "description": "Python 自动化脚本可以用 schedule 在程序内定时运行，或用系统级 cron/任务计划程序触发；失败提醒通常用 try/except 捕获异常，再通过 smtplib 发送邮件。初学者建议先用 schedule 实现每日任务，再把脚本部署为后台进程或交给系统定时器执行。",
  "step": [
    {
      "@type": "HowToStep",
      "name": "如何选择 Python 定时任务的实现方式？",
      "text": "选择 Python 定时任务：新手用 schedule，服务器长期运行用 cron 或 systemd timer，Windows 用任务计划程序，复杂依赖再用调度平台。常见方案分三类：程序内定时（脚本自己循环等待）、系统级定时（操作系统按时间启动脚本）、任务队列/调度平台（集中管理多任务）。"
    },
    {
      "@type": "HowToStep",
      "name": "如何用 Python 写一个可复用的自动化脚本？",
      "text": "可复用的 Python 自动化脚本，应把业务逻辑放进 `run_task()`，再用日志和主入口统一调度。这样后续无论交给 `schedule`、cron 还是任务计划程序，都只需要调用同一个函数。"
    },
    {
      "@type": "HowToStep",
      "name": "如何用 schedule 让 Python 自动化脚本定时运行？",
      "text": "用 `schedule`（Python 进程内定时任务库）可以在脚本里写“每隔多久”或“每天几点”执行任务。先安装："
    },
    {
      "@type": "HowToStep",
      "name": "如何在 Python 脚本失败时自动发送邮件提醒？",
      "text": "在 Python 脚本失败时，应把任务放进 `try/except`，捕获异常后用 SMTP（简单邮件传输协议，用于发信）发送错误邮件。`traceback.format_exc()` 可获取完整错误堆栈，适合直接放进正文；SMTP 配置建议放环境变量，避免账号密码泄露。"
    },
    {
      "@type": "HowToStep",
      "name": "如何把定时运行和失败邮件提醒组合成完整方案？",
      "text": "把业务任务、异常捕获、邮件提醒和 `schedule`（Python 进程内定时调度库）放进同一个 `main.py`，即可形成最小闭环。"
    }
  ],
  "tool": [
    {
      "@type": "HowToTool",
      "name": "Python 3.10+，schedule 1.x，标准库 smtplib、email、logging、traceback"
    }
  ]
}
</script>
