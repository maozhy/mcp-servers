from mcp.server.fastmcp import FastMCP
import smtplib
from email.message import EmailMessage
import os

# Initialize FastMCP server
mcp = FastMCP("email_server")


@mcp.tool()
async def send_email(to: list, sub: str, message: str, is_ok: bool = False) -> str:
    if is_ok == False:
        return "待用户二次确认"
    SMTP_SERVER = "smtp.qq.com"
    SMTP_PORT = 465
    SMTP_USER = "2383101175@qq.com"
    SMTP_PASS = "oqypfketbsaqecbi"
    SMTP_SSL = True

    msg = EmailMessage()
    msg["Subject"] = sub
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join(to)
    msg.set_content(message)

    try:
        if SMTP_SSL:
            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
                server.login(SMTP_USER, SMTP_PASS)
                server.send_message(msg)
                server.close()
        else:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASS)
                server.send_message(msg)
                server.close()
        return "邮件发送成功"
    except Exception as e:
        res = f"邮件发送失败: {type(e).__name__}: {e}"
        return res


if __name__ == "__main__":
    # Initialize and run the server
    mcp.run(transport="stdio")
