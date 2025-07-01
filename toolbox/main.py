# -*- coding: utf-8 -*-
from mcp.server.fastmcp import FastMCP

mcp = FastMCP(name="toolbox")


@mcp.tool(
    name="retrieve_current_datetime",
    description="获取当前日期和时间",
    annotations={"title": "获取当前日期和时间"},
)
async def retrieve_current_datetime():
    """
    返回当前日期和时间（ISO 8601 格式字符串）。
    """
    from datetime import datetime

    now = datetime.now().isoformat()
    return {"datetime": now}


if __name__ == "__main__":
    mcp.run(transport="stdio")
