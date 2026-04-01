"""
测试用例智能体的工具模块

本模块负责初始化和管理测试用例生成智能体所需的各类工具，
主要通过 MCP (Model Context Protocol) 协议连接外部服务，
为智能体提供文档解析等扩展能力。

MCP 协议说明：
    MCP (Model Context Protocol) 是一种开放协议，用于标准化 AI 模型与
    外部工具/服务之间的通信。它允许 AI 助手通过统一的接口访问各种能力，
    如文件系统、数据库、API 等。

依赖服务：
    - docling-mcp: 文档解析服务，支持 PDF、Word、Markdown 等格式的文档解析
      启动命令: uvx --from docling-mcp docling-mcp-server --transport sse --port 8003
"""

import asyncio

from langchain_mcp_adapters.client import MultiServerMCPClient
from .excel_tool import save_testcases_to_excel


def create_mcp_client() -> MultiServerMCPClient:
    """
    创建并配置 MCP 客户端
    
    创建一个 MultiServerMCPClient 实例，配置连接到多个 MCP 服务器。
    当前配置连接 docling-server 用于文档解析。
    
    Returns:
        MultiServerMCPClient: 配置好的 MCP 客户端实例
        
    Note:
        使用 SSE (Server-Sent Events) 传输协议与 MCP 服务器通信，
        确保 docling-mcp 服务器已在指定端口启动。
    """
    # 配置 MCP 服务器连接信息
    # 键名为服务器标识，在代码中可通过该名称引用特定服务器
    server_configs = {
        # Docling 文档解析服务配置
        # 该服务提供文档解析能力，可将 PDF、Word 等文档转换为结构化文本
        "docling-server": {
            # SSE 服务端点 URL
            "url": "http://127.0.0.1:8000/sse",
            # 传输协议：sse (Server-Sent Events)
            "transport": "sse",
        }
    }
    
    return MultiServerMCPClient(server_configs)


# ============================================================================
# 全局工具初始化
# ============================================================================

# 尝试加载 MCP 工具
mcp_tools = []
try:
    client = create_mcp_client()
    mcp_tools = asyncio.run(client.get_tools())
    print(f"✅ 已加载 {len(mcp_tools)} 个 MCP 工具")
except Exception as e:
    print(f"⚠️  MCP 工具加载失败: {e}")
    print("继续使用本地工具...")

# 添加本地工具
tools = mcp_tools + [save_testcases_to_excel]

print(f"📦 总共加载 {len(tools)} 个工具: {[t.name for t in tools]}")
