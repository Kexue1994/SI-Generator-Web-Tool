#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2025/11/5 20:54
# @Author  : Healer
# @File    : scf_app.py
# @Software: PyCharm


# -*- coding: utf-8 -*-
import json
import base64
import os
import sys
import tempfile
from io import BytesIO

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(__file__))

try:
    # 尝试导入您的主应用
    from app import SIGeneratorWeb

    HAS_MAIN_APP = True
except ImportError as e:
    print(f"导入主应用失败: {e}")
    HAS_MAIN_APP = False


def main_handler(event, context):
    """
    云函数主处理器
    """
    print("收到请求:", event.get('httpMethod', 'GET'))

    try:
        # 处理预检请求（CORS）
        if event.get('httpMethod') == 'OPTIONS':
            return {
                "isBase64Encoded": False,
                "statusCode": 200,
                "headers": {
                    'Access-Control-Allow-Origin': '*',
                    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
                    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
                    'Access-Control-Max-Age': '86400'
                },
                "body": ""
            }

        # 获取请求方法和路径
        http_method = event.get('httpMethod', 'GET')
        path = event.get('path', '/')

        # 路由处理
        if path == '/' or path == '/index.html':
            return serve_static_page()
        elif path == '/upload' and http_method == 'POST':
            return handle_file_upload(event)
        else:
            return serve_static_page()

    except Exception as e:
        print(f"处理请求时出错: {e}")
        return error_response(f"服务器错误: {str(e)}")


def serve_static_page():
    """返回静态HTML页面"""
    html_content = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SI Generator Tool - 腾讯云函数版</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Microsoft YaHei', Arial, sans-serif; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .container {
            max-width: 1000px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #2c3e50, #3498db);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
        }
        .header p {
            opacity: 0.9;
            font-size: 1.1em;
        }
        .content {
            padding: 40px;
        }
        .upload-area {
            border: 3px dashed #3498db;
            border-radius: 10px;
            padding: 40px;
            text-align: center;
            margin: 20px 0;
            background: #f8f9fa;
            transition: all 0.3s ease;
        }
        .upload-area:hover {
            border-color: #2980b9;
            background: #e8f4fc;
        }
        .upload-area h3 {
            color: #2c3e50;
            margin-bottom: 15px;
            font-size: 1.4em;
        }
        .btn {
            background: linear-gradient(135deg, #3498db, #2980b9);
            color: white;
            border: none;
            padding: 12px 30px;
            border-radius: 25px;
            font-size: 1.1em;
            cursor: pointer;
            transition: all 0.3s ease;
            margin: 10px;
        }
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(52, 152, 219, 0.4);
        }
        .btn:disabled {
            background: #bdc3c7;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        .instructions {
            background: #e8f4fc;
            padding: 25px;
            border-radius: 10px;
            margin: 25px 0;
            border-left: 5px solid #3498db;
        }
        .instructions h3 {
            color: #2c3e50;
            margin-bottom: 15px;
        }
        .instructions ul {
            list-style: none;
            padding-left: 20px;
        }
        .instructions li {
            margin: 10px 0;
            padding-left: 25px;
            position: relative;
        }
        .instructions li:before {
            content: "✓";
            color: #27ae60;
            font-weight: bold;
            position: absolute;
            left: 0;
        }
        .status {
            padding: 15px;
            border-radius: 8px;
            margin: 15px 0;
            text-align: center;
            font-weight: bold;
        }
        .status.success { background: #d4edda; color: #155724; }
        .status.error { background: #f8d7da; color: #721c24; }
        .status.info { background: #d1ecf1; color: #0c5460; }
        .feature-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }
        .feature-card {
            background: white;
            padding: 25px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            border: 2px solid transparent;
            transition: all 0.3s ease;
        }
        .feature-card:hover {
            border-color: #3498db;
            transform: translateY(-5px);
        }
        .feature-icon {
            font-size: 2.5em;
            margin-bottom: 15px;
        }
        .footer {
            text-align: center;
            padding: 20px;
            background: #f8f9fa;
            color: #6c757d;
            border-top: 1px solid #dee2e6;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 SI Generator Tool</h1>
            <p>基于腾讯云函数的SI文件生成工具</p>
        </div>

        <div class="content">
            <div class="feature-grid">
                <div class="feature-card">
                    <div class="feature-icon">🚀</div>
                    <h3>快速部署</h3>
                    <p>基于腾讯云函数，无需服务器管理</p>
                </div>
                <div class="feature-card">
                    <div class="feature-icon">📁</div>
                    <h3>批量处理</h3>
                    <p>支持多个Excel文件同时处理</p>
                </div>
                <div class="feature-card">
                    <div class="feature-icon">🔒</div>
                    <h3>数据安全</h3>
                    <p>文件在处理后自动清理，保障数据安全</p>
                </div>
            </div>

            <div class="instructions">
                <h3>使用说明</h3>
                <ul>
                    <li>准备包含'No SI Order'和'SI Template'工作表的Excel文件</li>
                    <li>系统将按指定列（默认O列）自动分组数据</li>
                    <li>为每个分组生成独立的SI文件</li>
                    <li>支持批量下载生成的文件</li>
                    <li>完全基于浏览器操作，无需安装任何软件</li>
                </ul>
            </div>

            <div class="upload-area">
                <h3>文件上传区域</h3>
                <p>当前版本运行在腾讯云函数环境中</p>
                <p>文件处理功能需要完整服务器部署</p>
                <p style="margin: 20px 0; color: #e74c3c; font-weight: bold;">
                    ⚠️ 云函数环境限制：文件处理功能需要完整服务器部署
                </p>
                <button class="btn" onclick="showMessage()">上传Excel文件</button>
                <button class="btn" onclick="showMessage()">生成SI文件</button>
            </div>

            <div id="statusMessage" class="status info" style="display: none;">
                提示信息将在这里显示
            </div>
        </div>

        <div class="footer">
            <p>Powered by 腾讯云函数 | 建议使用Chrome浏览器访问</p>
            <p>技术支持：请联系系统管理员获取完整部署版本</p>
        </div>
    </div>

    <script>
        function showMessage() {
            const statusDiv = document.getElementById('statusMessage');
            statusDiv.innerHTML = '💡 完整文件处理功能需要服务器部署。请联系管理员部署完整版本。';
            statusDiv.className = 'status info';
            statusDiv.style.display = 'block';
        }

        // 页面加载完成后的初始化
        document.addEventListener('DOMContentLoaded', function() {
            console.log('SI Generator Tool 已加载');
        });
    </script>
</body>
</html>
    """

    return {
        "isBase64Encoded": False,
        "statusCode": 200,
        "headers": {
            'Content-Type': 'text/html; charset=utf-8',
            'Access-Control-Allow-Origin': '*'
        },
        "body": html_content
    }


def handle_file_upload(event):
    """处理文件上传请求"""
    return {
        "isBase64Encoded": False,
        "statusCode": 200,
        "headers": {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
        },
        "body": json.dumps({
            "status": "success",
            "message": "文件上传功能需要在完整服务器环境中运行",
            "note": "云函数环境适合展示页面，文件处理建议使用服务器部署"
        })
    }


def error_response(message):
    """返回错误响应"""
    return {
        "isBase64Encoded": False,
        "statusCode": 500,
        "headers": {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*'
        },
        "body": json.dumps({
            "status": "error",
            "message": message
        })
    }


# 测试函数
def test_local():
    """本地测试函数"""
    test_event = {
        "httpMethod": "GET",
        "path": "/"
    }
    result = main_handler(test_event, None)
    print("测试结果:", result)


if __name__ == "__main__":
    test_local()