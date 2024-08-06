#!/bin/bash

# 检测 node 是否已安装
if ! command -v node &> /dev/null
then
    echo "Node.js 未安装. 正在安装 Node.js..."
    
    # 检查是否有权限安装 Node.js
    if [ "$EUID" -ne 0 ]
    then
        echo "请使用 root 用户或具有 sudo 权限的用户运行此脚本."
        exit 1
    fi

    # 使用 n 安装 Node.js
    if ! command -v n &> /dev/null
    then
        echo "安装 n (Node.js 版本管理工具)..."
        curl -L https://raw.githubusercontent.com/tj/n/master/bin/n -o n
        bash n lts
        rm n
    else
        n lts
    fi
    
    # 确保 node 命令可用
    export PATH="$PATH:/usr/local/bin"
fi

# 再次检测 node 是否已安装成功
if ! command -v node &> /dev/null
then
    echo "Node.js 安装失败."
    exit 1
fi

echo "Node.js 已安装. 正在启动 server.js..."
node server.js
