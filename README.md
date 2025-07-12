# E550电子秤重量显示与上传系统

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

## 项目简介

E550电子秤重量显示与上传系统是一个连接CH340电子秤的桌面应用程序，可实时显示重量数据并上传至远程服务器。该应用程序支持自动检测和连接电子秤设备，提供稳定性指示、数据上传和历史记录查看等功能。

## 主要功能

- 自动连接并识别CH340串口电子秤
- 实时显示重量数据
- 重量稳定性指示
- 记录包裹尺寸信息（长、宽、高）
- 将重量和尺寸数据上传到指定服务器
- 显示上传历史记录
- 支持设置管理和参数配置

## 安装要求

- Windows操作系统
- Python 3.6+
- 以下Python库:
  - pyserial==3.5
  - requests==2.31.0
  - ttkbootstrap==1.10.1

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用说明

1. 连接电子秤到计算机USB端口
2. 运行程序
3. 程序将自动尝试连接电子秤
4. 将包裹放置在电子秤上，等待重量稳定（指示灯变为绿色）
5. 输入扫描单号和可选的包裹尺寸
6. 点击"上传重量"或按回车键上传数据

## 配置说明

### 初始配置

首次使用时，需要在设置中配置以下参数：

1. 设备编号：电子秤的唯一标识，例如"SCALE_001"
2. API域名和端口：远程服务器地址和端口，例如"api.example.com"和"80"
3. 用户ID和安全密钥：API访问凭证，从服务提供商获取

> **注意**：程序初始包含的配置信息仅为示例，需要用户自行替换为有效的配置。

### 详细设置步骤

1. 运行程序后，点击菜单中的"设置"选项
2. 输入默认密码"password"（建议首次登录后修改）
3. 在设置界面中填写以下信息：
   - **设备编号**：您的电子秤设备编号
   - **API域名**：数据上传服务器的域名（不含http://）
   - **API端口**：服务器端口，通常为80或443
   - **用户ID**：您的API账户用户名
   - **安全密钥**：您的API账户密钥
   - **串口设置**：一般会自动检测，如需手动设置可在此处选择

4. 点击"保存"按钮完成配置

### 上传数据格式

当您点击"上传重量"按钮时，系统将以以下格式发送数据到服务器：

```json
{
  "deviceNo": "您的设备编号",
  "scanNo": "扫描的包裹单号",
  "weight": 12.345,     // 重量（千克）
  "length": 30.0,       // 长度（厘米，可选）
  "width": 20.0,        // 宽度（厘米，可选）
  "height": 15.0,       // 高度（厘米，可选）
  "pictureBase64": ""   // 图片数据（未使用）
}
```

### API响应格式

服务器预期返回的JSON格式如下：

```json
{
  "code": 0,            // 0表示成功
  "msg": "成功",        // 状态消息
  "data": {
    "scanNo": "包裹单号"
  }
}
```

> **安全提示**：请妥善保管您的API凭证，定期更改设置密码，避免在不安全网络环境中使用本软件。

## 使用PyInstaller打包

要将程序打包成单个可执行文件，可以使用PyInstaller。以下是打包步骤：

1. 安装PyInstaller
```bash
pip install pyinstaller
```

2. 进入项目目录
```bash
cd 程序路径\E550_Electronic_scales
```

3. 使用PyInstaller创建单文件
```bash
pyinstaller --onefile --windowed --icon=weight-scale.ico --add-data "weight-scale.ico;." "E550串口测试V63.py"
```

4. 打包完成后，可执行文件将位于`dist`文件夹中

## 故障排除

- **无法连接电子秤**：检查USB连接，确保CH340驱动已正确安装
- **重量显示不稳定**：确保电子秤放置在平稳表面，避免振动
- **上传数据失败**：检查网络连接，确认API配置正确
- **串口无法识别**：尝试在设置中手动选择正确的COM端口

## 隐私与安全

- 本应用程序不会收集或存储用户个人数据
- 敏感信息（如API密钥）存储在本地配置文件中
- 建议定期更改设置密码以增强安全性

## 许可

本项目采用 GNU General Public License v3.0 (GPL-3.0) 许可证

Copyright (C) 2023-2025

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
