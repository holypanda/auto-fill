# Auto-Fill 文档自动填写系统

批量填充金融/法律模板（.docx 和 .pdf），一次录入数据，自动生成多份文档。

## 功能概述

- **DOCX 模板填充** — 自动识别 `{占位符}`，替换后保留原始格式
- **PDF 表单填充** — 通过 JSON 映射配置，填充 PDF 表单字段
- **批量生成** — 勾选多个模板，一键生成 ZIP 打包下载
- **模板管理** — 上传/删除模板，在线预览 PDF 字段并配置映射
- **数据持久化** — 表单数据可保存为 JSON 文件，支持导入/导出及 localStorage 自动暂存

## 项目结构

```
auto-fill/
├── backend/
│   ├── main.py          # FastAPI 服务，提供 REST API
│   ├── engine.py        # 模板填充引擎（DOCX/PDF）
│   └── __init__.py
├── frontend/
│   └── index.html       # 单页前端应用（内联 CSS/JS）
├── templates/           # 模板文件目录（.docx / .pdf + .json）
├── output/              # 生成的文档输出目录
├── requirements.txt
└── README.md
```

## 快速部署

### 环境要求

- Python 3.10+

### 安装与启动

```bash
# 安装依赖
pip install -r requirements.txt

# 启动服务
uvicorn backend.main:app --host 0.0.0.0 --port 8000

# 访问
# http://localhost:8000
```

## 使用说明

### 1. 模板管理

切换到「模板管理」标签页：

- **上传模板** — 拖拽或点击上传 `.docx` / `.pdf` 文件
- **删除模板** — 点击模板卡片上的删除按钮
- **PDF 字段映射** — 点击 PDF 模板的「Field Mapping」按钮，可视化预览字段位置并配置占位符与 PDF 表单字段的映射关系

### 2. 填写表单

切换到「填写信息」标签页：

- 表单按分组展示：实体信息、地址、税务、董事、控权人等
- 根据勾选的模板动态显示相关字段
- 支持「全选 / 清除」快速操作模板选择

### 3. 生成文档

- **批量生成** — 勾选多个模板后点击底部「Fill N Templates」按钮，下载 ZIP 包
- **单个生成** — 在模板卡片上点击「Fill Single」，直接下载单个文件

### 4. 数据保存与加载

- **保存到文件** — 点击顶部「Save」按钮，导出当前表单数据为 JSON 文件
- **从文件加载** — 点击顶部「Load」按钮，导入之前保存的 JSON 数据
- **自动暂存** — 表单数据自动保存到浏览器 localStorage，关闭页面后可恢复

## 模板占位符规范

DOCX 模板中使用 `{FieldName}` 格式标记占位符，例如：

```
公司名称：{Entity Name}
注册地点：{Entity Location}
董事姓名：{Director Name}
```

系统会自动扫描模板中所有 `{...}` 占位符并生成对应的表单输入项。

## PDF 模板配置

PDF 模板需要两个配套 JSON 文件（放在 `templates/` 目录下）：

### 字段映射文件（`<template>.json`）

将占位符映射到 PDF 内部表单字段名：

```json
{
  "{Entity Name}": "topmostSubform[0].Page1[0].f1_1[0]",
  "{TIN}": "topmostSubform[0].Page2[0].f2_3[0]"
}
```

### 字段描述文件（`<template>_desc.json`，可选）

为 PDF 表单字段提供人类可读的描述：

```json
{
  "topmostSubform[0].Page1[0].f1_1[0]": "Line 1: 实体名称 Name of organization"
}
```

可通过前端「Field Mapping」功能可视化配置，无需手动编辑 JSON。

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/` | 前端页面 |
| GET | `/api/templates` | 获取所有模板及其占位符 |
| GET | `/api/fields` | 获取所有去重后的占位符字段 |
| POST | `/api/fill` | 批量填充模板，返回 ZIP |
| POST | `/api/fill-single` | 填充单个模板，返回文件 |
| GET | `/api/download-template/{name}` | 下载原始模板 |
| POST | `/api/upload-template` | 上传新模板 |
| DELETE | `/api/templates/{name}` | 删除模板及关联配置 |
| GET | `/api/pdf-fields/{name}` | 获取 PDF 表单字段列表 |
| GET | `/api/pdf-preview-labeled/{name}` | 获取 PDF 字段标注预览图 |
| GET | `/api/pdf-field-map/{name}` | 获取 PDF 字段映射 |
| POST | `/api/pdf-field-map/{name}` | 保存 PDF 字段映射 |
