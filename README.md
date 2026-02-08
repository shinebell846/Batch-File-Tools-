# Batch-File-Tools-
一个使用Python Tkinter编写的桌面应用程序，集成了多种文件处理功能，包括批量重命名、图片格式转换和Excel超链接转换
批量文件重命名（支持修改后缀）

图片格式批量转换
<img width="995" height="826" alt="屏幕截图 2026-02-08 120540" src="https://github.com/user-attachments/assets/12a5433c-303c-42ec-beff-390a27b2c6f7" />

Excel超链接转换（支持正则匹配）

🏗️ 架构设计
分层架构
text
MainApplication (主窗口)
├── BaseModule (基类)
├── RenameModule (重命名模块)
├── ConvertModule (图片转换模块)
└── HyperlinkModule (超链接模块)
核心特点
模块化设计：每个功能独立封装，便于维护和扩展

线程安全：后台任务使用线程处理，保持UI响应

队列日志：线程安全的日志系统，实时更新UI

配置持久化：超链接模块支持自定义样式保存

🔧 各模块功能详解
1. 批量重命名模块 (RenameModule)
功能：批量重命名文件夹中的文件

特性：

支持自定义前缀、排序方式、序号位数

新增后缀修改功能（关键修复）

自动检测文件名冲突

倒序处理避免覆盖问题

2. 图片格式转换模块 (ConvertModule)
功能：批量转换图片格式

支持格式：PNG, JPEG, BMP, WEBP, ICO

特性：

支持透明通道处理（RGBA转RGB）

ICO格式支持自定义尺寸

可单独选择文件或整个文件夹

JPEG/WEBP可调质量参数

3. Excel超链接转换模块 (HyperlinkModule)
功能：Excel中超链接与文本的相互转换

特性：

新增样式管理界面（Treeview展示）

预置常见网盘URL模式

支持自定义正则表达式

区分预置样式和自定义样式

配置持久化到JSON文件

💡 技术亮点
1. UI/UX设计
现代化配色方案（蓝色主题）

响应式布局

清晰的视觉层次

操作状态反馈（按钮状态变化）

2. 代码质量
遵守PEP8规范

完善的错误处理

清晰的注释

函数职责单一

3. 线程处理
python
# 后台任务处理模式
Thread(target=self.batch_rename, kwargs=params, daemon=True).start()
4. 日志系统
使用队列(Queue)线程安全

彩色日志标签（成功/警告/错误）

自动滚动到最新

🚀 GitHub上传建议
1. 仓库结构建议
text
toolbox-v4.0/
├── src/
│   ├── main.py              # 主程序入口
│   ├── modules/
│   │   ├── __init__.py
│   │   ├── base.py          # BaseModule
│   │   ├── rename.py        # RenameModule
│   │   ├── convert.py       # ConvertModule
│   │   └── hyperlink.py     # HyperlinkModule
│   └── config/
│       └── default_styles.py # 样式配置
├── requirements.txt
├── README.md
└── LICENSE
2. 依赖文件 (requirements.txt)
txt
Pillow>=9.0.0
openpyxl>=3.0.0
3. README.md建议内容
项目标题：多功能文件处理工具箱 v4.0

功能演示：GIF/截图展示三个功能

安装说明：pip install -r requirements.txt

使用说明：每个模块的详细步骤

特性列表：

🎨 现代化GUI界面

⚡ 多线程后台处理

📁 批量文件操作

🔧 可扩展模块设计

💾 配置持久化

4. 标签建议
text
python, tkinter, file-management, batch-processing, gui-application, 
excel-automation, image-processing, file-renamer
5. LICENSE选择
推荐：MIT License（适合开源工具）

或 Apache 2.0
