# Time Recorder

![image](https://github.com/user-attachments/assets/b0280fa1-f8e3-45c4-87e8-ecd0fa84cd5a)
![image](https://github.com/user-attachments/assets/82c4467b-0ba9-4b70-99aa-e0907f2c2fd8)

![image](https://github.com/user-attachments/assets/eb3ab2b5-3664-4558-b7b1-da8ee9845e28)
![image](https://github.com/user-attachments/assets/ef787152-e106-4799-a254-2a0faa58e8cd)
![image](https://github.com/user-attachments/assets/bb05b06a-c73f-4266-987a-c2bbe3c781d2)


### 项目名称
多功能计时与任务管理软件

### 应用场景
你是否曾因忙碌了一天，在结束后感觉到“今天咋这么忙？怎么好像什么都没干完就过了一天了？”
利用该软件，你可轻松记录每一天时间都画到哪里了，实现自我监督，并具备待办列表提醒未做事项。

### 项目描述
这是一个基于 Python `tkinter` 库开发的多功能桌面应用程序，主要用于计时、待办事项管理、数据可视化和便签记录。用户可以通过该软件记录工作或学习时间，管理待办事项，将数据导出为 Excel 文件，并对计时数据进行可视化展示。

### 功能特性
1. **计时功能**：支持开始、暂停、恢复和结束计时，自动计算暂停时长和用工时长，并将记录显示在表格中。
2. **待办事项管理**：提供一个待办列表，用户可以输入新增或从本地导入之前的待办事项，待办事项可作为计时备注选项。
3. **数据导出**：可以将计时记录和待办事项导出到 Excel 文件，同时支持导出计时数据的饼图和柱状图。
4. **数据可视化**：在可视化页签中，用户可以查看计时数据的饼图和柱状图，通过鼠标左键点击切换图表类型。
5. **便签功能**：提供一个便签页签，用户可以在其中记录文本信息。
6. **使用说明**：在使用说明页签中，提供了软件的详细使用说明和注意事项。
7. **窗口吸附效果**：窗口拖动到屏幕边缘时会自动隐藏，鼠标悬停在边缘时窗口会弹出。

### 使用的编程语言和框架
- **编程语言**：Python
- **主要框架**：`tkinter`（Python 内置的 GUI 库）、`openpyxl`（用于处理 Excel 文件）、`pandas`（用于数据处理）、`matplotlib`（用于数据可视化）

### 安装和运行步骤
#### 安装依赖
确保你已经安装了 Python 环境，然后使用以下命令安装所需的库：
```bash
pip install openpyxl pandas matplotlib
```

#### 运行代码
将代码保存为一个 Python 文件（例如 `timer_app.py`），然后在终端中运行以下命令：
```bash
python timer_app.py
```

### 代码示例
以下是一个简单的使用示例，展示如何启动应用程序：
```python
if __name__ == "__main__":
    import tkinter as tk
    root = tk.Tk()
    root.geometry('800x200+600+300')
    from your_module import TimerApp  # 替换为实际的模块名
    app = TimerApp(root)
    root.mainloop()
```

### 贡献指南
如果你想为这个项目做出贡献，可以按照以下步骤进行：
1. **Fork 项目**：在 GitHub 上 Fork 这个项目到你的仓库。
2. **创建分支**：在你的仓库中创建一个新的分支，用于开发新功能或修复问题。
3. **提交代码**：在分支上进行开发，并提交代码到你的仓库。
4. **发起 Pull Request**：在 GitHub 上发起一个 Pull Request，将你的分支合并到主项目中。

### 许可证
本项目采用 Apache-2.0 许可证

### 注意事项
- 导出 Excel 文件时，确保代码和 Excel 文件在同一个文件夹内。
- 当点击“结束计时”时，备注为必填项，否则无法结束计时。
- 软件关闭时会自动导出 Excel 文件，以防止数据丢失。



