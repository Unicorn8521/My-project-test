# 屏幕捕捉工具 (Screen Capture Tool)

一个功能强大的屏幕捕捉与操作记录工具，支持快捷键截图、截图编辑、步骤记录和多格式报告生成，帮助你轻松创建操作教程或故障报告。
<img width="752" height="665" alt="image" src="https://github.com/user-attachments/assets/0a42071d-b210-4224-9626-a48b67529893" />


## 功能特点

- **便捷截图**：通过自定义快捷键快速捕捉屏幕内容
- **截图编辑**：支持截图裁剪、自由标注和添加文字说明
- **步骤记录**：自动记录操作步骤和时间线，形成完整操作流程
- **多格式报告**：支持生成 Word (.docx)、PDF 和 Markdown 格式的操作报告
- **历史管理**：保存历史操作记录，支持重命名、删除和重新编辑
- **中文支持**：全面支持中文显示和输入，解决各类文档中的中文乱码问题

## 安装说明

### 前提条件

- Python 3.7 或更高版本

### 安装依赖

```bash
pip install pillow python-docx reportlab keyboard tkinter
```

对于 Linux 系统，还需要安装额外依赖以支持截图功能：

```bash
sudo apt install python3-xlib
```

## 使用方法

1. 运行程序：

   ```bash
   python screen_capture_tool.py
   ```

   或者直接运行screen_capture_tool.exe文件

2. 开始捕捉：

   - 点击 "开始捕捉" 按钮
   - 输入操作名称和描述
   - 使用默认快捷键 `Ctrl+Alt+O` 截取屏幕（可在程序中修改默认快捷键）

3. 编辑截图：

   - 截图后可进行裁剪
   - 支持自由绘制标注和添加文字说明
   - 为每步操作添加描述信息
  <img width="1916" height="1077" alt="image" src="https://github.com/user-attachments/assets/c9cca8cb-ec30-4cf1-9341-02f1c48315e3" />

  <img width="973" height="771" alt="image" src="https://github.com/user-attachments/assets/e4f6e53d-ea1c-4a3a-9593-bcc7bbab13de" />


4. 生成报告：

   - 完成所有操作后，点击 "停止捕捉"
   - 在编辑界面中完善步骤描述
   - 选择所需格式（Word/PDF/Markdown）导出报告
  <img width="440" height="352" alt="image" src="https://github.com/user-attachments/assets/f70c4491-55a1-41d9-bfdc-c4798c54dfb5" />


## 界面说明

- **主界面**：显示历史记录和控制按钮
- **截图预览**：全屏预览截图，支持区域选择和标注
- **编辑窗口**：修改步骤描述，调整步骤顺序
- **报告设置**：选择导出格式和保存路径

## 报告示例

生成的报告包含以下内容：

- 操作基本信息（名称、描述、时间、时长）
- 详细步骤列表，包含每个步骤的描述、时间和截图
- 自动编号的步骤和图片引用

## 自定义配置

可在代码中的 `Config` 类修改以下参数：

- 默认窗口大小
- 快捷键设置
- 字体配置
- 报告格式设置

## 致谢

本工具使用了以下优秀的开源库：

- [Pillow](https://python-pillow.org/) - 图像处理
- [python-docx](https://python-docx.readthedocs.io/) - Word 文档生成
- [reportlab](https://www.reportlab.com/) - PDF 文档生成
- [keyboard](https://github.com/boppreh/keyboard) - 键盘事件处理
- [Tkinter](https://docs.python.org/3/library/tkinter.html) - GUI 界面

## 问题反馈

如有任何问题或建议，请在 GitHub 仓库提交 Issue。
作者邮箱：2742234948@qq.com
欢迎大家提改进意见，有好的想法可以和作者多多交流哦！
