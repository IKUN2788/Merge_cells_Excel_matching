<img width="984" height="841" alt="PixPin_2026-02-10_16-31-23" src="https://github.com/user-attachments/assets/f7bd02ac-ace1-40eb-ad95-2239ee555847" />

# 合并单元格匹配工具（PyQt5）

## 功能
- 从“数据源Excel”读取含合并单元格的表格，自动展开合并单元格
- 选择 Key 列（可多选，用 “_” 拼接）与 Value 列，生成 key→value 映射
- 对“匹配源Excel”按 Key 进行匹配，将结果写入指定列（默认新增到最后一列）

## 运行
1. 安装依赖：

```bash
pip install -r requirements.txt
```

2. 启动：

```bash
python 合并单元格匹配工具.py
```

## 说明
- 默认将第一行作为表头（列名来自第一行）
- “累加”勾选时：
  - 纯数值：按 key 求和
  - 非数值：使用 “;” 拼接并去重
- 当“匹配写入列”选择“新增到最后一列”时，可自定义新增列表头；留空则自动生成“匹配_{Value列名}”

## 最小化打包（Windows）
1. 执行一键构建脚本：

```powershell
.\build_exe.ps1
```

2. 生成产物：
- dist\merge_cell_match_tool.exe

说明：
- exe 图标使用项目目录下的 app.ico（同时也会作为运行时资源打包）
- 若要进一步减小体积，可自行安装 UPX 并在 PyInstaller 中配置（此项目未强制依赖 UPX）
