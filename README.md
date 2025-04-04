# PDF提取工具 - 使用说明书

**版本：1.0版 (2025年3月)**

## 目录

1. [软件介绍](#1-软件介绍)
2. [快速上手](#2-快速上手)
3. [功能详解](#3-功能详解)
4. [使用场景](#4-使用场景)
5. [常见问题及解决](#5-常见问题及解决)
6. [注意事项](#6-注意事项)

## 1. 软件介绍

PDF提取工具是一款专为采购信息管理设计的软件。它能帮您从大量PDF文件中快速提取关键信息，如项目名称、采购预算、中标单位等，并整理到Excel表格中，大大提高工作效率。

特点：
- 批量处理多个PDF文件
- 自动识别采购信息
- 汇总整理到Excel表格
- 友好的图形界面操作
- 可追加新信息到已有表格

## 2. 快速上手

### 第一步：选择要处理的PDF文件
1. 点击【选择PDF文件】按钮选择一个或多个PDF文件
2. 或点击【选择文件夹】按钮选择一个包含PDF文件的文件夹
3. 如果需要处理文件夹下的所有子文件夹，请勾选【包含子文件夹】
4. 可以在【文件名过滤】框中输入关键词，只处理文件名中包含特定关键词的PDF

### 第二步：选择键名文件
1. 点击【选择键名文件】按钮选择一个包含要提取信息的键名列表的文本文件
   - 键名文件是一个简单的文本文件(.txt)，每行一个键名，例如"采购项目名称"、"中标金额"等
   - 软件会根据键名文件中列出的项目在PDF中寻找匹配信息
   - 软件附带了一个示例键名文件(key_names_example.txt)，可以根据需要修改

### 第三步：设置提取选项
1. 阅读顺序：根据PDF表格结构选择【从上到下】或【从左到右】
2. 如果希望保留未找到值的键名，勾选【允许值为空】
3. 项目信息处理：
   - 【每个文件夹作为不同项目】：每个文件夹生成一行Excel记录
   - 【所有文件夹作为同一项目】：所有文件夹合并为一行Excel记录

### 第四步：处理和导出
1. 直接导出到新Excel：
   - 点击【处理并导出】按钮
   - 选择保存位置和文件名
   - 等待处理完成

2. 追加到现有Excel：
   - 点击【选择现有Excel】按钮
   - 输入包含标题的行号（通常为1）
   - 点击【新增表格信息】按钮
   - 等待处理完成

## 3. 功能详解

### 文件选择
- 可选择单个或多个PDF文件
- 可选择整个文件夹(包含子文件夹选项)
- 可通过关键词过滤文件名

### 键名配置
- 通过文本文件管理要提取的信息项
- 无需编程知识，直接编辑文本文件添加或删除需要提取的项目

### 阅读顺序
- 从上到下：适合信息垂直排列的表格
- 从左到右：适合信息水平排列的表格

### 项目信息处理模式
- 每个文件夹作为不同项目模式：
  > 例如：文件夹A和文件夹B各包含若干PDF文件，选择此模式会在Excel中生成两行记录，分别对应A和B的信息。

- 所有文件夹作为同一项目模式：
  > 例如：一个项目的资料分散在多个文件夹中，选择此模式会将所有文件夹的信息合并，在Excel中只生成一行记录。

### Excel操作
- 处理并导出：创建新的Excel文件
- 新增表格信息：将提取的信息追加到现有Excel文件中
  > 注意：追加模式会自动添加"追加时间"列，便于追踪数据添加时间

## 4. 使用场景

### 场景一：初次整理信息
1. 将所有PDF文件按项目组织到不同文件夹
2. 选择【每个文件夹作为不同项目】模式
3. 使用【处理并导出】功能创建新Excel

### 场景二：更新现有项目库
1. 将新增PDF文件按项目组织到不同文件夹
2. 选择【每个文件夹作为不同项目】模式
3. 使用【选择现有Excel】和【新增表格信息】功能追加到现有Excel

### 场景三：补充项目信息
1. 将补充资料放入与原始项目相同的文件夹
2. 选择【所有文件夹作为同一项目】模式
3. 使用【新增表格信息】功能更新Excel

## 5. 常见问题及解决

### 问题：某些信息未被正确提取
- **解决方法**：
  - 检查键名文件中是否包含对应的键名
  - 确认PDF是文本格式而非扫描图片
  - 尝试切换阅读顺序(从上到下/从左到右)

### 问题：价格值提取不正确
- **解决方法**：
  - 软件会自动清理价格中的非数字字符
  - 如果价格格式特殊，可能需要手动调整

### 问题：时间相关字段混淆
- **解决方法**：
  - 在键名文件中使用更明确的名称，比如用"开标日期"代替"时间"
  - 软件已针对常见时间格式优化识别

### 问题：追加模式下某些文件夹被跳过
- **解决方法**：
  - 这是正常的保护机制，只有包含有效"采购项目名称"的项目才会被添加
  - 检查被跳过的文件夹中的PDF是否包含采购项目名称信息

### 问题：进度条显示完成但处理仍在继续
- **解决方法**：
  - 等待状态信息更新为"导出完成"或"已成功新增数据到Excel"
  - 大文件或多文件处理可能需要更多时间

## 6. 注意事项

- **关于键名文件**：键名文件的配置直接影响提取效果，建议先使用示例键名文件测试，再根据需要调整
- **关于PDF格式**：文本PDF(可复制文本的PDF)提取效果最好，扫描PDF可能无法正确提取
- **关于Excel标题行**：选择现有Excel时，请确保输入正确的标题行号(通常为1)
- **关于价格处理**：所有价格类信息(如控制价、预算金额)会自动只保留数字和逗号
- **关于数据追加**：追加模式下，只有采购项目名称不为空的项目才会被添加到Excel
- **系统要求**：本软件兼容Windows 7及以上操作系统
