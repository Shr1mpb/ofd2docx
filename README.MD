# 📄 ofd2docx 使用说明

## 🛠️ 功能说明
程序提供两种ofd转docx的方法：
- **✨ 方法1**：基于Aspose-pdf（🌟推荐使用，满足大部分场景） 转换为纯文本docx时基于apache-poi包
- **⚙️ 方法2**：基于LibreOffice（需额外安装软件）

> ⚠️ **注意**：  
> 🔸 如果两种方法都无法转换，建议使用WPS进行转换  
> 🔸 转换质量因文件复杂度而异，复杂文件可能出现格式问题  
> 🔸 程序不会修改原始文件，所有输出都在`out`文件夹

## 🖥️ 环境要求
- **JDK版本**：21（其他版本未测试）
- **系统要求**：
    - 🐧 Linux环境下请确保安装必要字体
    - 📦 使用方法2需提前安装LibreOffice

## 🚀 使用步骤

### 1. 📂 准备工作
- 1. 使用源码用maven自行打包项目jar包`ofd2docx-1.0-SNAPSHOT-jar-with-dependencies`
  2. 将所有OFD文件放入`ofds`文件夹
  3. 确保执行目录包含`ofds`文件夹
  4. （仅方法1转换不符要求时 使用方法2需要）[下载安装LibreOffice](https://www.libreoffice.org/)

### 2. ⚡ 执行转换
```bash
java -jar ofd2docx-1.0-SNAPSHOT-jar-with-dependencies.jar
```
### 3.🔘转换选项
| 选项 | 功能说明 | 输出格式 | 依赖要求 |
|------|----------|----------|----------|
| y1   | 使用Aspose-pdf完整转换 | DOCX（保留格式） | 无 |
| y1x  | 仅提取文本内容 | TXT（纯文本） | 无 |
| y1d  | 仅提取文本到DOCX | DOCX（纯文本） | 无 |
| y2   | 使用LibreOffice转换 | DOCX（保留格式） | 需LibreOffice |
| n    | 退出程序 | - | - |

### 4.📌注意事项
1. ✅转换文件保存在out文件夹
2. 🔒原始ofds文件夹内容保持不变
3. 🔢自动处理重名文件（添加序号）

### 5.⚠️重要提示
1. 转换大文件时请耐心等待 ⏳
2. 复杂表格/特殊排版可能转换不完美
3. 如遇问题可尝试：
   - 检查文件是否损坏
   - 使用WPS手动转换对比
   - Linux用户检查字体是否安装
4. 由于没有Log4j依赖，输出带有`Log4j API could not find a logging provider`的错误是正常现象，忽略即可

### 6.📄后记
- 本程序应需求而生，编码规范有待提高，都写到了一个类里，也不一定维护
- 如果想贡献代码请先提ISSUE
- 作者联系wx: ndgyy9000