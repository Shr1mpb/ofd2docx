#  ofd2docx 使用说明 

程序提供两种ofd2docx的方法 方法1足以应对大部分场景 方法2需要额外安装LibreOffice软件

如果两个方法都转换不了的话 依靠现有的开源的项目或者库也比较难做转换了 最好还是用WPS转一下
转换出的质量视文件会有所不同 复杂的文件可能会出现转换有问题的情况

运行环境： jdk21(其他没试)

1. 准备工作：
   - 将所有要转换的 OFD 文件放入程序目录下的 `ofds` 文件夹内。
   - 确保 `java -jar` 命令执行的目录中包含 `ofds` 文件夹。
   - 安装 LibreOffice（如果选择 y2 转换方式）。

2. 运行程序：
   - 在命令行中执行：`java -jar ofd2docx.jar`
   - 程序会提示输入选项：
     - 输入 `y1`：使用 Aspose-pdf 进行 PDF→DOCX 转换（无需额外依赖）。
     - 输入 `y2`：使用 LibreOffice 进行 PDF→DOCX 转换（需提前安装 LibreOffice）。
     - 输入 `n`：退出程序。

3. 输出结果：
   - 转换后的 DOCX 文件会保存在 `out` 文件夹中。
   - 原始 `ofds` 文件夹中的文件不会被修改。

4. 注意事项：
   - 转换逻辑：OFD → PDF → DOCX（两步转换）。
   - 转换质量可能因文件内容而异。
   - 如果选择 `y2`，请确保 LibreOffice 已正确安装并可用。


fc-cache -fv