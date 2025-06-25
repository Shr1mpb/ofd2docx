package cn.edu.twt;

import com.aspose.pdf.Document;
import com.aspose.pdf.SaveFormat;
import org.ofdrw.converter.export.PDFExporterIText;
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class Application {
    public static void main(String[] args) throws InterruptedException {
        // 打印欢迎信息
        printWelcomeMessage();

        // 获取用户输入
        Scanner scanner = new Scanner(System.in);
        String input = scanner.nextLine().trim().toLowerCase();

        // 处理用户选择
        if (input.equals("y1")) {
            System.out.println("开始转换...");
            convertOFDs(true);
        }else if (input.equals("y2")) {
            System.out.println("开始转换...");
            convertOFDs(false);
        }else if (input.equals("n")) {
            System.out.println("程序已退出");
            System.exit(0);
        } else {
            System.out.println("无效输入，程序已退出");
            System.exit(1);
        }
        scanner.close();
    }

    private static void printWelcomeMessage() {
        System.out.println("====================ofd2docx====================");
        System.out.println("请确认以下信息：");
        System.out.println("1. 将所有要转换的ofd文件放入ofds文件夹内");
        System.out.println("2. java -jar命令执行的目录中包含ofds文件夹");
        System.out.println("3. 程序不会影响原有的ofds文件夹中的文件，输出位于out文件夹中");
        System.out.println("4. 转换后的docx文件将保存在out文件夹中");
        System.out.println("5. 转换逻辑为：第一步先ofd转换为pdf，第二步再从pdf转换为docx");
        System.out.println("6. ofd转换为docx的转换质量可能会因文件而异");
        System.out.println("7. 【关键】输入y1使用Aspose-pdf包进行第二步转换 输入y2使用LibreOffice软件进行第二步转换");
        System.out.println("8. 【关键】y2的第二步转换依赖LibreOffice 请先提前安装");
        System.out.println("====================ofd2docx====================");
        System.out.print("确认无误后输入y开始转换，输入n退出程序 (y1/y2/n)? ");
    }

    /**
     * 转换ofd文件为docx文件
     */
    private static void convertOFDs(boolean y1) throws InterruptedException {
        String execDir = System.getProperty("user.dir");
        File ofdDir = new File(execDir, "ofds");
        File outDir = new File(execDir, "out");

        // 检查输入目录
        if (!ofdDir.exists() || !ofdDir.isDirectory()) {
            System.err.println("错误：ofds文件夹不存在或不是目录");
            System.exit(1);
        }

        // 创建输出目录
        if (!outDir.exists() && !outDir.mkdirs()) {
            System.err.println("错误：out文件夹不存在且无法创建out文件夹");
            System.exit(1);
        }

        // 获取OFD文件列表
        File[] ofdFiles = ofdDir.listFiles((dir, name) -> name.toLowerCase().endsWith(".ofd"));

        if (ofdFiles == null || ofdFiles.length == 0) {
            System.out.println("警告：ofds文件夹中没有找到.ofd文件");
            return;
        }

        System.out.println("找到 " + ofdFiles.length + " 个OFD文件待转换");

        // 开始转换
        int successCount = 0;
        boolean alreadyConverted = false;

        List<File> needDeleteFiles = new ArrayList<>();
        for (File ofdFile : ofdFiles) {
            try {
                String originalName = ofdFile.getName();
                Path docxPath = Paths.get(outDir.getPath(), originalName.replace(".ofd", ".docx"));
                Path pdfPath = Paths.get(outDir.getPath(), originalName.replace(".ofd", ".pdf"));

                System.out.print("正在转换: " + originalName + "... ");
                if(Files.exists(docxPath)) {
                    successCount++;
                    alreadyConverted = true;
                    System.out.println("已转换过");
                    continue;
                }

                // 转换为PDF
                try (PDFExporterIText pdfConverter = new PDFExporterIText(ofdFile.toPath(), pdfPath)) {
                    pdfConverter.export();
                } catch (Exception e) {
                    needDeleteFiles.add(pdfPath.toFile());
                    throw new RuntimeException("转换为PDF时失败" + e.getMessage());
                }
                // 把pdf转换为docx
                // 两种方法
                if (y1){
                    try (Document pdfDocument = new Document(pdfPath.toString())) {
                        pdfDocument.save(docxPath.toString(), SaveFormat.DocX);
                    } catch (Exception e) {
                        throw new RuntimeException("pdf转换为docx时失败" + e.getMessage());
                    }
                } else{
                    try {
                        // 调用 LibreOffice 命令行转换
                        convertPdfToDocxUsingLibreOffice(pdfPath.toString(), docxPath.toString());
                    } catch (Exception e) {
                        needDeleteFiles.add(docxPath.toFile());
                        throw new RuntimeException("pdf转换为docx时失败: " + e.getMessage());
                    }
                }

                System.out.println("成功");
                // 成功后删除pdf中间文件
                Files.deleteIfExists(pdfPath);
                successCount++;

            } catch (Exception e) {
                Path docxPath = Paths.get(outDir.getPath(), ofdFile.getName().replace(".ofd", ".docx"));
                needDeleteFiles.add(docxPath.toFile());
                System.err.println("\n转换失败: " + ofdFile.getName() + " | 原因: " + e.getMessage());
            }
        }


        // 输出统计信息
        System.out.println("\n==========转换完成==========");
        // 删除无用文件
        System.out.println("正在删除pdf中间文件与转换失败的文件...");
        System.gc();
        Thread.sleep(1000);// 等待1s 资源回收后删除 避免被占用
        for (File file : needDeleteFiles) {
            try{
                Files.delete(file.toPath());
            }catch (Exception e){
                System.out.println("删除 " + file.getName() + " 失败" + " | 原因: " + e.getMessage());
            }

        }
        if (alreadyConverted) {
            System.out.println("已转换过的文件被算在转换成功个数当中");
        }
        System.out.println("成功: " + successCount + " 个");
        System.out.println("失败: " + (ofdFiles.length - successCount) + " 个");
        System.out.println("输出目录: " + outDir.getAbsolutePath());
    }

    /**
        * 使用 LibreOffice 命令行将 PDF 转换为 DOCX
        * @param pdfPath 输入 PDF 文件路径
        * @param docxPath 输出 DOCX 文件路径
    */
    private static void convertPdfToDocxUsingLibreOffice(String pdfPath, String docxPath) throws IOException, InterruptedException {
        // 构造 LibreOffice 命令
        String libreOfficePath = getLibreOfficePath(); // 自动获取安装路径
        String command = String.format(
                "\"%s\" --infilter=\"writer_pdf_import\" --convert-to docx \"%s\" --outdir \"%s\"",
                libreOfficePath,
                pdfPath,
                Paths.get(pdfPath).getParent().toString()
        );

        // 执行命令
        Process process = Runtime.getRuntime().exec(command);
        int exitCode = process.waitFor();

        // 检查是否成功
        if (exitCode != 0) {
            try (BufferedReader errorReader = new BufferedReader(
                    new InputStreamReader(process.getErrorStream()))) {
                StringBuilder errorMsg = new StringBuilder();
                String line;
                while ((line = errorReader.readLine()) != null) {
                    errorMsg.append(line).append("\n");
                }
                throw new IOException("LibreOffice 转换失败 (Exit Code: " + exitCode + "): " + errorMsg);
            }
        }
    }

        /**
            * 获取 LibreOffice 可执行文件路径（跨平台支持）
         */
    private static String getLibreOfficePath() {
        String osName = System.getProperty("os.name").toLowerCase();
        String defaultPath = "soffice"; // Linux/macOS 默认在 PATH 中

        if (osName.contains("win")) {
            // Windows 常见安装路径
            String[] possiblePaths = {
                    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
                    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
                    "D:\\LibreOffice\\program\\soffice.exe"
            };
            for (String path : possiblePaths) {
                if (Files.exists(Paths.get(path))) {
                    return path;
                }
            }
            System.err.println("未找到 LibreOffice 安装路径，请确认已安装");
            System.exit(1);
        }
        return defaultPath;
    }
}
