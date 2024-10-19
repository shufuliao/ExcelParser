package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static void main(String[] args) {
        // 創建文件選擇器
        JFileChooser fileChooser = new JFileChooser();
        // 只允許選擇 .xlsx 文件
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx");
        fileChooser.setFileFilter(filter);

        // 顯示打開文件對話框
        int returnValue = fileChooser.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            String filePath = selectedFile.getAbsolutePath();

            try {
                FileInputStream file = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(file);
                ExcelParser parser = new ExcelParser(workbook);
                parser.run();
                workbook.close();
                file.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("沒有選擇文件");
        }
    }
}