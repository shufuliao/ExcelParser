package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelParser {
    private final Workbook workbook;

    public ExcelParser(Workbook workbook) {
        this.workbook = workbook;
    }

    public void run() {
        readSheet();
    }

    private void readSheet() {
        Sheet sheet = workbook.getSheetAt(0);
        String currentBenefit = "";
        String currentCoverage = "";

        System.out.println("[Benefit, Coverage, Category, Plan Name, Coverage Value]");

        for (Row row : sheet) {
            Cell firstCell = row.getCell(0);

            if (isChangeBenefitCell(firstCell)) {
                currentBenefit = getCellValueAsString(firstCell);
            }
            else if (isChangeCoverage(firstCell)) {
                currentCoverage = getCellValueAsString(firstCell);
            }

            Cell categoryCell = row.getCell(1);
            if (categoryCell == null) {continue;}

            String category = getCellValueAsString(categoryCell);

            // 從 F 欄開始處理 Plan 資料
            extractPlanData(sheet, row, currentBenefit, currentCoverage, category);
        }
    }

    private boolean isChangeBenefitCell(Cell cell) {
        boolean isBenefit = false;

        if (cell != null) {
            isBenefit = cell.getCellStyle().getFillForegroundColor() == 0;
        }

        return isBenefit;
    }

    private boolean isChangeCoverage(Cell cell) {
        boolean isCoverage = false;

        if (cell != null) {
            isCoverage = cell.getCellStyle().getFillForegroundColor() != 0;
        }

        return isCoverage;
    }

    private void extractPlanData(Sheet sheet, Row row, String benefit, String coverage, String category) {
        for (int i = 5; i < row.getLastCellNum(); i++) {
            Cell planNameCell = sheet.getRow(0).getCell(i);
            Cell coverageNameCell = row.getCell(i);

            if (planNameCell != null && coverageNameCell != null) {
                String planName = getCellValueAsString(planNameCell);
                String coverageName = getCellValueAsString(coverageNameCell);

                printFormattedData(benefit, coverage, category, planName, coverageName);
            }
        }
    }

    private static void printFormattedData(String benefit, String coverage, String category, String planName, String coverageName) {
        System.out.printf("%s, %s, %s, %s, %s\n", benefit, coverage, category, planName, coverageName);
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        // 根據單元格類型來處理
        CellType cellType = cell.getCellType();
        return switch (cellType) {
            case STRING ->
                cell.getStringCellValue();
            case NUMERIC ->
                // 處理數字類型（包括浮點數）
                String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN ->
                // 處理布爾類型
                String.valueOf(cell.getBooleanCellValue());
            case FORMULA ->
                // 如果是公式，則返回公式結果
                String.valueOf((int) cell.getNumericCellValue());
            default -> "";
        };
    }
}
