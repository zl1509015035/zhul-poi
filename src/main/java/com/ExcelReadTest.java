package com;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Date;

/**
 * @author zhul
 */
public class ExcelReadTest {

    String path = "D:\\Code\\zhul-poi";

    /**
     * 03版本的读excel操作
     *
     * @throws Exception
     */
    @Test
    public void testRead03() throws Exception {

        //获取文件流
        FileInputStream inputStream = new FileInputStream(path + "\\新增用户03.xls");

        //1.创建一个工作簿,并读取流
        Workbook workbook = new HSSFWorkbook(inputStream);
        //2.创建一个工作表，并获取流中的sheet表
        Sheet sheet = workbook.getSheetAt(0);
        //3.创建一个行,并获取流中的行
        Row row = sheet.getRow(0);
        //4.创建一个单元格，获取流中的单元格
        Cell cell = row.getCell(0);

        System.out.println(cell.getStringCellValue());

        inputStream.close();
        System.out.println("生成完成！");
    }

    /**
     * 07版本的读excel操作
     *
     * @throws Exception
     */
    @Test
    public void testRead07() throws Exception {

        //获取文件流
        FileInputStream inputStream = new FileInputStream(path + "\\新增用户07.xlsx");

        //1.创建一个工作簿,并读取流
        Workbook workbook = new XSSFWorkbook(inputStream);
        //2.创建一个工作表，并获取流中的sheet表
        Sheet sheet = workbook.getSheetAt(0);
        //3.创建一个行,并获取流中的行
        Row row = sheet.getRow(0);
        //4.创建一个单元格，获取流中的单元格
        Cell cell = row.getCell(0);

        System.out.println(cell.getStringCellValue());

        inputStream.close();
        System.out.println("生成完成！");
    }

    /**
     * 读取不同类型的数据
     */
    @Test
    public void testCellType() throws Exception {
        //获取文件流
        FileInputStream inputStream = new FileInputStream(path + "\\明细表.xls");

        //1.创建一个工作簿,并读取流
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        //获取标题内容
        Row rowTitle = sheet.getRow(0);
        if (rowTitle != null) {
            //得到这一行有多少列
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    //获取cell的类型
                    int cellType = cell.getCellType();
                    //获取string类型的数值
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + "|");
                }
            }
            System.out.println();
        }

        // 获取表中内容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 0; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {
                //读取列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
//                    System.out.print("["+(rowNum+1)+"-"+(cellNum+1)+"]");
                    Cell cell = rowData.getCell(cellNum);
                    if (cell != null) {
                        int cellType = cell.getCellType();
                        String cellValue = "";
                        switch (cellType) {
                            // 字符串
                            case HSSFCell.CELL_TYPE_STRING:
                                System.out.print("【String】");
                                cellValue = cell.getStringCellValue();
                                break;
                            // 布尔
                            case HSSFCell.CELL_TYPE_BOOLEAN:
                                System.out.print("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            // 空
                            case HSSFCell.CELL_TYPE_BLANK:
                                System.out.print("【BLANK】");
                                break;
                            // 数字（日期、普通数字）
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                System.out.print("【NUMERIC】");
                                // 日期
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                } else {
                                    // 不是日期格式，防止数字过长！
                                    System.out.print("【转换为字符串输出】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print("【数据类型错误】");
                                break;
                            default:
                                throw new IllegalStateException("Unexpected value: " + cellType);
                        }
                        System.out.print(cellValue);
                    }
                }
            }
        }
        inputStream.close();
    }

    /**
     * 工具类，传入一个文件流，进行操作（地址）
     *
     * @param inputStream
     * @throws Exception
     */
    public void testCellType(FileInputStream inputStream) throws Exception {
        //1.创建一个工作簿,并读取流
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        //获取标题内容
        Row rowTitle = sheet.getRow(0);
        if (rowTitle != null) {
            //得到这一行有多少列
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    //获取cell的类型
                    int cellType = cell.getCellType();
                    //获取string类型的数值
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + "|");
                }
            }
            System.out.println();
        }

        // 获取表中内容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 0; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {
                //读取列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
//                    System.out.print("["+(rowNum+1)+"-"+(cellNum+1)+"]");
                    Cell cell = rowData.getCell(cellNum);
                    if (cell != null) {
                        int cellType = cell.getCellType();
                        String cellValue = "";
                        switch (cellType) {
                            // 字符串
                            case HSSFCell.CELL_TYPE_STRING:
                                System.out.print("【String】");
                                cellValue = cell.getStringCellValue();
                                break;
                            // 布尔
                            case HSSFCell.CELL_TYPE_BOOLEAN:
                                System.out.print("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            // 空
                            case HSSFCell.CELL_TYPE_BLANK:
                                System.out.print("【BLANK】");
                                break;
                            // 数字（日期、普通数字）
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                System.out.print("【NUMERIC】");
                                // 日期
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                } else {
                                    // 不是日期格式，防止数字过长！
                                    System.out.print("【转换为字符串输出】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print("【数据类型错误】");
                                break;
                            default:
                                throw new IllegalStateException("Unexpected value: " + cellType);
                        }
                        System.out.print(cellValue);
                    }
                }
            }
        }
        inputStream.close();
    }

}




