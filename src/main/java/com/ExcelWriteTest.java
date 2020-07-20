package com;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Date;

/**
 * poi学习
 *
 * @author zhul
 */

public class ExcelWriteTest {

    String path = "D:\\Code\\zhul-poi";

    @Test
    public void testWrite03() throws Exception {
        //1.创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("客户统计表");
        //3.创建一个行
        Row row1 = sheet.createRow(0);
        //4.创建一个单元格
        row1.createCell(0);
        //(1,1)
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增客 户");
        //(1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        //第二行
        //(2,1)
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表(io流)
        FileOutputStream fileOutputStream = new FileOutputStream(path + "\\新增用户03.xls");

        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();

        System.out.println("生成完成！");
    }

    @Test
    public void testWrite07() throws Exception {
        //1.创建一个工作簿 07
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("客户统计表");
        //3.创建一个行
        Row row1 = sheet.createRow(0);
        //4.创建一个单元格
        row1.createCell(0);
        //(1,1)
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增客 户");
        //(1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        //第二行
        //(2,1)
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表(io流)
        FileOutputStream fileOutputStream = new FileOutputStream(path + "\\新增用户07.xlsx");

        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();

        System.out.println("生成完成！");
    }


    @Test
    public void testWrite03BigData() throws Exception {
        long start = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建页码
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNumber = 0; rowNumber < 65536; rowNumber++) {
            Row row = sheet.createRow(rowNumber);
            for (int cellNumber = 0; cellNumber < 10; cellNumber++) {
                Cell cell = row.createCell(cellNumber);
                cell.setCellValue(cellNumber);
            }
        }
        System.out.println("over!");
        FileOutputStream fileOutputStream = new FileOutputStream(path + "\\testWrite03BigData.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-start)/1000);
    }

    /**
     *  07xlsx 使用XSSFWorkbook导出数据，导出大数据时效率低
     * @throws Exception
     */
    @Test
    public void testWrite07BigData() throws Exception {
        long start = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        //创建页码
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNumber = 0; rowNumber < 65537; rowNumber++) {
            Row row = sheet.createRow(rowNumber);
            for (int cellNumber = 0; cellNumber < 10; cellNumber++) {
                Cell cell = row.createCell(cellNumber);
                cell.setCellValue(cellNumber);
            }
        }
        System.out.println("over!");
        FileOutputStream fileOutputStream = new FileOutputStream(path + "\\testWrite07BigData.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-start)/1000);
    }

    /**
     *  07xlsx 使用SXSSFWorkbook导出大数据，效率高
     * @throws Exception
     */
    @Test
    public void testWrite07SxssfBigData() throws Exception {
        long start = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new SXSSFWorkbook();
        //创建页码
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNumber = 0; rowNumber < 65537; rowNumber++) {
            Row row = sheet.createRow(rowNumber);
            for (int cellNumber = 0; cellNumber < 10; cellNumber++) {
                Cell cell = row.createCell(cellNumber);
                cell.setCellValue(cellNumber);
            }
        }
        System.out.println("over!");
        FileOutputStream fileOutputStream = new FileOutputStream(path + "\\testWrite07SxssfBigData.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        //清除临时文件!
        ((SXSSFWorkbook)workbook).dispose();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-start)/1000);
    }


}
