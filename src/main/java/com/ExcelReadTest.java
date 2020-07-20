package com;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelReadTest {

    String path = "D:\\Code\\zhul-poi";

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

}
