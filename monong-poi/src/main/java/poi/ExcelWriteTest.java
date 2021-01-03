package poi;

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
import java.io.IOException;


public class ExcelWriteTest {
    String Path = "D:\\IdeaProjects\\excel\\";

    @Test
    public void testWrite03() throws Exception {
        //1创建一个工作簿03
        Workbook workbook = new HSSFWorkbook();
        //2创建一个工作表
        Sheet sheet = workbook.createSheet("统计表");
        //3创建一个行 （1，1）
        Row row1 = sheet.createRow(0);
        //4创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增人数");
        //（1，2）
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        //创建第二 （2，1）
        Row row2 = sheet.createRow(1);
        //创建一个单元格
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        //（2，2）
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表（io流）  03版本就是使用xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(Path+"03统计表.xls");
        workbook.write((fileOutputStream));
        //关闭流
        fileOutputStream.close();

        System.out.println("输出excel成功");
    }
    @Test
    public void testWrite07() throws Exception {
        //1创建一个工作簿07
        Workbook workbook = new XSSFWorkbook();
        //2创建一个工作表
        Sheet sheet = workbook.createSheet("统计表");
        //3创建一个行 （1，1）
        Row row1 = sheet.createRow(0);
        //4创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增人数");
        //（1，2）
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        //创建第二 （2，1）
        Row row2 = sheet.createRow(1);
        //创建一个单元格
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        //（2，2）
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成一张表（io流）  07版本就是使用xlsx结尾
        FileOutputStream fileOutputStream = new FileOutputStream(Path+"统计表07.xlsx");
        workbook.write((fileOutputStream));
        //关闭流
        fileOutputStream.close();

        System.out.println("输出excel成功");
    }

    @Test
    public void testWrite03BigData() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个簿
        Workbook workbook = new HSSFWorkbook();
        //创建一张表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536; rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0 ;cellNum < 10 ;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream outputStream = new FileOutputStream(Path + "03大文件.xls");
        workbook.write(outputStream);
        outputStream.close();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
    }

    @Test
    public void testWrite07BigData() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个簿
        Workbook workbook = new XSSFWorkbook();
        //创建一张表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536; rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0 ;cellNum < 10 ;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream outputStream = new FileOutputStream(Path + "07大文件.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
    }

    @Test
    public void testWrite07BigDatas() throws IOException {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个簿
        Workbook workbook = new SXSSFWorkbook();
        //创建一张表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536; rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0 ;cellNum < 10 ;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream outputStream = new FileOutputStream(Path + "07大文件加速版.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        //清楚临时文件
        ((SXSSFWorkbook) workbook).dispose();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
    }


}
