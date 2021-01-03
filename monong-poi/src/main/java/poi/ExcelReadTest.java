package poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelReadTest {
    String Path = "D:\\IdeaProjects\\excel\\";

    @Test
    public void testRead03() throws Exception {
        //读取一个工作簿03
        FileInputStream fileInputStream = new FileInputStream(Path + "excel03统计表.xls");
        //创建一个工作簿  可以像execl表一样操作
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        //得到表
        Sheet sheet = workbook.getSheetAt(0);
        //得到行
        Row row = sheet.getRow(0);
        //得到列
        Cell cell = row.getCell(1);
        //读取值的时候，一定要注意类型
        System.out.println(cell.getNumericCellValue());
        fileInputStream.close();
    }

    @Test
    public void testRead07() throws Exception {
        //读取一个工作簿07
        FileInputStream fileInputStream = new FileInputStream(Path + "excel统计表07.xlsx");
        //创建一个工作簿  可以像execl表一样操作
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        //得到表
        Sheet sheet = workbook.getSheetAt(0);
        //得到行
        Row row = sheet.getRow(0);
        //得到列
        Cell cell = row.getCell(1);

        //读取值的时候，一定要注意类型
        System.out.println(cell.getNumericCellValue());
        fileInputStream.close();
    }

    @Test
    public void testCellType() throws IOException {
        //读取一个工作簿03
        FileInputStream fileInputStream = new FileInputStream(Path + "多数据类型表03.xls");
        //创建一个工作簿  可以像execl表一样操作
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        workbook.getSheetAt(0);
        //获取标题内容

    }
}
