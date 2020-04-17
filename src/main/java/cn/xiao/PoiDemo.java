package cn.xiao;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author：xiao
 * @date：Created in 2020/4/14 20:44
 * @description：测试导入导出
 */
public class PoiDemo {
    public static void main(String[] args) throws Exception {

        // getExcel(); // 获取单元格数据

        setExcel(); // 写入

    }

    public static void setExcel() throws Exception {
        Workbook wb = new XSSFWorkbook();
        FileOutputStream fs = new FileOutputStream("f:\\test2");
        Sheet user = wb.createSheet("用户表");
        wb.createSheet("订单表");
        // 创建行
        Row row = user.createRow(0);
        // 创建单元格
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("name");
        row.createCell(2).setCellValue("age");
        row.createCell(3).setCellValue("desc");
        // 创建行
        for (int i = 0; i < 10; i++) {
            Row row1 = user.createRow(i + 1);
            row1.createCell(0).setCellValue(1000 + i);
            row1.createCell(1).setCellValue("lina" + i);
            row1.createCell(2).setCellValue(10 + i);
            row1.createCell(3).setCellValue("betf" + i);
        }
        wb.close();
        fs.close();
        System.out.println("over");

    }

    public static void getExcel() throws IOException {
        /**
         * poi函数库中：
         *      工作簿：XSSWorkbook
         *      工作表：XSSFSheet
         *      行和单元格：ROW  / Cell
         */
        //读取一个excel表格中的数据
        //1. 获取excel工作簿对象
        XSSFWorkbook workbook = new XSSFWorkbook("f:\\test.xlsx");
        //2. 获取工作表
        XSSFSheet sheet = workbook.getSheetAt(0);//下标从0开始
        //3. 获取行对象
        for (Row row: sheet) {
            //4. 获取当前行中的单元格中的内容
            for (Cell cell : row) {
                //输出单元格中的内容
                System.out.println(cell);
            }

        }
    }
}
