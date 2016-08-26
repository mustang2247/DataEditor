package DataEditor;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

/**
 * Created by muskong on 2016/8/25.
 */
public class ReadExcelFile {

    static Logger logger = Logger.getLogger(ReadExcelFile.class);
    /**
     * 读取xlsx格式
     * @throws IOException
     */
    public static void readXLSXFile() throws IOException
    {
        logger.info("readXLSXFile begin!");
        System.out.println("\n");
        InputStream ExcelFileToRead = new FileInputStream("/Users/muskong/随机任务配置.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row;//行
        XSSFCell cell;//列

        Iterator rows = sheet.rowIterator();
        int rowNum = 0;

        while (rows.hasNext())
        {
            logger.debug("readXLSXFile rows:  " + rowNum);
//            System.out.println("\n");

            row=(XSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext())
            {
                cell=(XSSFCell) cells.next();

                if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING)
                {
                    logger.info(cell.getStringCellValue()+" ");
                }
                else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC)
                {
                    logger.info(cell.getNumericCellValue()+" ");
                }
                else
                {
                    //U Can Handel Boolean, Formula, Errors
                }
            }
            System.out.println("\n");
            rowNum++;
        }

    }

    public static void readXLSXFileToJson() throws IOException
    {
        InputStream ExcelFileToRead = new FileInputStream("/Users/muskong/随机任务配置.xlsx");
        logger.info("readXLSXFile  begin!");
        System.out.println("\n");
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

        if(wb.getNumberOfSheets() <= 0) return;

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            int num = sheet.getPhysicalNumberOfRows();
            logger.info(num);
            if(num <= 2) {
                logger.info("data sheet error!!");
                continue;
            }
            System.out.println(sheet.getPhysicalNumberOfRows());
            WriteJsonFile.fileToJson(sheet);
        }

    }


    /**
     * 读取xls格式
     * @throws IOException
     */
    public static void readXLSFile() throws IOException
    {
        InputStream ExcelFileToRead = new FileInputStream("C:/Test.xls");
        HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

        HSSFSheet sheet=wb.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;

        Iterator rows = sheet.rowIterator();

        while (rows.hasNext())
        {
            row=(HSSFRow) rows.next();
            Iterator cells = row.cellIterator();

            while (cells.hasNext())
            {
                cell=(HSSFCell) cells.next();

                if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING)
                {
                    System.out.print(cell.getStringCellValue()+" ");
                }
                else if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC)
                {
                    System.out.print(cell.getNumericCellValue()+" ");
                }
                else
                {
                    //U Can Handel Boolean, Formula, Errors
                }
            }
            System.out.println();
        }

    }

}
