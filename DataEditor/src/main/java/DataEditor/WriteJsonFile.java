package DataEditor;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Iterator;

/**
 * Created by muskong on 2016/8/25.
 */
public class WriteJsonFile {
    static Logger logger = Logger.getLogger(WriteJsonFile.class);

    public static void fileToJson(XSSFSheet sheet){

        Iterator rows = sheet.rowIterator();

        XSSFRow row;//行
        XSSFCell cell;//列

        // Start constructing JSON.
        JSONObject json = new JSONObject();

        // Iterate through the rows.
        JSONArray jsonRows = new JSONArray();

        int rowIndex = 0;


        while (rows.hasNext())//读取一行数据
        {
            row=(XSSFRow) rows.next();
            JSONObject jRow = new JSONObject();

            // Iterate through the cells.
            JSONArray jsonCells = new JSONArray();

            Iterator cells = row.cellIterator();
            while (cells.hasNext())//读取每列数据
            {
                cell=(XSSFCell) cells.next();
                if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING)
                {
                    jsonCells.add(cell.getStringCellValue());
                }
                else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC)
                {
                    jsonCells.add(cell.getNumericCellValue());
                }
                else {
                    jsonCells.add(null);
                    logger.info(" 无效数据！ ");
                }

            }
            jRow.put("cell", jsonCells);
            jsonRows.add(jRow);
        }

        json.put("rows", jsonRows);

        System.out.println(json.toJSONString());
    }
}
