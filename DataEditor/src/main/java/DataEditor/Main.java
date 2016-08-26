package DataEditor;

import org.apache.log4j.Logger;

import java.io.IOException;

/**
 * Created by muskong on 2016/8/25.
 */
public class Main {
    static Logger logger = Logger.getLogger(Main.class);
    public static void main(String[] args) throws IOException {
        logger.info("read data beginning!");

//        WriteExcelFile.writeXLSFile();
//        ReadExcelFile.readXLSFile();

//        WriteExcelFile.writeXLSXFile();
        ReadExcelFile.readXLSXFileToJson();
//        ReadExcelFile.readXLSXFile();

        logger.info("read data end!");
    }
}
