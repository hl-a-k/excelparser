import com.alibaba.fastjson.JSONArray;
import org.apache.commons.io.IOUtils;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;

public class BookInfoExcelParserTests {
    @Test
    void excelParserSuccessTest() throws IOException, InstantiationException, IllegalAccessException {
        String destDirPath = "res";
        File excelFile = new File("res/bookInfo.xls");
        String errorExcel = "error/error.xls";
        BookInfoExcelParser parser = new BookInfoExcelParser(destDirPath);
        JSONArray cfg = JSONArray.parseArray(IOUtils.toString(BookInfoExcelParserTests.class.getClassLoader().getResourceAsStream("bookInfo.json")));
        boolean isValid = parser.verify(excelFile, errorExcel, cfg);
        Assertions.assertTrue(isValid);
        List<HashMap> list = parser.extract(HashMap.class);
        Assertions.assertEquals(4, list.size());
    }

    @Test
    void excelParserFailTest() throws IOException, InstantiationException, IllegalAccessException {
        String destDirPath = "res1";
        File excelFile = new File("res1/bookInfo.xls");
        String errorExcel = "error/error.xls";
        File errorFile = new File(errorExcel);
        if(errorFile.isFile()) {
            errorFile.delete();
        }
        BookInfoExcelParser parser = new BookInfoExcelParser(destDirPath);
        JSONArray cfg = JSONArray.parseArray(IOUtils.toString(BookInfoExcelParserTests.class.getClassLoader().getResourceAsStream("bookInfo.json")));
        boolean isValid = parser.verify(excelFile, errorExcel, cfg);
        Assertions.assertFalse(isValid);


        Assertions.assertTrue(errorFile.isFile());

    }

}
