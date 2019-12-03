import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import lombok.AllArgsConstructor;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.codehaus.groovy.runtime.ReflectionMethodInvoker;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.math.BigInteger;
import java.security.MessageDigest;
import java.text.ParseException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Pattern;


public class BookInfoExcelParser {
    List<List<Object>> rows;
    List<Error> errors;
    Map<String, Integer> headMap = new HashMap<>();
    JSONArray cfg;

    private final String[] imgSuffix = new String[]{"jpg", "png"};
    private String destDirPath;


    public BookInfoExcelParser(String destDirPath) {
        this.destDirPath = destDirPath;
    }

    private Integer getCategoryId( String value) {
        return 0;
    }

    private Integer getAuthorId(String value) {
        return 0;
    }

    private void verifyHead() {
        List<Object> head = rows.get(0);
        Set<String> headCfg = new HashSet<>();
        for (int i = 0; i < cfg.size(); i++) {
            headCfg.add(cfg.getJSONObject(i).getString("name"));
        }
        for (int i = 0; i < head.size(); i++) {
            String column = head.get(i).toString();
            if (headCfg.contains(column)) {
                if (headMap.containsKey(column)) {
                    errors.add(new Error(0, i, "列重复!"));
                } else {
                    headMap.put(column, i);
                }
            } else {
                errors.add(new Error(0, i, "未知列!"));
            }
        }
        int rn = head.size();
        for (String h : headCfg) {
            if (h.startsWith("*") && !headMap.containsKey(h)) {
                errors.add(new Error(0, rn++, "必填列!", h));
            }
        }
    }

    Set<String> bookCodes = new HashSet<>();
    String currBookCode = null;

    private Error bookCode(String name, Object value, int r, int c) {
        if (name.startsWith("*") && value == null) {
            return new Error(r, c, "必填!");
        }

        if (bookCodes.contains(value.toString())) {
            return new Error(r, c, "书号重复!");
        } else {
            bookCodes.add(value.toString());
            this.currBookCode = value.toString();
        }
        return null;
    }

    private Error string(String name, Object value, int r, int c) {
        if (name.startsWith("*") && value == null) {
            return new Error(r, c, "必填!");
        }
        return null;
    }

    private Error bigCategory(String name, Object value, int r, int c) {
        if (name.startsWith("*") && value == null) {
            return new Error(r, c, "必填!");
        }
        return null;
    }

    private Error smallCategory(String name, Object value, int r, int c) {
        if (name.startsWith("*") && value == null) {
            return new Error(r, c, "必填!");
        }
        return null;
    }

    private String getImageSuffix() {
        for (String suffix : imgSuffix) {
            File imgFile = new File(destDirPath + "/" + currBookCode + "." + suffix);
            if (imgFile.exists()) {
                return suffix;
            }
        }
        return null;
    }

    private Error coverPath(String name, Object value, int r, int c) {
        String suffix = getImageSuffix();
        if (name.startsWith("*") && suffix == null) {
            return new Error(r, c, "必填!");
        }
        return null;
    }

    private String getFilePath(String prefix) {
        String filePath = null;
        File epubFile = new File(destDirPath + "/" + prefix + currBookCode + ".epub");
        if (epubFile.exists()) {
            filePath = epubFile.getAbsolutePath();
        }
        return filePath;
    }

    private Error filePath(String name, Object value, int r, int c) {
        String filePath = getFilePath("");

        if (name.startsWith("*") && filePath == null) {
            return new Error(r, c, "必填!");
        }
        return null;
    }

    private Error tryFilePath(String name, Object value, int r, int c) {
        String filePath = getFilePath("_");

        if (name.startsWith("*") && filePath == null) {
            return new Error(r, c, "必填!");
        }
        return null;
    }

    private Error Double(String name, Object value, int r, int c) {
        if (name.startsWith("*") && value == null) {
            return new Error(r, c, "必填!");
        }
        if(value != null) {
            try{
                new Double(value.toString());
            } catch (Exception e) {
                return new Error(r, c, "数字格式不正确!");
            }
        }

        return null;
    }

    private Error Int(String name, Object value, int r, int c) {
        if (name.startsWith("*") && value == null) {
            return new Error(r, c, "必填!");
        }
        if(value != null) {
            try{
                new Double(value.toString());
            } catch (Exception e) {
                return new Error(r, c, "数字格式不正确!");
            }
        }
        return null;
    }

    private Error author(String name, Object value, int r, int c) {
        if (name.startsWith("*") && value == null) {
            return new Error(r, c, "必填!");
        }
        if (value != null) {
            if (getAuthorId(value.toString()) == null) {
                return new Error(r, c, "作者名字不正确");
            }
        }
        return null;
    }

    private Error date(String name, Object value, int r, int c) {
        if (name.startsWith("*") && value == null) {
            return new Error(r, c, "必填!");
        }
        Pattern p = Pattern.compile("\\d\\d\\d\\d-\\d\\d-\\d\\d");
        if (!p.matcher(value.toString()).matches()) {
            return new Error(r, c, "日期格式!");
        }
        return null;
    }

    public boolean verify(File excelFile, String errorPath, JSONArray cfg) throws IOException {
        errors = new ArrayList<>();
        rows = ExcelUtil.readExcel2003(excelFile);
        this.cfg = cfg;
        if (rows.size() < 2) {
            errors.add(new Error(0, 0, "Excel无数据", "Excel无数据"));
        } else {
            verifyHead();
            for (int r = 1; r < rows.size(); r++) {
                for (int i = 0; i < cfg.size(); i++) {
                    JSONObject cfgObj = cfg.getJSONObject(i);
                    String name = cfgObj.getString("name");
                    String verifier = cfgObj.getString("verifier");
                    if (verifier == null) {
                        verifier = cfgObj.getString("parser");
                    }
                    if (headMap.containsKey(name)) {
                        int c = headMap.get(name);
                        Object value = null;
                        try{
                            value = rows.get(r).get(c);
                        }catch (Exception e){
//                            e.printStackTrace();
                        }
//                        Object value = rows.get(r).get(c);
                        Error error = (Error) ReflectionMethodInvoker.invoke(this, verifier, new Object[]{name, value, r, c});
                        if (error != null) {
                            errors.add(error);
                        }
                    }
                }
            }
        }
        if (errors.size() != 0) {
            File dst = new File(errorPath);
            HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(excelFile));
            HSSFSheet sheet = hwb.getSheetAt(0);
            HSSFCellStyle style = hwb.createCellStyle();
            style.setFillForegroundColor(IndexedColors.RED.getIndex());
            style.setFillPattern(CellStyle.SOLID_FOREGROUND);

            HSSFPatriarch p = sheet.createDrawingPatriarch();
            for (Error error : errors) {
                HSSFComment comment = p.createComment(new HSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
                //输入批注信息
                comment.setString(new HSSFRichTextString(error.msg));
                Cell cell = sheet.getRow(error.r).getCell(error.c);
                if (cell == null) {
                    cell = sheet.getRow(error.r).createCell(error.c);
                }
                if (error.columnName != null) {
                    cell.setCellValue(error.columnName);
                }
                //将批注添加到单元格对象中
                cell.setCellComment(comment);
                cell.setCellStyle(style);
            }

            hwb.write(new FileOutputStream(dst));
        }
        return errors.size() == 0;
    }

    public void string(Object obj, String field, Object value) throws InvocationTargetException, IllegalAccessException {
        if (value != null) {
            BeanUtils.setProperty(obj, field, value);
        }

    }
    public void bookCode(Object obj, String field, Object value) throws InvocationTargetException, IllegalAccessException {
        if (value != null) {
            BeanUtils.setProperty(obj, field, value);
            this.currBookCode = value.toString();
        }

    }

    public void bigCategory(Object obj, String field, Object value) throws InvocationTargetException, IllegalAccessException {
        BeanUtils.setProperty(obj, field, getCategoryId(value.toString()));
    }

    public void smallCategory(Object obj, String field, Object value) throws InvocationTargetException, IllegalAccessException {
        BeanUtils.setProperty(obj, field, getCategoryId( value.toString()));
    }

    public void coverPath(Object obj, String field, Object value) throws InvocationTargetException, IllegalAccessException, FileNotFoundException {
        String imgSuffix = getImageSuffix();
        if (imgSuffix != null) {
            BeanUtils.setProperty(obj, field,  imgSuffix);
        }

    }

    private String getMd5(File file) {
        try {
            byte[] uploadBytes = FileUtils.readFileToByteArray(file);
            MessageDigest md5 = MessageDigest.getInstance("MD5");
            byte[] digest = md5.digest(uploadBytes);
            String hashString = new BigInteger(1, digest).toString(16);
            return hashString;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    public void filePath(Object obj, String field, Object value) throws InvocationTargetException, IllegalAccessException, FileNotFoundException {
        String filePath = getFilePath("");
        if(filePath != null) {
        }
    }

    public void tryFilePath(Object obj, String field, Object value) throws InvocationTargetException, IllegalAccessException, FileNotFoundException {
        String filePath = getFilePath("_");
        if (filePath != null) {
            File f = new File(filePath);
            BeanUtils.setProperty(obj, field,filePath);
        }
    }

    public void Double(Object obj, String field, Object value) throws InvocationTargetException, IllegalAccessException {
        if (value != null) {
            BeanUtils.setProperty(obj, field, new Double(value.toString()));
        }
    }

    public void Int(Object obj, String field, Object value) throws InvocationTargetException, IllegalAccessException {
        if (value != null) {
            BeanUtils.setProperty(obj, field, new Double(value.toString()).intValue());
        }

    }

    public void author(Object obj, String field, Object value) throws InvocationTargetException, IllegalAccessException {
        if(value != null) {
            BeanUtils.setProperty(obj, field, getAuthorId(value.toString()));
        }
    }

    public void date(Object obj, String name, Object value) throws InvocationTargetException, IllegalAccessException, ParseException {
        if (value != null) {
            DateTimeFormatter df = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
            LocalDateTime ldt = LocalDateTime.parse(value + " 00:00:00",df);

            BeanUtils.setProperty(obj, name, ldt);
        }
    }

    public <T> List<T> extract(Class<T> tclass) throws IllegalAccessException, InstantiationException {
        List<T> list = new ArrayList<>();
        for (int r = 1; r < rows.size(); r++) {
            T obj = tclass.newInstance();
            for (int i = 0; i < this.cfg.size(); i++) {
                JSONObject cfgObj = cfg.getJSONObject(i);
                String name = cfgObj.getString("name");
                String field = cfgObj.getString("field");
                String parser = cfgObj.getString("parser");

                if (headMap.containsKey(name)) {
                    int c = headMap.get(name);
                    Object value = rows.get(r).get(c);
                    ReflectionMethodInvoker.invoke(this, parser, new Object[]{obj, field, value});
                }
            }
            list.add(obj);
        }
        return list;
    }

    @AllArgsConstructor
    static class Error {
        int r, c;
        String msg, columnName;

        public Error(int r, int c, String msg) {
            this.r = r;
            this.c = c;
            this.msg = msg;
        }
    }
}
