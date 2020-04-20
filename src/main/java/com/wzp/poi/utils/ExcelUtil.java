package com.wzp.poi.utils;

import com.wzp.poi.bean.Student;
import com.wzp.poi.bean.User;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.ObjectUtils;
import org.springframework.util.StringUtils;

import java.io.*;
import java.lang.reflect.Field;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DecimalFormat;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Slf4j
public class ExcelUtil<T> {

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";


    /**
     * 导出数据，生成excel
     *
     * @param dataList   数据源
     * @param clazz      实体类类型
     * @param sheetName  表名
     * @param ColumnName 字段名
     * @param filename   生成的文件名
     * @throws NoSuchFieldException
     * @throws IllegalAccessException
     */

    public static void export(List dataList, Class<?> clazz, String sheetName, List<String> ColumnName, String filename) throws NoSuchFieldException, IllegalAccessException {
        HSSFWorkbook workBook = new HSSFWorkbook(); //创建一个excel
        HSSFSheet sheet = workBook.createSheet(sheetName); //创建一个工作簿
        sheet.setDefaultColumnWidth(10 * 256);    //单元格宽度
        sheet.setDefaultRowHeight((short) (256));
        HSSFDataFormat format = workBook.createDataFormat();
        HSSFRow row = sheet.createRow(0);
        row.setHeight((short) (256));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, ColumnName.size() - 1));
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("学生信息");
        cell.setCellStyle(getFormat(workBook, null, null, IndexedColors.RED.index, true, true, false, false, null));
        HSSFRow row1 = sheet.createRow(1);
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        StringBuilder stringBuilder = new StringBuilder();
        for (int i = 0; i < ColumnName.size(); i++) {
            HSSFCell column = row1.createCell(i);
            column.setCellValue(ColumnName.get(i));
        }
        for (int i = 0; i < dataList.size(); i++) {
            HSSFRow data = sheet.createRow(i + 2); //创建一行
            for (int j = 0; j < ColumnName.size(); j++) { //遍历属性名
                Field field = null;
                String[] strings = ColumnName.get(j).split("_");
                for (String string : strings) {
                    if (stringBuilder.length() == 0) {
                        stringBuilder.append(string.toLowerCase());
                    } else {
                        stringBuilder.append(string.substring(0, 1).toUpperCase());
                        stringBuilder.append(string.substring(1).toLowerCase());
                    }
                }
                field = clazz.getDeclaredField(stringBuilder.toString()); //反射获取对应实体类中的属性
                field.setAccessible(true);
                HSSFCell element = data.createCell(j);
                if (field.getGenericType().toString().equals("class java.lang.String")) {
                    element.setCellValue((String) field.get(dataList.get(i)));
                    if (field.getName().equals("sex") && ((field.get(dataList.get(i))).equals("女"))) {
                        element.setCellStyle(getFormat(workBook, null, null, null, false, false, false, false, IndexedColors.PINK.index));
                    }
                } else if (field.getName().toLowerCase().contains("date")) {
                    element.setCellValue(simpleDateFormat.format(field.get(dataList.get(i))));
                } else if (field.getType().getName().toLowerCase().contains("integer") || field.getType().getName().toLowerCase().contains("long")) {
                    element.setCellValue((Long) field.get(dataList.get(i)));
                }
                stringBuilder.delete(0, stringBuilder.length());

            }
        }
        sheet.autoSizeColumn(4); //自适应列宽度
        sheet.autoSizeColumn(5);
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream("D:\\" + filename + ".xlsx");
            workBook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 生成导入模板
     *
     * @// TODO: 2019/7/7 字段需要从配置文件中获取（K-V结构）
     */

    public static void excelTemplate() {
        HSSFWorkbook workBook = new HSSFWorkbook(); //创建一个excel
        HSSFSheet sheet = workBook.createSheet("批量导入学生"); //创建一个工作簿
        HSSFRow row = sheet.createRow(0);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("导入学生");
        HSSFRow row1 = sheet.createRow(1);
        String[] strings = {"姓名", "性别", "住址", "联系方式"};
        for (int i = 0; i < 4; i++) {
            HSSFCell property = row1.createCell(i);
            property.setCellValue(strings[i]);
        }
        //下拉菜单
        CellRangeAddressList regions = new CellRangeAddressList(2, 65535, 1, 1);  //范围
        DVConstraint constraint = DVConstraint.createExplicitListConstraint(new String[]{"男", "女"});
        HSSFDataValidation dataValidation = new HSSFDataValidation(regions, constraint);
        sheet.addValidationData(dataValidation);
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream("D:\\excelTemplate.xlsx");
            workBook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static List<User> intoDatabase() throws Exception {
        File file = new File("D:\\excelTemplate.xlsx");
        Workbook workbook = getWorkbook(new FileInputStream(file),file.getName());
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;
        List<User> userList = new ArrayList<>();
        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            sheet = workbook.getSheetAt(sheetNum);
            if (sheet == null) {
                continue;
            }
            //遍历当前sheet中的所有行,去掉表头
            for (int rowNum = 2; rowNum < sheet.getLastRowNum() + 1; rowNum++) {
                List<Object> temp = new ArrayList<>();
                //读取一行
                row = sheet.getRow(rowNum);
                //遍历所有列
                for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                    cell = row.getCell(cellNum);
                    temp.add(getCellValue(cell));
                }
                User user = new User();
                user.setName((String) temp.get(0));
                user.setSex((String) temp.get(1));
                user.setAddress((String) temp.get(2));
                user.setPhone(((Double) temp.get(3)).longValue());
                user.setCreateDate(System.currentTimeMillis());
                user.setLastLoginDate(System.currentTimeMillis());
                userList.add(user);
            }
        }
        return userList;
    }


    private static Workbook getWorkbook(InputStream inputStream, String fileName) throws Exception {
        Workbook workbook = null;
        String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
        if (XLSX.equals(suffix)) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (XLS.equals(suffix)) {
            workbook = new HSSFWorkbook(inputStream);
        }
        if (workbook == null) {
            throw new Exception("文件后缀名必须为xlsx或xls");
        }
        return workbook;
    }

    /**
     * 设置的单元格格式
     *
     * @param FontName            字体
     * @param fontHeightInPoints  字号
     * @param color               文字颜色
     * @param alignment           水平居中
     * @param verticalAlignment   垂直居中
     * @param italic              是否斜体
     * @param bold                是否加粗
     * @param fillForegroundColor 单元格背景色
     * @return
     */

    private static HSSFCellStyle getFormat(HSSFWorkbook workbook, String FontName, Short fontHeightInPoints, Short color, boolean alignment, boolean verticalAlignment,
                                           boolean italic, boolean bold, Short fillForegroundColor) {

        HSSFCellStyle cellStyle = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        if (!StringUtils.isEmpty(FontName)) {
            font.setFontName(FontName);
        }
        if (!ObjectUtils.isEmpty(fontHeightInPoints)) {
            font.setFontHeightInPoints(fontHeightInPoints);
        }
        if (!ObjectUtils.isEmpty(color)) {
            font.setColor(color);
        }
        if (!ObjectUtils.isEmpty(alignment) && alignment) {
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
        }
        if (!ObjectUtils.isEmpty(verticalAlignment) && verticalAlignment) {
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        }
        if (!ObjectUtils.isEmpty(fillForegroundColor)) {
            cellStyle.setFillForegroundColor(fillForegroundColor);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        font.setItalic(italic);
        font.setBold(bold);
        cellStyle.setFont(font);
        return cellStyle;
    }

    public static List<Student> readExcel(File file) throws Exception {

        FileInputStream inputStream = null;
        Workbook workbook = null;
        try {
            if (!file.exists()) {
                throw new Exception("文件不存在");
            }

            // 获取Excel工作簿
            inputStream = new FileInputStream(file);
            workbook = getWorkbook(inputStream, file.getName());

            // 读取excel中的数据
            return parseExcel(workbook);
        } catch (Exception e) {
            log.error("解析Excel失败，文件名：" + file.getName() + " 错误信息：" + e.getMessage());
        } finally {
            try {
                if (null != workbook) {
                    workbook.close();
                }
                if (null != inputStream) {
                    inputStream.close();
                }
            } catch (Exception e) {
                log.error("关闭数据流出错！错误信息：" + e.getMessage());
            }
        }
        return null;
    }

    /**
     * 解析Excel数据
     * @param workbook Excel工作簿对象
     * @return 解析结果
     */
    private static List<Student> parseExcel(Workbook workbook) {
        List<Student> resultDataList = new ArrayList<>();

        // 解析sheet
        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            Sheet sheet = workbook.getSheetAt(sheetNum);

            // 校验sheet是否合法
            if (sheet == null) {
                continue;
            }

            // 获取第一行数据
            int firstRowNum = sheet.getFirstRowNum();
            Row firstRow = sheet.getRow(firstRowNum);
            if (null == firstRow) {
                log.error("解析Excel失败，在第一行没有读取到任何数据！");
            }

            // 解析每一行的数据，构造数据对象
            int rowStart = firstRowNum + 1;
            int rowEnd = sheet.getPhysicalNumberOfRows();
            for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
                Row row = sheet.getRow(rowNum);

                if (null == row) {
                    continue;
                }

                Student resultData = convertRowToData(row);
                resultDataList.add(resultData);
            }
        }

        return resultDataList;
    }

    /**
     * 提取每一行中需要的数据，构造成为一个结果数据对象
     *
     * 当该行中有单元格的数据为空或不合法时，忽略该行的数据
     *
     * @param row 行数据
     * @return 解析后的行数据对象，行数据错误时返回null
     */
    private static Student convertRowToData(Row row) {
        Student resultData = new Student();

        Cell cell;
        int cellNum = 0;
        // 获取loginId
        cell = row.getCell(cellNum++);
        String loginId = convertCellValueToString(cell);
        resultData.setLoginId(loginId);
        // 获取名称
        cell = row.getCell(cellNum++);
        String name = convertCellValueToString(cell);
        resultData.setName(name);
        // 获取学号
        cell = row.getCell(cellNum);
        String regNo = convertCellValueToString(cell);
        resultData.setRegNo(regNo);
        return resultData;
    }

    /**
     * 将单元格内容转换为字符串
     * @param cell
     * @return
     */
    private static String convertCellValueToString(Cell cell) {
        if(cell==null){
            return null;
        }
        String returnValue = null;
        switch (cell.getCellTypeEnum()) {
            case NUMERIC:   //数字
                Double doubleValue = cell.getNumericCellValue();

                // 格式化科学计数法，取一位整数
                DecimalFormat df = new DecimalFormat("0");
                returnValue = df.format(doubleValue);
                break;
            case STRING:    //字符串
                returnValue = cell.getStringCellValue();
                break;
            case BOOLEAN:   //布尔
                Boolean booleanValue = cell.getBooleanCellValue();
                returnValue = booleanValue.toString();
                break;
            case BLANK:     // 空值
                break;
            case FORMULA:   // 公式
                returnValue = cell.getCellFormula();
                break;
            case ERROR:     // 故障
                break;
            default:
                break;
        }
        return returnValue;
    }

    /**
     * 格式化单元格内容
     *
     * @param cell
     * @return
     */
    private static Object getCellValue(Cell cell) {
        Object value;
        switch (cell.getCellTypeEnum()) {
            case STRING:  //字符串类型
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                value = cell.getNumericCellValue();
                break;
            case BLANK:
                value = null;
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            default:
                value = "data error";
        }
        return value;
    }


    public static void main(String[] args) throws Exception {
        Class.forName("com.mysql.jdbc.Driver");
        String url = "jdbc:mysql://192.168.20.52:33306/zkyjy_learning_space?useUnicode=true&characterEncoding=UTF-8";
        String user = "zkyjy";
        String password = "9mOx4ulyjW";
        Connection conn = DriverManager.getConnection(url, user, password);
        Statement stat = conn.createStatement();
        String sql = "select * from sso_user where LOGIN_ID in ({0})";
        StringBuilder stringBuilder = new StringBuilder();
        List<Student> studentList = readExcel(new File("C:\\Users\\50381\\Desktop\\students.xls"));
        int i = 1;
        Iterator<Student> iterator = studentList.iterator();
        List<String> result = new ArrayList<>();
        ResultSet resultSet = null;
        while (iterator.hasNext()) {
            stringBuilder.append("'").append(iterator.next().getLoginId()).append("',");

            if (i % 100 == 0) {
                resultSet = stat.executeQuery(MessageFormat.format(sql,stringBuilder.substring(0,stringBuilder.length() - 1)));
                while (resultSet.next()) {
                    result.add(resultSet.getString("ID") + resultSet.getString("LOGIN_ID"));
                }
                stringBuilder.delete(0,stringBuilder.length());
            }

            i++;
        }

        resultSet = stat.executeQuery(MessageFormat.format(sql,stringBuilder.substring(0,stringBuilder.length() - 1)));
        while (resultSet.next()) {
            result.add(resultSet.getString("ID") + resultSet.getString("LOGIN_ID"));
        }

        fileWrite("C:\\Users\\50381\\Desktop\\abc.txt",result,true);
        resultSet.close();
        stat.close();
        conn.close();
    }

    /**
     * 写入文件
     * @param path
     * @param contents
     * @param append
     * @return
     */
    public static File fileWrite(String path, List<String> contents, boolean append) {
        File file = new File(path);

        createFile(file);

        FileWriter fileWritter = null;
        try {
            fileWritter = new FileWriter(file,append);
            for (String content : contents) {
                fileWritter.write(content + "\r\n");
                fileWritter.flush();
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fileWritter != null) {
                try {
                    fileWritter.flush();
                    fileWritter.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        return file;
    }

    /**
     * 如果文件不存在，创建
     *
     * @param files
     */
    public static void createFile(File... files) {
        for (File file : files) {
            if (!file.exists()) {
                file.getParentFile().mkdirs();
                try {
                    file.createNewFile();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
