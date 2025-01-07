package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.sql.*;
import java.util.*;
public class Helper {
    // MySQL数据库连接配置
    private static final String DB_URL = "jdbc:mysql://localhost:3306/datatest";
    private static final String USER = "root";
    private static final String PASS = "123456";

    /**
     * 把excel表格数据存到数据库中
     *
     * @param filePath 表格路径
     * @throws Exception
     */
    public static void exlToSql(String filePath) throws Exception {

        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);

        Sheet sheet = workbook.getSheetAt(0);  // 获取第一个工作表
        Iterator<Row> rowIterator = sheet.iterator();

        // 获取表头（第一行）
        Row headerRow = sheet.getRow(0);
        int colNum = headerRow.getPhysicalNumberOfCells();
        List<String> headers = new ArrayList<>();
        for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
            headers.add(headerRow.getCell(i).getStringCellValue()); // 获取列名
        }

        Connection conn = DriverManager.getConnection(DB_URL, USER, PASS);
        conn.setAutoCommit(false); // 开始事务

        String insertSQL = "INSERT INTO datatable (名称, 车型, 所属产品, PCC分类, `describe`, pic) VALUES (?, ?, ?, ?, ?, ?)";
        PreparedStatement stmt = conn.prepareStatement(insertSQL);

        // 跳过第一行（表头），开始处理数据
        rowIterator.next();

        // 拿到所有图片的信息
        Map<Integer, byte[]> m=extractImagesFromExcel(filePath);

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // 读取前四列
            String col1 = getCellValue(row.getCell(0));
            String col2 = getCellValue(row.getCell(1));
            String col3 = getCellValue(row.getCell(2));
            String col4 = getCellValue(row.getCell(3));

            // 读取后续的描述列 k1-kN
            StringBuilder description = new StringBuilder();
            for (int i = 4; i < colNum - 1; i++) {
                String k = headers.get(i);  // 获取列名
                String v = getCellValue(row.getCell(i));  // 获取对应列的值

                if (i > 4) description.append(", ");
                description.append(k).append(":").append(v);
            }


            // 插入数据到数据库
            stmt.setString(1, col1);
            stmt.setString(2, col2);
            stmt.setString(3, col3);
            stmt.setString(4, col4);
            stmt.setString(5, description.toString());
            stmt.setBytes(6,m.get(row.getRowNum()));
            System.out.println("pic row num:"+row.getRowNum() );

            stmt.addBatch();
        }

        stmt.executeBatch(); // 执行批量插入
        conn.commit(); // 提交事务

        stmt.close();
        conn.close();
        workbook.close();
        fis.close();
    }

    /**
     * 获取单元格的值
     *
     * @param cell 单元格
     * @return
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // 如果是日期格式
                    return cell.getDateCellValue().toString();
                } else {
                    // 如果是数字类型，判断是否为整数
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == (int) numericValue) {
                        return String.valueOf((int) numericValue);  // 转换为整数字符串
                    } else {
                        return String.valueOf(numericValue);  // 保留小数
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    /**
     * 从 Excel 文件中提取图片，并返回一个 Map，包含图片所在行号与图片数据的映射。
     *
     * @param filePath Excel 文件路径
     * @return Map<Integer, byte[]> key 为行号，value 为图片的二进制数据
     * @throws Exception 异常
     */
    public static Map<Integer, byte[]> extractImagesFromExcel(String filePath) throws Exception {
        Map<Integer, byte[]> imageMap = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);

            // 获取绘图对象
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            if (drawing == null) {
                return imageMap;
            }

            // 遍历绘图对象中的图片
            for (XSSFShape shape : drawing.getShapes()) {
                if (shape instanceof XSSFPicture) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    XSSFPictureData pictureData = picture.getPictureData();

                    // 获取图片的锚点（确定图片所在的单元格位置）
                    ClientAnchor anchor = picture.getClientAnchor();
                    int rowNum = anchor.getRow1();  // 图片所在的行号

                    // 存储图片数据到 Map
                    imageMap.put(rowNum, pictureData.getData());
                }
            }
        }
        return imageMap;
    }


    /**
     * 从数据库中获取图片的二进制数据 测试使用
     *
     * @param name
     * @return
     */
    public static byte[] getImageFromDatabase(String name) {
        byte[] imageBytes = null;

        String query = "SELECT pic FROM datatable WHERE 名称 = ?";  // 假设主键字段是 id

        try (Connection conn = DriverManager.getConnection(DB_URL, USER, PASS);
             PreparedStatement stmt = conn.prepareStatement(query)) {

            stmt.setString(1, name);

            // 执行查询
            ResultSet rs = stmt.executeQuery();

            if (rs.next()) {
                // 获取 BLOB 数据
                imageBytes = rs.getBytes("pic");  // 获取 pic 字段的二进制数据
            }

        } catch (SQLException e) {
            e.printStackTrace();
        }

        return imageBytes;
    }

    /**
     * 保存图片（默认图片后缀为png） 测试使用
     *
     * @param imageData 图片的二进制数据
     */
    public static void saveImage(byte[] imageData) {
        try (FileOutputStream fos = new FileOutputStream("output_image.png")) {
            fos.write(imageData);
            System.out.println("图片已保存为 output_image.png");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
