package cn.gc.tools;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;


/**
 * <p/>
 * <li>Description: 使用Poi实现Excel导出工具类</li>
 * <li>@author: Guicheng.Zhou</li>
 * <li>Date: 2017/10/16 16:01</li>
 * <li>使用规则:1、创建PoiExportExcel对象;2、创建工作簿对象(createWorkbook);3、写入数据(writeToExcel);4、关闭(close)</li>
 */
public class PoiExportExcel {
    /**
     * <li>response :http请求的响应对象 </li>
     */
    private HttpServletResponse response;

    /**
     * <li>response :工作簿对象 </li>
     */
    private Workbook workbook;

    /**
     * <li>cellStyle :单元格样式 </li>
     */
    private CellStyle cellStyle;

    /**
     * <li>rowOffset :当前操作行数偏移量 </li>
     */
    private int rowOffset = 0;

    /**
     * <li>sheetNum :sheet数 </li>
     */
    private int sheetNum = 1;

    /**
     * <li>sheet :sheet对象 </li>
     */
    private Sheet sheet;

    /**
     * <li>columnNum :定义所需列数 </li>
     */
    private int columnNum;

    /**
     * <li>flushNum :数据在内存缓存数，超过即写入磁盘，默认1000 </li>
     */
    private int flushNum = 1000;

    /**
     * <li>filePath :文件路径 </li>
     */
    private String fileDir;

    /**
     * <li>fileName :文件名 </li>
     */
    private String fileName;

    /**
     * <li>sheetName :sheet名 </li>
     */
    private String sheetName;

    /**
     * <li>out :输出流对象 </li>
     */
    private OutputStream out = null;

    /**
     * <li> PoiExportExcel 的构造函数. </li>
     *
     * @param response  the response
     * @param fileName  the file name
     * @param sheetName the sheet name
     * @throws IOException the io exception
     */
    public PoiExportExcel(HttpServletResponse response, String fileName, String sheetName) throws IOException {
        this(response, null, fileName, sheetName);
    }

    /**
     * <li> PoiExportExcel 的构造函数. </li>
     *
     * @param response  the response
     * @param fileDir   文件路径,如：c:/aaa(不带斜杠)
     * @param fileName  the file name
     * @param sheetName the sheet name
     * @throws IOException the io exception
     */
    public PoiExportExcel(HttpServletResponse response, String fileDir, String fileName, String sheetName)
            throws IOException {
        this.response = response;
        this.fileDir = fileDir;
        this.fileName = String.format("%s_%d.xls", fileName, System.currentTimeMillis());
        this.sheetName = sheetName;
        if (fileDir != null) { //如果有文件路径
            out = new FileOutputStream(String.format("%s/%s", fileDir, fileName));
        } else { //直接写到输出流中
            out = response.getOutputStream();
            response.setCharacterEncoding("UTF-8");
            response.setContentType("text/html");
            response.setHeader("Content-disposition",
                    "attachment; filename=" + new String(this.fileName.getBytes("utf-8"), "ISO8859-1"));
        }
    }

    /**
     * <li>Description: 设置数据在内存缓存数，超过即写入磁盘，默认1000 </li>
     *
     * @param flushNum 内存缓存数
     */
    public void setFlushNum(int flushNum) {
        this.flushNum = flushNum;
    }

    /**
     * <li>Description: 创建工作簿 </li>
     *
     * @param tableHeads 报表表头
     */
    public void createWorkbook(Object tableHeads[]) {
        createWorkbook(null, tableHeads);
    }

    /**
     * <li>Description: 创建工作簿 </li>
     *
     * @param tableName  报表名
     * @param tableHeads 报表表头
     */
    public void createWorkbook(String tableName, Object tableHeads[]) {
        workbook = new SXSSFWorkbook(this.flushNum); //内存缓存数，超过即写入磁盘
        sheet = workbook.createSheet(this.sheetName + "_" + sheetNum++); // 创建工作表
        cellStyle = this.getCellStyle(workbook);
        columnNum = tableHeads.length;
        if (!StringUtils.isEmpty(tableName)) { // 设置第一行一些样式
            CellRangeAddress region = new CellRangeAddress(0, 0, 0, columnNum - 1);
            sheet.addMergedRegion(region); //合并
            Row rowm = sheet.createRow(rowOffset++); // 产生表格标题行
            Cell cellTiltle = rowm.createCell(0);
            cellTiltle.setCellStyle(cellStyle);
            cellTiltle.setCellValue(tableName);
        }
        Row rowRowName = sheet.createRow(rowOffset++); // 创建表头
        rowRowName.setRowStyle(cellStyle);
        for (int i = 0; i < tableHeads.length; i++)
            rowRowName.createCell(i); //创建单元格
        Iterator<Cell> it = rowRowName.cellIterator();
        Arrays.asList(tableHeads).forEach(head -> {
            Cell cell = it.next();
            cell.setCellStyle(cellStyle);
            cell.setCellValue(String.valueOf(head));
        }); // 将列头设置到sheet的单元格中
    }

    /**
     * <li>Description: 写数据到Excel </li>
     *
     * @param tableDatas 数据
     * @throws NullPointerException the null pointer exception
     */
    public void writeToExcel(List<Object[]> tableDatas) throws NullPointerException {
        if (workbook == null)
            throw new NullPointerException("workbook is null");
        Object[] data;
        for (int i = 0; i < tableDatas.size(); i++) { // 将数据设置到sheet对应的单元格中
            Row row;
            try {
                row = sheet.createRow(rowOffset++); // 创建数据行
            } catch (IllegalArgumentException e) { //一个sheet最多1048576行,新建sheet
                sheet = workbook.createSheet(this.sheetName + "_" + sheetNum++); // 创建工作表
                rowOffset = 0;
                writeToExcel(tableDatas.subList(i, tableDatas.size()));
                break;
            }
            data = tableDatas.get(i);
            for (int j = 0; j < data.length; j++)
                row.createCell(j); //创建单元格
            Iterator<Cell> it1 = row.cellIterator();
            Arrays.asList(data).forEach(obj -> {
                Cell cell = it1.next();
                cell.setCellStyle(cellStyle);
                if (obj != null) {
                    if (obj instanceof Integer) {
                        cell.setCellValue(((Integer) obj).longValue());
                    } else if (obj instanceof Byte) {
                        cell.setCellValue(((Byte) obj).longValue());
                    } else if (obj instanceof String) {
                        cell.setCellValue((String) obj);
                    } else if (obj instanceof Double) {
                        cell.setCellValue(((Double) obj).longValue());
                    }
                }
            });
        }
    }

    /**
     * <li>Description: 关闭输出流对象 </li>
     *
     * @throws IOException          the io exception
     * @throws NullPointerException the null pointer exception
     */
    public void close() throws IOException, NullPointerException {
        try {
            if (workbook == null)
                throw new NullPointerException("workbook is null");
            workbook.write(out);
        } catch (IOException e) {
            throw e;
        } finally {
            try {
                if (out != null)
                    out.close();
            } catch (IOException e) {
                throw e;
            }
        }
    }

    /**
     * <li>Description: 获得单元样式 </li>
     *
     * @param workbook Workbook对象
     * @return 单元样式
     */
    private CellStyle getCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER); // 居中
        style.setWrapText(true); // 自动换行
        style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        return style;
    }

    /**
     * <li>Description: 测试 </li>
     *
     * @param args TODO
     */
    public static void main(String[] args) {
        long start = System.currentTimeMillis();
        PoiExportExcel poi = null;
        try {
            poi = new PoiExportExcel(null, "F:", "text.xls", "sheet");
            Object[] objArr1 = { 1, 2, 3, 4, 5 };
            poi.createWorkbook(null, objArr1);
            List<Object[]> list = new ArrayList<>();
            for (int i = 1; i <= 2000000; i++) {
                Object[] objArr2 = { i + 1, i + 1, i + 1, i + 1, i + 1 };
                list.add(objArr2);
                if (i % 100000 == 0) {
                    long s = System.currentTimeMillis();
                    poi.writeToExcel(list);
                    System.out.println("写完100000 耗时：" + (System.currentTimeMillis() - s) + " ms");
                    list.clear();
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (poi != null) {
                    poi.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                System.out.println(System.currentTimeMillis() - start);
            }
        }
    }
}
