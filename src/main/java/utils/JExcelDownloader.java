package utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import resource.ExcelData;
import resource.ExcelDataFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;

public class JExcelDownloader {
    private int ROW_INDEX = 0;

    private Workbook workbook;
    private Sheet sheet;

    public JExcelDownloader() {
        this.workbook = new XSSFWorkbook();
        this.sheet = this.workbook.createSheet("Sheet");
    }

    public void excelDownload(List<?> data, Class<?> clazzType, HttpServletResponse response, String fileName, boolean useSeq) throws Exception {
        ExcelDataFactory excelDataFactory = new ExcelDataFactory();
        ExcelData excelData = excelDataFactory.getExcelData(clazzType);

        setHttpHeader(response, fileName);

        renderHeader(excelData, clazzType, useSeq);
        renderBody(excelData, data, clazzType, useSeq);

        write(response.getOutputStream());
    }

    private void setHttpHeader(HttpServletResponse response, String fileName) {
        response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ".xlsx");
        response.setContentType("application/download;charset=utf-8");
        response.setHeader("Content-Transfer-Encoding", "binary");
    }

    private void renderHeader(ExcelData excelData, Class<?> clazzType, boolean useSeq) throws Exception {
        Row row = this.sheet.createRow(ROW_INDEX++);

        int col = 0;
        for(Field field : clazzType.getDeclaredFields()) {
            if(field.isAnnotationPresent(annotation.Cell.class)) {
                if (col == 0 && useSeq) {
                    setCellValue("No", row, col++);
                }
                String headerName = excelData.getHeaderMap().get(field.getName());
                setCellValue(headerName, row, col++);
            }
        }
    }

    private void renderBody(ExcelData excelData, List<?> data, Class<?> clazzType, boolean useSeq) throws Exception {
        for(Object object : data) {
            Row row = this.sheet.createRow(ROW_INDEX++);

            int col = 0;
            for(String fieldName : excelData.getFieldName()) {
                if(col==0 && useSeq) {
                    setCellValue(Integer.toString(ROW_INDEX-1), row, col++);
                }
                Field field = getField(clazzType, fieldName);
                field.setAccessible(true);
                Object cellValue = field.get(object);

                setCellValue((String) cellValue, row, col++);
            }
        }
    }

    private void setCellValue(String value, Row row, int col) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value);
    }

    private void write(OutputStream outputStream) throws IOException {
        this.workbook.write(outputStream);
        outputStream.close();
    }

    private Field getField(Class<?> clazz, String fieldName) {
        for(Field field : clazz.getDeclaredFields()) {
            if(field.getName().equals(fieldName)) {
                return field;
            }
        }
        return null;
    }
}
