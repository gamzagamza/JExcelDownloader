package resource;

import annotation.Cell;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelDataFactory {

    public ExcelData getExcelData(Class<?> clazz) {
        List<String> fieldName = new ArrayList<>();
        Map<String, String> headerMap = new HashMap<>();

        for(Field field : clazz.getDeclaredFields()) {
            if(field.isAnnotationPresent(Cell.class)) {
                Cell cell = field.getAnnotation(Cell.class);
                headerMap.put(field.getName(), cell.headerName());
                fieldName.add(field.getName());
            }
        }

        return new ExcelData(fieldName, headerMap);
    }
}
