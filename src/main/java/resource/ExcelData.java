package resource;

import java.util.List;
import java.util.Map;

public class ExcelData {
    private List<String> fieldName;
    private Map<String, String> headerMap;

    public ExcelData(List<String> fieldName, Map<String, String> headerMap) {
        this.fieldName = fieldName;
        this.headerMap = headerMap;
    }

    public List<String> getFieldName() {
        return fieldName;
    }

    public Map<String, String> getHeaderMap() {
        return headerMap;
    }
}
