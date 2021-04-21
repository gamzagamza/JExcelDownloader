package resource.data;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.List;
import java.util.Map;

public class ExcelData {
    private List<String> fieldName;
    private Map<String, String> headerMap;
    Map<String, CellStyle> headerStyleMap;
    Map<String, CellStyle> bodyStyleMap;

    public ExcelData(List<String> fieldName, Map<String, String> headerMap,
                     Map<String, CellStyle> headerStyleMap, Map<String, CellStyle> bodyStyleMap) {
        this.fieldName = fieldName;
        this.headerMap = headerMap;
        this.headerStyleMap = headerStyleMap;
        this.bodyStyleMap = bodyStyleMap;
    }

    public List<String> getFieldName() {
        return fieldName;
    }

    public Map<String, String> getHeaderMap() {
        return headerMap;
    }

    public Map<String, CellStyle> getHeaderStyleMap() { return headerStyleMap; }

    public Map<String, CellStyle> getBodyStyleMap() { return bodyStyleMap; }
}
