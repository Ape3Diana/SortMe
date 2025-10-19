package Model;

import java.util.*;

public class ExcelRow {
    private Map<String, Object> fields=new LinkedHashMap<>();

    public Object get(String fieldName){
        return fields.get(fieldName);
    }

    public void set(String fieldName, Object value){
        fields.put(fieldName, value);
    }

    public List<String> getFieldNames() {
        return new ArrayList<>(fields.keySet());
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder("[ ");
        for (Map.Entry<String, Object> entry : fields.entrySet()) {
            sb.append(entry.getKey())
                    .append(" = ")
                    .append(entry.getValue())
                    .append("; ");
        }
        if (!fields.isEmpty()) {
            sb.setLength(sb.length() - 2);
        }
        sb.append(" ]");
        return sb.toString();
    }
}
