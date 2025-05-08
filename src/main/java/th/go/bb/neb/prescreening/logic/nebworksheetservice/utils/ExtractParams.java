package th.go.bb.neb.prescreening.logic.nebworksheetservice.utils;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


public class ExtractParams {

    /**
     * แยกค่าพารามิเตอร์ที่เป็น String จากคำสั่ง
     * เช่น FILE("path/to/file.xlsx") จะได้ "path/to/file.xlsx"
    */
    public static String extractStringParam(String param, String command) {
        String pattern = command + "\\(\"([^\"]+)\"\\)";
        Pattern r = Pattern.compile(pattern);
        Matcher m = r.matcher(param);
        if (m.find()) {
            return m.group(1);
        }
        return "";
    }

    /**
     * แยกค่าพารามิเตอร์หลายค่าจากคำสั่ง
     * เช่น FIX("KEY", "VALUE") จะได้ ["KEY", "VALUE"]
     */
    public static String[] extractMultipleParams(String param, String command) {
        // ตัดส่วนหัวและท้ายของคำสั่งออก
        String content = param.substring(command.length() + 1, param.length() - 1);

        List<String> params = new ArrayList<>();
        StringBuilder currentParam = new StringBuilder();
        boolean inQuotes = false;
        int bracketCount = 0;

        for (int i = 0; i < content.length(); i++) {
            char c = content.charAt(i);

            if (c == '"') {
                inQuotes = !inQuotes;
                currentParam.append(c);
            } else if (c == '[') {
                bracketCount++;
                currentParam.append(c);
            } else if (c == ']') {
                bracketCount--;
                currentParam.append(c);
            } else if (c == ',' && !inQuotes && bracketCount == 0) {
                // พบตัวคั่นพารามิเตอร์
                params.add(currentParam.toString().trim());
                currentParam = new StringBuilder();
            } else {
                currentParam.append(c);
            }
        }

        // เพิ่มพารามิเตอร์สุดท้าย
        if (currentParam.length() > 0) {
            params.add(currentParam.toString().trim());
        }

        // แยกเครื่องหมายคำพูดออกจากค่า
        for (int i = 0; i < params.size(); i++) {
            String p = params.get(i);
            if (p.startsWith("\"") && p.endsWith("\"")) {
                params.set(i, p.substring(1, p.length() - 1));
            }
        }

        return params.toArray(new String[0]);
    }

    /**
     * แยกค่าพารามิเตอร์ที่เป็น Array จากคำสั่ง
     * เช่น ["VALUE1", "VALUE2"] จะได้ ["VALUE1", "VALUE2"]
     */
    public static String[] extractArrayParam(String param) {
        if (!param.startsWith("[") || !param.endsWith("]")) {
            return new String[0];
        }

        // ตัดวงเล็บเปิดและปิดออก
        String content = param.substring(1, param.length() - 1);

        List<String> params = new ArrayList<>();
        StringBuilder currentParam = new StringBuilder();
        boolean inQuotes = false;

        for (int i = 0; i < content.length(); i++) {
            char c = content.charAt(i);

            if (c == '"') {
                inQuotes = !inQuotes;
            } else if (c == ',' && !inQuotes) {
                // พบตัวคั่นพารามิเตอร์
                String value = currentParam.toString().trim();
                if (value.startsWith("\"") && value.endsWith("\"")) {
                    value = value.substring(1, value.length() - 1);
                }
                params.add(value);
                currentParam = new StringBuilder();
                continue;
            }

            currentParam.append(c);
        }

        // เพิ่มพารามิเตอร์สุดท้าย
        if (currentParam.length() > 0) {
            String value = currentParam.toString().trim();
            if (value.startsWith("\"") && value.endsWith("\"")) {
                value = value.substring(1, value.length() - 1);
            }
            params.add(value);
        }

        return params.toArray(new String[0]);
    }

    /**
     * แยกค่าพารามิเตอร์สำหรับ CHKSHEETDUPLICATE
     * เช่น [("SHEET1", RANGE1, ["EXCEPT1"]),("SHEET2", RANGE2, [])]
     * จะได้ List ของ Map ที่มีข้อมูล sheetName, rangeStr, และ exceptValues
     */
    public static List<Map<String, Object>> extractSheetConfigs(String content) {
        List<Map<String, Object>> sheetConfigs = new ArrayList<>();

        // ตัดวงเล็บเปิดและปิดของ array ออก
        content = content.substring(1, content.length() - 1).trim();

        // แยกแต่ละ tuple
        List<String> tuples = new ArrayList<>();
        StringBuilder currentTuple = new StringBuilder();
        boolean inQuotes = false;
        int parenthesisCount = 0;
        int bracketCount = 0;

        for (int i = 0; i < content.length(); i++) {
            char c = content.charAt(i);

            if (c == '"') {
                inQuotes = !inQuotes;
                currentTuple.append(c);
            } else if (c == '(') {
                parenthesisCount++;
                currentTuple.append(c);
            } else if (c == ')') {
                parenthesisCount--;
                currentTuple.append(c);
            } else if (c == '[') {
                bracketCount++;
                currentTuple.append(c);
            } else if (c == ']') {
                bracketCount--;
                currentTuple.append(c);
            } else if (c == ',' && !inQuotes && parenthesisCount == 0 && bracketCount == 0) {
                // พบตัวคั่น tuple
                tuples.add(currentTuple.toString().trim());
                currentTuple = new StringBuilder();
            } else {
                currentTuple.append(c);
            }
        }

        // เพิ่ม tuple สุดท้าย
        if (currentTuple.length() > 0) {
            tuples.add(currentTuple.toString().trim());
        }

        // แยกข้อมูลในแต่ละ tuple
        for (String tuple : tuples) {
            // ตัดวงเล็บเปิดและปิดของ tuple ออก
            if (tuple.startsWith("(") && tuple.endsWith(")")) {
                tuple = tuple.substring(1, tuple.length() - 1).trim();

                // แยกพารามิเตอร์ในแต่ละ tuple
                List<String> params = new ArrayList<>();
                StringBuilder currentParam = new StringBuilder();
                inQuotes = false;
                bracketCount = 0;

                for (int i = 0; i < tuple.length(); i++) {
                    char c = tuple.charAt(i);

                    if (c == '"') {
                        inQuotes = !inQuotes;
                        currentParam.append(c);
                    } else if (c == '[') {
                        bracketCount++;
                        currentParam.append(c);
                    } else if (c == ']') {
                        bracketCount--;
                        currentParam.append(c);
                    } else if (c == ',' && !inQuotes && bracketCount == 0) {
                        // พบตัวคั่นพารามิเตอร์
                        params.add(currentParam.toString().trim());
                        currentParam = new StringBuilder();
                    } else {
                        currentParam.append(c);
                    }
                }

                // เพิ่มพารามิเตอร์สุดท้าย
                if (currentParam.length() > 0) {
                    params.add(currentParam.toString().trim());
                }

                // สร้าง Map เก็บข้อมูล
                if (params.size() >= 2) {
                    Map<String, Object> sheetConfig = new HashMap<>();

                    // ชื่อชีท
                    String sheetName = params.get(0);
                    if (sheetName.startsWith("\"") && sheetName.endsWith("\"")) {
                        sheetName = sheetName.substring(1, sheetName.length() - 1);
                    }
                    sheetConfig.put("sheetName", sheetName);

                    // ช่วงของคอลัมน์
                    String rangeStr = params.get(1);
                    sheetConfig.put("rangeStr", rangeStr);

                    // ค่ายกเว้น
                    if (params.size() >= 3 && !params.get(2).equals("NULL")) {
                        String exceptParam = params.get(2);
                        Set<Object> exceptValues = new HashSet<>();
                        String[] exceptList = extractArrayParam(exceptParam);
                        for (String except : exceptList) {
                            exceptValues.add(except);
                        }
                        sheetConfig.put("exceptValues", exceptValues);
                    }

                    sheetConfigs.add(sheetConfig);
                }
            }
        }

        return sheetConfigs;
    }

    public static Map<String, Object> extractCrossFileCompare(String content , String type) {
        Map<String, Object> result = new HashMap<>();

        String condition = new String();
        String arrayPart = content;

        if (type.equals("CROSSFILECOMPARE")) {
            if (content.startsWith("CROSSFILECOMPARE(") && content.endsWith(")")) {
                content = content.substring("CROSSFILECOMPARE(".length(), content.length() - 1).trim();
                int commaIndex = content.indexOf(',');
                if (commaIndex == -1) {
                    throw new IllegalArgumentException("ไม่พบ , คั่น condition และ array");
                }

                condition = content.substring(0, commaIndex).trim();
                arrayPart = content.substring(commaIndex + 1).trim();
            } else {
                throw new IllegalArgumentException("รูปแบบไม่ถูกต้อง: " + content);
            }
        }else if (type.equals("CROSSFILEDUPLICATED")) {
            if (content.startsWith("CROSSFILEDUPLICATED(") && content.endsWith(")")) {
                content = content.substring("CROSSFILEDUPLICATED(".length(), content.length() - 1).trim();
            } else {
                throw new IllegalArgumentException("รูปแบบไม่ถูกต้อง: " + content);
            }
        }


        if (arrayPart.startsWith("[") && arrayPart.endsWith("]")) {
            arrayPart = arrayPart.substring(1, arrayPart.length() - 1).trim();
        }

        List<Map<String, Object>> sheetConfigs = new ArrayList<>();
        StringBuilder currentTuple = new StringBuilder();
        boolean inQuotes = false;
        int bracketCount = 0;

        for (int i = 0; i < arrayPart.length(); i++) {
            char c = arrayPart.charAt(i);

            if (c == '"') {
                inQuotes = !inQuotes;
                currentTuple.append(c);
            } else if (c == '[') {
                bracketCount++;
                currentTuple.append(c);
            } else if (c == ']') {
                bracketCount--;
                currentTuple.append(c);
            } else if (c == ',' && !inQuotes && bracketCount == 0) {
                String tuple = currentTuple.toString().trim();
                if (!tuple.isEmpty()) {
                    sheetConfigs.add(parseCrossFileTuple(tuple));
                }
                currentTuple.setLength(0);
            } else {
                currentTuple.append(c);
            }
        }

        if (currentTuple.length() > 0) {
            String tuple = currentTuple.toString().trim();
            if (!tuple.isEmpty()) {
                sheetConfigs.add(parseCrossFileTuple(tuple));
            }
        }

        if (type.equals("CROSSFILECOMPARE")) {
            result.put("condition", condition);
        }
        result.put("sheetConfigs", sheetConfigs);
        return result;
    }

    public static Map<String, Object> parseCrossFileTuple(String tuple) {
        tuple = tuple.trim();
        if (tuple.startsWith("[") && tuple.endsWith("]")) {
            tuple = tuple.substring(1, tuple.length() - 1).trim();
        }

        List<String> parts = new ArrayList<>();
        boolean inQuotes = false;
        StringBuilder current = new StringBuilder();

        for (int i = 0; i < tuple.length(); i++) {
            char c = tuple.charAt(i);
            if (c == '"') {
                inQuotes = !inQuotes;
            } else if (c == ',' && !inQuotes) {
                parts.add(current.toString().replace("\"", "").trim());
                current.setLength(0);
                continue;
            }
            current.append(c);
        }
        if (current.length() > 0) {
            parts.add(current.toString().replace("\"", "").trim());
        }

        if (parts.size() < 3) {
            throw new IllegalArgumentException("tuple ไม่ครบ 3 ค่า: " + tuple);
        }

        Map<String, Object> map = new HashMap<>();
        map.put("fileName", parts.get(0));
        map.put("sheetName", parts.get(1));
        map.put("rangeStr", parts.get(2));
        return map;
    }
} 
