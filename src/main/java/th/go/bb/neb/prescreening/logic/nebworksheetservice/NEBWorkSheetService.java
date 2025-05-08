package th.go.bb.neb.prescreening.logic.nebworksheetservice;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import th.go.bb.neb.prescreening.logic.nebworksheetservice.utils.ExcelReader;
import th.go.bb.neb.prescreening.logic.nebworksheetservice.utils.ExtractParams;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class NEBWorkSheetService {

    /**
     * อ่านข้อมูลจากไฟล์ Excel และแปลงเป็นรูปแบบ wssResult
     * โดยรวมผลลัพธ์จาก readExcelVariables และการตรวจสอบค่าซ้ำเข้าด้วยกัน
     * โดยนำค่าจาก readExcelVariables ใส่ใน wssResult.data เสมอ
     * ไม่ว่าจะพบค่าซ้ำหรือไม่
     * 
     * @param fileName ชื่อไฟล์ Excel
     * @param params   พารามิเตอร์ในรูปแบบ String[] โดยมีรูปแบบดังนี้:
     *                 params[0] = "SHEET(\"ชื่อชีท\")"
     *                 params[1] = "FIX(\"FISCAL_YEAR\", \"2568\")"
     *                 params[2] = "FIX(\"TEMPLATE_ID\",
     *                 \"21811913-aa9b-48f8-86cb-ed4041b17a99\")"
     *                 params[3] = "COLUMN(\"AMOUNT\", \"Z4\")"
     *                 params[4] = "COUNT(\"QTY\", \"B15:EOF\")"
     *                 params[5] = "CHKDUPLICATE(\"คำนวณเงินเดือน\", B15:EOF,
     *                 [\"ตำแหน่งว่าง\"])"
     *                 params[6] = "CHKDUPLICATE(\"คำนวณบำเหน็จบำนาญ\", C12:EOF,
     *                 NULL)"
     * @return JSON string ในรูปแบบ wssResult
     */
    public static String processExcelToWssResult(String fileName, String[] params) {
        try {
            if (params.length < 1) {
                throw new IllegalArgumentException("ต้องระบุพารามิเตอร์อย่างน้อย 1 ค่า: sheetName");
            }
            File file = new File(fileName);
            Workbook workbook = createWorkbook(file, fileName);

            // กำหนดค่าเริ่มต้น
            String sheetName = "";
            Map<String, Object> variables = new HashMap<>();
            Set<Object> validates = new HashSet<>();

            // แยกพารามิเตอร์แต่ละรายการ
            for (String param : params) {
                if (param.startsWith("SHEET(")) {
                    // รูปแบบ SHEET("ชื่อชีท")
                    sheetName = ExtractParams.extractStringParam(param, "SHEET");
                } else if (param.startsWith("FIX(")) {
                    // รูปแบบ FIX("KEY", "VALUE")
                    String[] parts = ExtractParams.extractMultipleParams(param, "FIX");
                    if (parts.length >= 2) {
                        variables.put(parts[0], "FIX " + parts[1]);
                    }
                } else if (param.startsWith("COLUMN(")) {
                    // รูปแบบ COLUMN("KEY", "CELL_REF", "SHEET_NAME")
                    String[] parts = ExtractParams.extractMultipleParams(param, "COLUMN");
                    if (parts.length >= 2) {
                        Map<String, Object> columnConfig = new HashMap<>();
                        columnConfig.put("type", "COLUMN");
                        columnConfig.put("cellRef", parts[1]);

                        // ตรวจสอบว่ามีการระบุชื่อชีทหรือไม่
                        if (parts.length >= 3) {
                            columnConfig.put("sheetName", parts[2]);
                        }

                        variables.put(parts[0], columnConfig);
                    }
                } else if (param.startsWith("COUNT(")) {
                    // รูปแบบ COUNT("KEY", "RANGE", "SHEET_NAME")
                    String[] parts = ExtractParams.extractMultipleParams(param, "COUNT");
                    if (parts.length >= 2) {
                        Map<String, Object> countConfig = new HashMap<>();
                        countConfig.put("type", "COUNT");
                        countConfig.put("rangeStr", parts[1]);

                        // ตรวจสอบว่ามีการระบุชื่อชีทหรือไม่
                        if (parts.length >= 3) {
                            countConfig.put("sheetName", parts[2]);
                        }

                        variables.put(parts[0], countConfig);
                    }
                } else if (param.startsWith("ROWBY(")) {
                    System.out.println("================[ROWBY]");
                    String[] parts = ExtractParams.extractMultipleParams(param, "ROWBY");

                    // รูปแบบใหม่: ROWBY(searchCondition, columnRefs, sheetName)
                    // ตัวอย่าง: ROWBY(A EQUAL "รวมทั้งสิ้น", [COLUMN("AMOUNT", "C?",
                    // "11.ปัจจัยพื้นฐาน"),COLUMN("AMOUNT", "F?", "11.ปัจจัยพื้นฐาน")],
                    // "11.ปัจจัยพื้นฐาน")
                    if (parts.length >= 3) {
                        System.out.println("================[ROWBY 3]");
                        // แยกเงื่อนไขการค้นหา (searchCondition)
                        String searchConditionStr = parts[0];
                        String[] searchConditionParts = searchConditionStr.split(" ", 3);

                        if (searchConditionParts.length >= 3) {
                            System.out.println("================[ROWBY 33]");
                            String searchColumn = searchConditionParts[0];
                            String searchCondition = searchConditionParts[1];
                            String searchValue = searchConditionParts[2];

                            // ถ้า searchValue อยู่ในเครื่องหมายคำพูด ให้ตัดออก
                            if (searchValue.startsWith("\"") && searchValue.endsWith("\"")) {
                                searchValue = searchValue.substring(1, searchValue.length() - 1);
                            }

                            // สร้าง config สำหรับ ROWBY
                            Map<String, Object> rowByConfig = new HashMap<>();
                            rowByConfig.put("type", "ROWBY");
                            rowByConfig.put("searchSheetName", parts[2]); // ใช้ sheet เดียวกันสำหรับค้นหาและอ่านข้อมูล
                            rowByConfig.put("readSheetName", parts[2]);
                            rowByConfig.put("searchColumn", searchColumn);
                            rowByConfig.put("searchCondition", searchCondition);
                            rowByConfig.put("searchValue", searchValue);
                            rowByConfig.put("columnRefs", parts[1]);
                            System.out.println("=========================parts[1]" + parts[1]);
                            // ใช้ชื่อตัวแปรเป็น "ROWBY" + ลำดับ
                            String varName = "ROWBY";// + variables.size();
                            variables.put(varName, rowByConfig);
                        }
                    }
                } else if (param.startsWith("ROW(")) {
                    // รูปแบบ ROW("KEY", "SHEET_NAME", "ROW_RANGE", "[COL1,COL2,COL3]",
                    // "{COL1:VAR1,COL2:VAR2}")
                    String[] parts = ExtractParams.extractMultipleParams(param, "ROW");
                    if (parts.length >= 4) {
                        Map<String, Object> rowConfig = new HashMap<>();
                        rowConfig.put("type", "ROW");
                        rowConfig.put("sheetName", parts[1]);
                        rowConfig.put("rowRange", parts[2]);
                        rowConfig.put("columns", parts[3]);

                        // ตรวจสอบว่ามีการระบุ mapping หรือไม่
                        if (parts.length >= 5) {
                            rowConfig.put("mapping", parts[4]);
                        }

                        variables.put(parts[0], rowConfig);
                    }
                } else if (param.startsWith("CHKDUPLICATE(")) {
                    // รูปแบบ CHKDUPLICATE("SHEET_NAME", RANGE, ["EXCEPT1", "EXCEPT2"])
                    String[] parts = ExtractParams.extractMultipleParams(param, "CHKDUPLICATE");
                    if (parts.length >= 2) {
                        String targetSheetName = parts[0]; // ชื่อชีทที่ต้องการตรวจสอบ
                        String rangeStr = parts[1]; // ช่วงของคอลัมน์ที่ต้องการตรวจสอบ

                        Map<String, Object> checkDuplicate = new HashMap<>();
                        checkDuplicate.put("sheetName", targetSheetName);
                        checkDuplicate.put("condition", "CHKDUPLICATE " + rangeStr);

                        // ตรวจสอบว่ามีการระบุค่ายกเว้นหรือไม่
                        if (parts.length >= 3 && !parts[2].equals("NULL")) {
                            Set<Object> exceptValues = new HashSet<>();
                            String[] exceptList = ExtractParams.extractArrayParam(parts[2]);
                            for (String except : exceptList) {
                                exceptValues.add(except);
                            }
                            checkDuplicate.put("exceptValues", exceptValues);
                        }

                        validates.add(checkDuplicate);
                    }
                } else if (param.startsWith("CHKSHEETDUPLICATE(")) {
                    // รูปแบบ CHKSHEETDUPLICATE([("SHEET1", RANGE1, ["EXCEPT1"]),("SHEET2", RANGE2,
                    // [])])
                    String content = param.substring("CHKSHEETDUPLICATE(".length(), param.length() - 1).trim();

                    // ตรวจสอบว่าเป็นรูปแบบ array หรือไม่
                    if (content.startsWith("[") && content.endsWith("]")) {
                        // แยกแต่ละ tuple ในรูปแบบ ("SHEET", RANGE, [EXCEPTS])
                        List<Map<String, Object>> sheetConfigs = ExtractParams.extractSheetConfigs(content);

                        // สร้าง validate สำหรับการตรวจสอบข้ามชีท
                        Map<String, Object> checkSheetDuplicate = new HashMap<>();
                        checkSheetDuplicate.put("condition", "CHKSHEETDUPLICATE");
                        checkSheetDuplicate.put("sheetConfigs", sheetConfigs);

                        validates.add(checkSheetDuplicate);
                    }
                }
            }

            // เพิ่ม validates เข้าไปใน variables
            if (!validates.isEmpty()) {
                variables.put("VALIDATES", validates);
            }

            // ตรวจสอบและดึงข้อมูลการตรวจสอบค่าซ้ำจาก variables
            String rangeStr = null;
            Set<Object> exceptValues = null;

            // ดึงข้อมูลการตรวจสอบค่าซ้ำจาก variables
            if (variables.containsKey("VALIDATES")) {
                Object validateObj = variables.get("VALIDATES");
                if (validateObj instanceof Set) {
                    Set<?> validateSet = (Set<?>) validateObj;
                    for (Object validate : validateSet) {
                        if (validate instanceof Map) {
                            Map<?, ?> condition = (Map<?, ?>) validate;
                            if (condition.containsKey("condition") && condition.get("condition") instanceof String) {
                                String conditionStr = (String) condition.get("condition");
                                if (conditionStr.startsWith("CHKDUPLICATE ")) {
                                    rangeStr = conditionStr.substring("CHKDUPLICATE ".length());
                                }
                            }
                            if (condition.containsKey("exceptValues") && condition.get("exceptValues") instanceof Set) {
                                exceptValues = (Set<Object>) condition.get("exceptValues");
                            }
                        }
                    }
                }
            }

            // ถ้าไม่มีการระบุช่วงเซลล์สำหรับตรวจสอบค่าซ้ำ ให้ใช้ค่าเริ่มต้น
            if (rangeStr == null) {
                rangeStr = "B15:EOF";
            }

            // อ่านข้อมูลจาก Excel
            Map<String, Object> data = ExcelReader.readExcelVariables(workbook, fileName, sheetName, variables);

            // ตรวจสอบค่าซ้ำในแต่ละชีทที่กำหนด
            Map<Object, List<String>> duplicates = new HashMap<>();
            List<String> crossSheetDuplicates = new ArrayList<>();
            List<String> sheetNotFoundErrors = new ArrayList<>();

            if (variables.containsKey("VALIDATES")) {
                Object validateObj = variables.get("VALIDATES");
                if (validateObj instanceof Set) {
                    Set<?> validateSet = (Set<?>) validateObj;
                    for (Object validate : validateSet) {
                        if (validate instanceof Map) {
                            Map<?, ?> condition = (Map<?, ?>) validate;

                            // ตรวจสอบค่าซ้ำในชีทเดียว
                            if (condition.containsKey("sheetName") && condition.containsKey("condition")) {
                                String targetSheetName = (String) condition.get("sheetName");
                                String conditionStr = (String) condition.get("condition");
                                System.out.println("============================" + conditionStr);
                                if (conditionStr.startsWith("CHKDUPLICATE ")) {
                                    String targetRangeStr = conditionStr.substring("CHKDUPLICATE ".length());
                                    Set<Object> targetExceptValues = null;
                                    if (condition.containsKey("exceptValues")) {
                                        targetExceptValues = (Set<Object>) condition.get("exceptValues");
                                    }

                                    try {
                                        // ตรวจสอบค่าซ้ำในชีทที่กำหนด
                                        Map<Object, List<String>> sheetDuplicates = ExcelReader
                                                .checkDuplicateValuesInRange(workbook, targetSheetName, targetRangeStr,
                                                        targetExceptValues);

                                        // รวมผลลัพธ์
                                        duplicates.putAll(sheetDuplicates);
                                    } catch (IllegalArgumentException e) {
                                        // กรณีไม่พบชีท
                                        if (e.getMessage().contains("ไม่พบชีท")) {
                                            sheetNotFoundErrors.add("ไม่พบชีท '" + targetSheetName + "' ในไฟล์");
                                        } else {
                                            throw e; // ส่งต่อข้อผิดพลาดอื่นๆ
                                        }
                                    }
                                }
                            }
                            // ตรวจสอบค่าซ้ำระหว่างชีท
                            else if (condition.containsKey("condition") && "CHKSHEETDUPLICATE".equals(condition.get("condition"))) {
                                if (condition.containsKey("sheetConfigs")) {
                                    List<Map<String, Object>> sheetConfigs = (List<Map<String, Object>>) condition.get("sheetConfigs");

                                    // เก็บข้อมูลจากแต่ละชีท
                                    Map<String, List<String>> allSheetValues = new LinkedHashMap<>();

                                    // อ่านข้อมูลจากแต่ละชีท
                                    for (Map<String, Object> sheetConfig : sheetConfigs) {
                                        String targetSheetName = (String) sheetConfig.get("sheetName");
                                        String targetRangeStr = (String) sheetConfig.get("rangeStr");
                                        Set<Object> targetExceptValues = null;
                                        if (sheetConfig.containsKey("exceptValues")) {
                                            targetExceptValues = (Set<Object>) sheetConfig.get("exceptValues");
                                        }

                                        try {
                                            // อ่านค่าจากชีท
                                            List<String> sheetValues = ExcelReader.getAllValuesInRange(
                                                    workbook, targetSheetName, targetRangeStr, targetExceptValues);
                                            // เก็บข้อมูล
                                            allSheetValues.put(targetSheetName, sheetValues);
                                        } catch (IllegalArgumentException e) {
                                            // กรณีไม่พบชีท
                                            if (e.getMessage().contains("ไม่พบชีท")) {
                                                sheetNotFoundErrors.add("ไม่พบชีท '" + targetSheetName + "' ในไฟล์");
                                            } else {
                                                throw e; // ส่งต่อข้อผิดพลาดอื่นๆ
                                            }
                                        }
                                    }

                                    // ตรวจสอบค่าซ้ำระหว่างชีท
                                    if (allSheetValues.size() >= 2) {
                                        // System.out.println(allSheetValues);
                                        List<String> duplicateValues = new ArrayList<>();

                                        Iterator<List<String>> it = allSheetValues.values().iterator();
                                        if (it.hasNext()) {
                                            List<String> list1 = it.next();
                                            if (it.hasNext()) {
                                                List<String> list2 = it.next();
                                                Set<String> set1 = new HashSet<>(list1);
                                                for (String value : list2) {
                                                    if (set1.contains(value)) {
                                                        duplicateValues.add(value);
                                                    }
                                                }
                                            }
                                        }

                                        System.out.println("ค่าที่ซ้ำกันคือ: " + duplicateValues);
                                        crossSheetDuplicates.addAll(duplicateValues);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // สร้าง JSON string ในรูปแบบ wssResult
            JSONObject wssResult = new JSONObject();
            JSONObject wssResultObj = new JSONObject();

            // ส่วนของ data - ใส่ข้อมูลจาก readExcelVariables เสมอ
            wssResultObj.put("data", new JSONObject(data));

            // ตรวจสอบว่ามีข้อผิดพลาดเกี่ยวกับชีทหรือไม่
            if (!sheetNotFoundErrors.isEmpty()) {
                // กรณีไม่พบชีท
                wssResultObj.put("status", "fail");

                JSONObject errorMessage = new JSONObject();
                errorMessage.put("errorCode", "SHEET_NOT_FOUND");
                errorMessage.put("errorMessage", String.join(", ", sheetNotFoundErrors));

                wssResultObj.put("errorMessage", errorMessage);
            } else {
                // ส่วนของ status
                boolean isSuccess = duplicates.isEmpty() && crossSheetDuplicates.isEmpty();
                wssResultObj.put("status", isSuccess ? "success" : "fail");

                // ส่วนของ errorMessage
                if (isSuccess) {
                    // ไม่พบค่าซ้ำ
                    wssResultObj.put("errorMessage", new JSONObject());
                } else {
                    // พบค่าซ้ำ - สร้าง errorMessage ที่มีข้อมูลเกี่ยวกับค่าซ้ำ
                    JSONObject errorMessage = new JSONObject();

                    if (!duplicates.isEmpty() && !crossSheetDuplicates.isEmpty()) {
                        // พบค่าซ้ำทั้งในชีทเดียวและระหว่างชีท
                        errorMessage.put("errorCode", "DUPLICATE_AND_CROSS_SHEET_DUPLICATE");
                        errorMessage.put("errorMessage", "พบค่าซ้ำในข้อมูลทั้งในชีทเดียวและระหว่างชีท");
                    } else if (!duplicates.isEmpty()) {
                        // พบค่าซ้ำในชีทเดียว
                        errorMessage.put("errorCode", "DUPLICATE");
                        errorMessage.put("errorMessage", "พบค่าซ้ำในข้อมูล");
                    } else {
                        // พบค่าซ้ำระหว่างชีท
                        errorMessage.put("errorCode", "CROSS_SHEET_DUPLICATE");
                        errorMessage.put("errorMessage", "พบค่าซ้ำระหว่างชีท");
                    }

                    // แปลง Map<Object, List<String>> เป็น JSONObject สำหรับค่าซ้ำในชีทเดียว
                    if (!duplicates.isEmpty()) {
                        JSONObject duplicateValues = new JSONObject();
                        for (Map.Entry<Object, List<String>> entry : duplicates.entrySet()) {
                            String key = entry.getKey() != null ? entry.getKey().toString() : "null";
                            duplicateValues.put(key, entry.getValue());
                        }
                        errorMessage.put("duplicateValues", duplicateValues);
                    }

                    // แปลง Map<Object, List<String>> เป็น JSONObject สำหรับค่าซ้ำระหว่างชีท
                    if (!crossSheetDuplicates.isEmpty()) {
                        errorMessage.put("crossSheetDuplicateValues", new JSONObject(crossSheetDuplicates));
                    }

                    wssResultObj.put("errorMessage", errorMessage);
                }
            }

            wssResult.put("wssResult", wssResultObj);
            return wssResult.toString();
        } catch (IOException e) {
            // กรณีเกิดข้อผิดพลาด
            JSONObject wssResult = new JSONObject();
            JSONObject wssResultObj = new JSONObject();

            wssResultObj.put("data", new JSONObject());
            wssResultObj.put("status", "fail");

            JSONObject errorMessage = new JSONObject();
            errorMessage.put("errorCode", "E001");
            errorMessage.put("errorMessage", e.getMessage());

            wssResultObj.put("errorMessage", errorMessage);
            wssResult.put("wssResult", wssResultObj);

            return wssResult.toString();
        }
    }

    public static String processMutiFileExcelToWssResult(List<String> fileNameList, String[] params) {
        try {
            if (params.length < 1) {
                throw new IllegalArgumentException("ต้องระบุพารามิเตอร์อย่างน้อย 1 ค่า: sheetName");
            }
            if (fileNameList.size() != 2) {
                throw new IllegalArgumentException("ต้องระบุ 2 ไฟล์สำหรับเปรียบเทียบ");
            }

            List<Workbook> workbooks = new ArrayList<>();
            for (String fileName : fileNameList) {
                File file = new File(fileName);
                Workbook workbook = createWorkbook(file, fileName);
                workbooks.add(workbook);
            }

            // กำหนดค่าเริ่มต้น
            Map<String, Object> variables = new HashMap<>();
            Set<Object> validates = new HashSet<>();
            for (String param : params) {
                if (param.startsWith("CROSSFILEDUPLICATED(")) {
                    // ตรวจสอบว่าเป็นรูปแบบ array หรือไม่
                    String type = "CROSSFILEDUPLICATED";
                    Map<String, Object> sheetConfigs = ExtractParams.extractCrossFileCompare(param , type);
                    Map<String, Object> checkSheetDuplicate = new HashMap<>();
                    checkSheetDuplicate.put("type", type);
                    checkSheetDuplicate.put("detail", sheetConfigs);
                    validates.add(checkSheetDuplicate);
                }else if (param.startsWith("CROSSFILECOMPARE(")) {
                    String type = "CROSSFILECOMPARE";
                    Map<String, Object> sheetConfigs = ExtractParams.extractCrossFileCompare(param , type);
                    Map<String, Object> checkSheetDuplicate = new HashMap<>();
                    checkSheetDuplicate.put("type", type);
                    checkSheetDuplicate.put("detail", sheetConfigs);
                    validates.add(checkSheetDuplicate);
                }
            }

            if (!validates.isEmpty()) {
                variables.put("VALIDATES", validates);
            }
            
            if (variables.containsKey("VALIDATES")) {
                Object validateObj = variables.get("VALIDATES");
                if (validateObj instanceof Set) {
                    Set<?> validateSet = (Set<?>) validateObj;
                    for (Object validate : validateSet) {
                        if (validate instanceof Map) {
                            Map<?, ?> condition = (Map<?, ?>) validate;

                            if (condition.containsKey("type") && "CROSSFILEDUPLICATED".equals(condition.get("type"))) {
                                Object detailObj = condition.get("detail");
                                if (detailObj instanceof Map) {
                                    Map<?, ?> detail = (Map<?, ?>) detailObj;
                            
                                    // if ("EQUAL".equals(detail.get("condition"))) {
                                        Object sheetConfigObj = detail.get("sheetConfigs");
                                        if (sheetConfigObj instanceof List) {
                                            @SuppressWarnings("unchecked")
                                            List<Map<String, Object>> sheetConfigs = (List<Map<String, Object>>) sheetConfigObj;
                                            List<String> mismatchedValues = checkCrossFileEqualCondition(fileNameList, workbooks, sheetConfigs);

                                            if (!mismatchedValues.isEmpty()) {
                                                String errorCode = "CROSS_FILE_DUPLICATED";
                                                String errorMessage = "ข้อมูลในสองไฟล์มีข้อมูลซ้ำ";
                                                return createWssResult(mismatchedValues ,errorCode , errorMessage);
                                            }
                                        }
                                    // }
                                }
                            }else if (condition.containsKey("type") && "CROSSFILECOMPARE".equals(condition.get("type"))) { 
                                Object detailObj = condition.get("detail");
                                if (detailObj instanceof Map) {
                                    Map<?, ?> detail = (Map<?, ?>) detailObj;

                                    if ("EQUAL".equals(detail.get("condition"))) {
                                        Object sheetConfigObj = detail.get("sheetConfigs");
                                        if (sheetConfigObj instanceof List) {
                                            @SuppressWarnings("unchecked")
                                            List<Map<String, Object>> sheetConfigs = (List<Map<String, Object>>) sheetConfigObj;
                                            List<String> mismatchedValues = checkCrossCompare(fileNameList, workbooks, sheetConfigs);

                                            if (!mismatchedValues.isEmpty()) {
                                                String errorCode = "CROSS_FILE_COMPARE";
                                                String errorMessage = "ข้อมูลในสองไฟล์มีค่าไม่เท่ากัน";
                                                return createWssResult(mismatchedValues ,errorCode , errorMessage);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            JSONObject wssResult = new JSONObject();

            return wssResult.toString();
        } catch (IOException e) {
            // กรณีเกิดข้อผิดพลาด
            JSONObject wssResult = new JSONObject();
            JSONObject wssResultObj = new JSONObject();

            wssResultObj.put("data", new JSONObject());
            wssResultObj.put("status", "fail");

            JSONObject errorMessage = new JSONObject();
            errorMessage.put("errorCode", "E001");
            errorMessage.put("errorMessage", e.getMessage());

            wssResultObj.put("errorMessage", errorMessage);
            wssResult.put("wssResult", wssResultObj);

            return wssResult.toString();
        }
    }

    private static String createWssResult(List<String> mismatchedValues, String errorCode, String errorMessage) {
        JSONObject wssResultObj = new JSONObject();
        wssResultObj.put("status", "fail");
    
        JSONObject error = new JSONObject();
        error.put("errorCode", errorCode);
        error.put("errorMessage", errorMessage);
        error.put("notMatchedValues", mismatchedValues);
    
        wssResultObj.put("errorMessage", error);
        JSONObject wssResult = new JSONObject();
        wssResult.put("wssResult", wssResultObj);
        return wssResult.toString();
    }

    public static Workbook createWorkbook(File file, String fileName) throws IOException {
        try (FileInputStream fis = new FileInputStream(file)) {
            // ตรวจสอบประเภทไฟล์และสร้าง Workbook
            Workbook workbook;
            if (fileName.toLowerCase().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (fileName.toLowerCase().endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else {
                throw new IllegalArgumentException("ไฟล์ไม่ใช่รูปแบบ Excel (.xls หรือ .xlsx)");
            }
            return workbook;
        }
    }

    private static List<String> checkCrossCompare(List<String> fileNames, List<Workbook> workbooks,
        List<Map<String, Object>> sheetConfigs) throws IOException {
            if (sheetConfigs.size() != 2) {
                throw new IllegalArgumentException("CROSSFILEDUPLICATED ต้องมี 2 ชุดข้อมูล");
            }

            Map<String, Workbook> fileWorkbookMap = new HashMap<>();
            for (int i = 0; i < fileNames.size(); i++) {
                fileWorkbookMap.put(new File(fileNames.get(i)).getName(), workbooks.get(i));
            }

            Set<String> values1 = new HashSet<>();
            Set<String> values2 = new HashSet<>();

            for (int i = 0; i < 2; i++) {
                Map<String, Object> cfg = sheetConfigs.get(i);
                String fileName = (String) cfg.get("fileName");
                String sheetName = (String) cfg.get("sheetName");
                String rangeStr = (String) cfg.get("rangeStr");

                Workbook wb = fileWorkbookMap.get(fileName);
                if (wb == null) {
                    throw new IllegalArgumentException("ไม่พบ Workbook สำหรับไฟล์: " + fileName);
                }

                List<String> values = ExcelReader.getAllValuesInRange(wb, sheetName, rangeStr, null);
                if (i == 0) {
                    values1.addAll(values);
                } else {
                    values2.addAll(values);
                }
            }

            double sumValues1 = 0.0;
            double sumValues2 = 0.0;


            List<String> notEq = new ArrayList<>();
            for (String v : values1) {
                if (v != null && v.trim().matches("^-?\\d+(\\.\\d+)?$")) {
                    sumValues1 += Double.parseDouble(v.trim());
                }
            }
            for (String v : values2) {
                if (v != null && v.trim().matches("^-?\\d+(\\.\\d+)?$")) {
                    sumValues2 += Double.parseDouble(v.trim());
                }
            }


            if (Double.compare(sumValues1, sumValues2) != 0) {
                notEq.add(sumValues1 + " ผลรวมไม่เท่ากับ " + sumValues2);
            }

            return notEq; // ถ้าว่าง = เท่ากัน, ถ้าไม่ว่าง = มีความต่าง
    }

    private static List<String> checkCrossFileEqualCondition(
                List<String> fileNames,
                List<Workbook> workbooks,
                List<Map<String, Object>> sheetConfigs
        ) throws IOException {
            if (sheetConfigs.size() != 2) {
                throw new IllegalArgumentException("CROSSFILEDUPLICATED ต้องมี 2 ชุดข้อมูล");
            }

            Map<String, Workbook> fileWorkbookMap = new HashMap<>();
            for (int i = 0; i < fileNames.size(); i++) {
                fileWorkbookMap.put(new File(fileNames.get(i)).getName(), workbooks.get(i));
            }

            Set<String> values1 = new HashSet<>();
            Set<String> values2 = new HashSet<>();

            for (int i = 0; i < 2; i++) {
                Map<String, Object> cfg = sheetConfigs.get(i);
                String fileName = (String) cfg.get("fileName");
                String sheetName = (String) cfg.get("sheetName");
                String rangeStr = (String) cfg.get("rangeStr");

                Workbook wb = fileWorkbookMap.get(fileName);
                if (wb == null) {
                    throw new IllegalArgumentException("ไม่พบ Workbook สำหรับไฟล์: " + fileName);
                }

                List<String> values = ExcelReader.getAllValuesInRange(wb, sheetName, rangeStr, null);
                if (i == 0) {
                    values1.addAll(values);
                } else {
                    values2.addAll(values);
                }
            }

            // หาค่าที่ซ้ำกัน
            List<String> notMatched = new ArrayList<>();
            for (String v : values1) {
                if (values2.contains(v)) {
                    notMatched.add(v);
                }
            }

            return notMatched;
        }


    /**
     * ตัวอย่างการใช้งาน
     */
    public static void main(String[] args) {
        try {
            // กำหนดชื่อไฟล์ Excel
            // String fileName =
            // "/Users/anirut/Documents/anirut/AVLGB_Flile/ฐานข้อมูลการตั้งงบประมาณรายการท้องถิ่น/01007_1_1_01_เงินอุดหนุนสำหรับการจัดการศึกษาตั้งแต่ระดับอนุบาลจนจบการศึกษาขั้นพื้นฐาน.xlsx";
            List<String> fileName = new ArrayList<>();
            fileName.add(
                    "/Users/anirut/Downloads/01007_1_1_07_เงินอุดหนุนสำหรับการจัดการศึกษาภาคบังคับ (เงินเดือนครู ค่าจ้างประจำ).xlsx");
            // fileName.add("null");
            fileName.add(
                    "/Users/anirut/Downloads/01007_1_1_07_เงินอุดหนุนสำหรับการจัดการศึกษาภาคบังคับ EDIT.xlsx");
            String[] XXXXX_1_1_07 = {
                    // "FIX(\"FISCAL_YEAR\", \"2568\")", // ใช้ได้
                    // "FIX(\"TEMPLATE_ID\", \"21811913-aa9b-48f8-86cb-ed4041b17a99\")",// ใช้ได้
                    // "COLUMN(\"ITEM\", \"A7\", \"คำนวณเงินเดือน\")", // ใช้ได้
                    // "COLUMN(\"AMOUNT\", \"Z4\", \"คำนวณเงินเดือน\")",
                    // "COUNT(\"QTY\", \"A7:EOF\", \"คำนวณเงินเดือน\")", // ใช้ได้
                    // "CHKDUPLICATE(\"1.1 แบบชั้นเคลื่อน (ภาพรวม)\", A6:EOF, [\"ปวช.1\"])", //
                    // ใช้ได้
                    // "CHKDUPLICATE(\"คำนวณบำเหน็จบำนาญ\", C12:EOF, [])", // ใช้ได้
                    // "CHKSHEETDUPLICATE([(\"แบบคำนวณ (ภาพรวม)\", B15:EOF,
                    // [\"ค่าจัดการเรียนการสอน\"]), (\"Topup\", B15:EOF,
                    // [\"ค่าจัดการเรียนการสอน\"])])" // ใช้ได้
                    // "CHKSHEETDUPLICATE([(\"คำนวณเงินเดือน\", B15:EOF, [\"ตำแหน่งว่าง\"]),
                    // (\"คำนวณบำเหน็จบำนาญ\", C12:EOF, [])])" // ใช้ได้
                    "CROSSFILECOMPARE( EQUAL ,[ [\"01007_1_1_07_เงินอุดหนุนสำหรับการจัดการศึกษาภาคบังคับ (เงินเดือนครู ค่าจ้างประจำ).xlsx\", \"คำนวณบำเหน็จบำนาญ\", \"H12:EOF\"], [\"01007_1_1_07_เงินอุดหนุนสำหรับการจัดการศึกษาภาคบังคับ EDIT.xlsx\", \"คำนวณบำเหน็จบำนาญ\", \"H12:EOF\"]])"
            };
            // "CHKSHEETDUPLICATE([(\"แบบคำนวณ (ภาพรวม)\", B15:EOF,
            // [\"ค่าจัดการเรียนการสอน\"]), (\"Topup\", B15:EOF,
            // [\"ค่าจัดการเรียนการสอน\"])])"
            // "CROSSFILEDUPLICATED([ชื่อไฟล์ , ชื่อsheet , range ของข้อมูล, ค่าที่ยกเว้น] ,
            // [ชื่อไฟล์ , ชื่อsheet , range ของข้อมูล, ค่าที่ยกเว้น])"
            String wssResult = "";
            if (fileName.size() == 1) {
                wssResult = processExcelToWssResult(fileName.get(0), XXXXX_1_1_07);
            } else {
                wssResult = processMutiFileExcelToWssResult(fileName, XXXXX_1_1_07);
            }

            // String fileName =
            // "/Users/arthit/Downloads/อปท/1.การจัดบริการสาธารณะด้านการศึกษา/01007_1_1_01_เงินอุดหนุนสำหรับการจัดการศึกษาตั้งแต่ระดับอนุบาลจนจบการศึกษาขั้นพื้นฐาน.xlsx";
            // String[] XXXXX_1_1_01 = {
            // "COLUMN(\"ITEM\", \"A2\", \"แบบคำนวณ (ภาพรวม)\")",
            // "ROW(\"SUBITEM\", \"แบบคำนวณ (ภาพรวม)\", \"5:9\", \"[A,F,I]\",
            // [\"ITEM\",\"QTY\",\"AMOUNT\"])"
            // };
            // String wssResult = processExcelToWssResult(fileName, XXXXX_1_1_01);

            // String fileName =
            // "/Users/arthit/Downloads/อปท/1.การจัดบริการสาธารณะด้านการศึกษา/01007_1_1_02_เงินอุดหนุนสำหรับสนับสนุนค่าใช้จ่ายในการจัดการศึกษาสำหรับศูนย์พัฒนาเด็กเล็ก.xlsx";
            // String[] XXXXX_1_1_02 = {
            // "COLUMN(\"ITEM\", \"A8\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"AMOUNT\", \"D10\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"QTY\", \"B10\", \"5.1 แบบคำนวณ\")",

            // "COLUMN(\"SUBITEM1\", \"A12\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"AMOUNT1\", \"D17\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"QTY1\", \"B17\", \"5.1 แบบคำนวณ\")",

            // "COLUMN(\"SUBITEM2\", \"A19\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"AMOUNT2\", \"D23\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"QTY2\", \"B23\", \"5.1 แบบคำนวณ\")",

            // "COLUMN(\"SUBITEM3\", \"A24\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"AMOUNT3\", \"D28\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"QTY3\", \"B28\", \"5.1 แบบคำนวณ\")",

            // "COLUMN(\"SUBITEM4\", \"A29\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"AMOUNT4\", \"D33\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"QTY4\", \"B33\", \"5.1 แบบคำนวณ\")",

            // "COLUMN(\"SUBITEM5\", \"A34\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"AMOUNT5\", \"D38\", \"5.1 แบบคำนวณ\")",
            // "COLUMN(\"QTY5\", \"B38\", \"5.1 แบบคำนวณ\")"

            // };
            // String wssResult = processExcelToWssResult(fileName, XXXXX_1_1_02);

            // String fileName =
            // "/Users/arthit/Downloads/อปท/1.การจัดบริการสาธารณะด้านการศึกษา/01007_1_1_03_เงินอุดหนุนสำหรับสนับสนุนอาหารเสริม(นม).xlsx";
            // String[] XXXXX_1_1_03 = {
            // "COLUMN(\"ITEM\", \"A6\", \"3. หลักเกณฑ์ อาหารเสริม นม\")",
            // "COLUMN(\"AMOUNT\", \"F21\", \"3. หลักเกณฑ์ อาหารเสริม นม\")",
            // "COLUMN(\"QTY\", \"C21\", \"3. หลักเกณฑ์ อาหารเสริม นม\")",

            // "COLUMN(\"SUBITEM1\", \"A4\", \"3.1 แบบคำนวณ\")",
            // "COLUMN(\"AMOUNT1\", \"F8\", \"3.1 แบบคำนวณ\")",
            // "COLUMN(\"QTY1\", \"C8\", \"3.1 แบบคำนวณ\")",

            // "COLUMN(\"SUBITEM2\", \"A10\", \"3.1 แบบคำนวณ\")",
            // "COLUMN(\"AMOUNT2\", \"F20\", \"3.1 แบบคำนวณ\")",
            // "COLUMN(\"QTY2\", \"C20\", \"3.1 แบบคำนวณ\")"

            // };
            // String wssResult = processExcelToWssResult(fileName, XXXXX_1_1_03);

            // String fileName =
            // "/Users/arthit/Downloads/อปท/1.การจัดบริการสาธารณะด้านการศึกษา/01007_1_1_04_เงินอุดหนุนสำหรับสนับสนุนอาหารกลางวัน.xlsx";
            // String[] XXXXX_1_1_04 = {
            // "COLUMN(\"ITEM\", \"A8\", \"4.อาหารกลางวัน\")",
            // "COLUMN(\"AMOUNT\", \"F59\", \"4.1แบบคำนวณ\")",
            // "COLUMN(\"QTY\", \"C59\", \"4.1แบบคำนวณ\")",

            // "COLUMN(\"SUBITEM1\", \"A3\", \"4.1แบบคำนวณ\")",
            // "COLUMN(\"AMOUNT1\", \"F11\", \"4.1แบบคำนวณ\")",
            // "COLUMN(\"QTY1\", \"C11\", \"4.1แบบคำนวณ\")",

            // "COLUMN(\"SUBITEM2\", \"A13\", \"4.1แบบคำนวณ\")",
            // "COLUMN(\"AMOUNT2\", \"F58\", \"4.1แบบคำนวณ\")",
            // "COLUMN(\"QTY2\", \"C58\", \"4.1แบบคำนวณ\")"

            // };
            // String wssResult = processExcelToWssResult(fileName, XXXXX_1_1_04);

            // String fileName =
            // "/Users/arthit/Downloads/อปท/1.การจัดบริการสาธารณะด้านการศึกษา/XXXXX_1_1_05_เงินอุดหนุนสำหรับส่งเสริมศักยภาพการจัดการศึกษาท้องถิ่น
            // (ค่าปัจจัยพื้นฐานสำหรับนักเรียนยากจน).xlsx";
            // String[] XXXXX_1_1_05 = {
            // "COLUMN(\"ITEM\", \"A3\", \"11.ปัจจัยพื้นฐาน\")",
            // "ROWBY(A EQUAL \"รวมทั้งสิ้น\",[COLUMN(\"AMOUNT\", \"C?\",
            // \"11.ปัจจัยพื้นฐาน\"),COLUMN(\"AMOUNT\", \"F?\", \"11.ปัจจัยพื้นฐาน\")],
            // \"11.ปัจจัยพื้นฐาน\")"
            // };
            // String wssResult = processExcelToWssResult(fileName, XXXXX_1_1_05);

            // แสดงผลลัพธ์
            System.out.println(wssResult);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
