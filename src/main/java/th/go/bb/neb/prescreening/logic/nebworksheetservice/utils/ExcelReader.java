package th.go.bb.neb.prescreening.logic.nebworksheetservice.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * คลาสสำหรับเก็บพารามิเตอร์ในรูปแบบต่างๆ
 */
abstract class ExcelParameter {
    protected String name;
    
    public ExcelParameter(String name) {
        this.name = name;
    }
    
    public String getName() {
        return name;
    }
    
    public abstract Object process(ExcelReader reader, String fileName, String sheetName) throws IOException;
}

/**
 * พารามิเตอร์สำหรับค่าคงที่
 */
class FixParameter extends ExcelParameter {
    private String value;
    
    public FixParameter(String name, String value) {
        super(name);
        this.value = value;
    }
    
    @Override
    public Object process(ExcelReader reader, String fileName, String sheetName) {
        return value;
    }
    
    public static FixParameter parse(String name, String value) {
        return new FixParameter(name, value);
    }
}

/**
 * พารามิเตอร์สำหรับอ่านค่าจากคอลัมน์
 */
class ColumnParameter extends ExcelParameter {
    private String cellReference;
    
    public ColumnParameter(String name, String cellReference) {
        super(name);
        this.cellReference = cellReference;
    }
    
    @Override
    public Object process(ExcelReader reader, String fileName, String sheetName) throws IOException {
        File file = new File(fileName);
        Workbook workbook = createWorkbook(file , fileName);
        return ExcelReader.readCellValue(workbook, sheetName, cellReference);
    }

    public static Workbook createWorkbook(File file , String fileName) throws IOException {
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
    
    public static ColumnParameter parse(String name, String cellReference) {
        return new ColumnParameter(name, cellReference);
    }
}

/**
 * พารามิเตอร์สำหรับนับจำนวนเซลล์
 */
class CountParameter extends ExcelParameter {
    private String rangeStr;
    
    public CountParameter(String name, String rangeStr) {
        super(name);
        this.rangeStr = rangeStr;
    }
    
    @Override
    public Object process(ExcelReader reader, String fileName, String sheetName) throws IOException {
        return ExcelReader.countCells(fileName, sheetName, rangeStr);
    }
    
    public static CountParameter parse(String name, String rangeStr) {
        return new CountParameter(name, rangeStr);
    }
}

/**
 * พารามิเตอร์สำหรับตรวจสอบค่าซ้ำ
 */
class CheckDuplicateParameter extends ExcelParameter {
    private String targetSheetName;
    private String rangeStr;
    private Set<Object> exceptValues;
    
    public CheckDuplicateParameter(String name, String targetSheetName, String rangeStr, Set<Object> exceptValues) {
        super(name);
        this.targetSheetName = targetSheetName;
        this.rangeStr = rangeStr;
        this.exceptValues = exceptValues;
    }
    
    public String getTargetSheetName() {
        return targetSheetName;
    }
    
    public String getRangeStr() {
        return rangeStr;
    }
    
    public Set<Object> getExceptValues() {
        return exceptValues;
    }
    
    @Override
    public Object process(ExcelReader reader, String fileName, String sheetName) throws IOException {
        // ไม่ได้ใช้ process โดยตรง แต่จะใช้ข้อมูลจาก getter
        return null;
    }
    
    public static CheckDuplicateParameter parse(String name, String targetSheetName, String rangeStr, List<String> exceptValuesList) {
        Set<Object> exceptValues = null;
        if (exceptValuesList != null && !exceptValuesList.isEmpty()) {
            exceptValues = new HashSet<>(exceptValuesList);
        }
        return new CheckDuplicateParameter(name, targetSheetName, rangeStr, exceptValues);
    }
}

public class ExcelReader {

    /**
     * อ่านข้อมูลจากไฟล์ Excel ตามพารามิเตอร์ที่ระบุ
     * 
     * @param fileName ชื่อไฟล์ Excel
     * @param sheetName ชื่อชีทที่ต้องการอ่าน
     * @param variables Map ของตัวแปรและคำสั่งที่ใช้ในการอ่าน เช่น {"AMOUNT": "Z4", "QTY": "COUNT B15:EOF"}
     * @return Map ของผลลัพธ์การอ่าน เช่น {"amount": 999, "qty": 99, "agencyCode": "01007"}
     * @throws IOException หากมีข้อผิดพลาดในการอ่านไฟล์
     */
    public static Map<String, Object> readExcelVariables( Workbook workbook,String fileName, String sheetName, 
                                                        Map<String, Object> variables) throws IOException {
        Map<String, Object> result = new HashMap<>();
        
        // ดึง agency code จากชื่อไฟล์ (5 ตัวอักษรแรก)
        File file = new File(fileName);
        String baseName = file.getName();
        String agencyCode = "";
        if (baseName.length() >= 5) {
            agencyCode = baseName.substring(0, 5);
        }
        result.put("agencyCode", agencyCode);

            try {
                // เลือกชีท (ถ้ามีการระบุชื่อชีท)
                Sheet sheet = null;
                if (!sheetName.isEmpty()) {
                    sheet = workbook.getSheet(sheetName);
                    if (sheet == null) {
                        throw new IllegalArgumentException("ไม่พบชีท '" + sheetName + "' ในไฟล์");
                    }
                }
                
                // ประมวลผลแต่ละตัวแปร
                for (Map.Entry<String, Object> entry : variables.entrySet()) {
                    String varName = entry.getKey().toLowerCase();
                    
                    // ข้ามการประมวลผลสำหรับตัวแปร VALIDATES
                    if ("validates".equals(varName)) {
                        continue;
                    }
                    
                    Object value = entry.getValue();
                    
                    // กรณีที่ค่าเป็น String
                    if (value instanceof String) {
                        String instruction = ((String) value).trim();
                        
                        // กรณีที่เป็นค่าคงที่ (FIX)
                        if (instruction.startsWith("FIX ")) {
                            String fixedValue = instruction.substring(4).trim(); // ตัด "FIX " ออก
                            result.put(varName, fixedValue);
                        }
                        // กรณีที่เป็นการอ่านเซลล์เดียว (เช่น Z4)
                        else if (isSingleCellReference(instruction)) {
                            Object cellValue = readCellValue(sheet, instruction);
                            result.put(varName, cellValue);
                        }
                        // กรณีที่เป็นการนับจำนวน (COUNT B15:EOF หรือ COUNT B15:B100000)
                        else if (instruction.toUpperCase().startsWith("COUNT ")) {
                            String rangeStr = instruction.substring(6).trim(); // ตัด "COUNT " ออก
                            int count = countNonEmptyCells(sheet, rangeStr);
                            result.put(varName, count);
                        }
                        // กรณีอื่นๆ
                        else {
                            result.put(varName, "คำสั่งไม่รองรับ: " + instruction);
                        }
                    }
                    // กรณีที่ค่าเป็น Map และเป็นการนับจำนวน (COUNT)
                    else if (value instanceof Map && ((Map<?, ?>) value).containsKey("type") && "COUNT".equals(((Map<?, ?>) value).get("type"))) {
                        Map<String, Object> countConfig = (Map<String, Object>) value;
                        String rangeStr = (String) countConfig.get("rangeStr");
                        
                        try {
                            int count;
                            
                            // ตรวจสอบว่ามีการระบุชื่อชีทหรือไม่
                            if (countConfig.containsKey("sheetName")) {
                                String targetSheetName = (String) countConfig.get("sheetName");
                                count = countCells(fileName, targetSheetName, rangeStr);
                            } else {
                                // ใช้ชีทหลัก
                                count = countNonEmptyCells(sheet, rangeStr);
                            }
                            
                            // เก็บผลลัพธ์
                            result.put(varName, count);
                        } catch (IllegalArgumentException e) {
                            // กรณีไม่พบชีทหรือข้อผิดพลาดอื่นๆ
                            result.put(varName, "ข้อผิดพลาด: " + e.getMessage());
                        }
                    }
                    // กรณีที่ค่าเป็น Map และเป็นการอ่านเซลล์เดียวจากชีทที่กำหนด (COLUMN)
                    else if (value instanceof Map && ((Map<?, ?>) value).containsKey("type") && "COLUMN".equals(((Map<?, ?>) value).get("type"))) {
                        Map<String, Object> columnConfig = (Map<String, Object>) value;
                        String cellRef = (String) columnConfig.get("cellRef");
                        
                        try {
                            Object cellValue;
                            
                            // ตรวจสอบว่ามีการระบุชื่อชีทหรือไม่
                            if (columnConfig.containsKey("sheetName")) {
                                String targetSheetName = (String) columnConfig.get("sheetName");
                                cellValue = readCellValue(workbook , targetSheetName, cellRef);
                            } else {
                                // ใช้ชีทหลัก
                                cellValue = readCellValue(sheet, cellRef);
                            }
                            
                            // เก็บผลลัพธ์
                            result.put(varName, cellValue);
                        } catch (IllegalArgumentException e) {
                            // กรณีไม่พบชีทหรือข้อผิดพลาดอื่นๆ
                            result.put(varName, "ข้อผิดพลาด: " + e.getMessage());
                        }
                    }
                    // กรณีที่ค่าเป็น Map และเป็นการค้นหาแถวตามเงื่อนไข (ROWBY)
                    else if (value instanceof Map && ((Map<?, ?>) value).containsKey("type") && "ROWBY".equals(((Map<?, ?>) value).get("type"))) {
                        Map<String, Object> rowByConfig = (Map<String, Object>) value;
                        String searchSheetName = (String) rowByConfig.get("searchSheetName");
                        String readSheetName = (String) rowByConfig.get("readSheetName");
                        String searchColumn = (String) rowByConfig.get("searchColumn");
                        String searchCondition = (String) rowByConfig.get("searchCondition");
                        String searchValue = (String) rowByConfig.get("searchValue");
                        String columnRefsStr = (String) rowByConfig.get("columnRefs");
                        
                        try {
                            // แยกคอลัมน์ที่ต้องการอ่านจากสตริง (เช่น "[COLUMN(\"AMOUNT\", \"C?\"),COLUMN(\"AMOUNT\", \"F?\")]")
                            if (!columnRefsStr.startsWith("[") || !columnRefsStr.endsWith("]")) {
                                throw new IllegalArgumentException("รูปแบบคอลัมน์ไม่ถูกต้อง: " + columnRefsStr);
                            }
                            
                            String columnRefsContent = columnRefsStr.substring(1, columnRefsStr.length() - 1);
                            
                            // แยกแต่ละ COLUMN
                            List<String> columnRefs = new ArrayList<>();
                            StringBuilder currentColumnRef = new StringBuilder();
                            boolean inQuotes = false;
                            int parenthesisCount = 0;
                            
                            for (int i = 0; i < columnRefsContent.length(); i++) {
                                char c = columnRefsContent.charAt(i);
                                
                                if (c == '"') {
                                    inQuotes = !inQuotes;
                                    currentColumnRef.append(c);
                                } else if (c == '(') {
                                    parenthesisCount++;
                                    currentColumnRef.append(c);
                                } else if (c == ')') {
                                    parenthesisCount--;
                                    currentColumnRef.append(c);
                                } else if (c == ',' && !inQuotes && parenthesisCount == 0) {
                                    // พบตัวคั่น COLUMN
                                    // String columnRef = extractColumnRef(currentColumnRef.toString().trim());
                                    // if (columnRef != null) {
                                        columnRefs.add(currentColumnRef.toString().trim());
                                        // columnRefs.add(columnRef);
                                    // }
                                    currentColumnRef = new StringBuilder();
                                } else {
                                    currentColumnRef.append(c);
                                }
                            }
                            
                            // เพิ่ม COLUMN สุดท้าย
                            if (currentColumnRef.length() > 0) {
                                // String columnRef = extractColumnRef(currentColumnRef.toString().trim());
                                // if (columnRef != null) {
                                    columnRefs.add(currentColumnRef.toString().trim());
                                // }
                            }
                            
                            // ค้นหาแถวและอ่านค่า
                            List<Object> rowValues = findRowByConditionAndReadColumns(fileName, searchSheetName, readSheetName,
                                                                                   searchColumn, searchCondition, searchValue, columnRefs);
                            
                            // เก็บผลลัพธ์
                            result.put(varName, rowValues);
                        } catch (IllegalArgumentException e) {
                            // กรณีไม่พบชีทหรือข้อผิดพลาดอื่นๆ
                            result.put(varName, "ข้อผิดพลาด: " + e.getMessage());
                        }
                    }
                    // กรณีที่ค่าเป็น Map และเป็นการอ่านแถว (ROW)
                    else if (value instanceof Map && ((Map<?, ?>) value).containsKey("type") && "ROW".equals(((Map<?, ?>) value).get("type"))) {
                        Map<String, Object> rowConfig = (Map<String, Object>) value;
                        String targetSheetName = (String) rowConfig.get("sheetName");
                        String rowRangeStr = (String) rowConfig.get("rowRange");
                        String columnsStr = (String) rowConfig.get("columns");
                        String mappingStr = (String) rowConfig.get("mapping");
                        
                        try {
                            // อ่านข้อมูลจากแถวและคอลัมน์ที่กำหนด
                            List<Map<String, Object>> rowsData;
                            if (mappingStr != null) {
                                rowsData = readRowsAndColumnsWithMapping(fileName, targetSheetName, rowRangeStr, columnsStr, mappingStr);
                            } else {
                                rowsData = readRowsAndColumns(fileName, targetSheetName, rowRangeStr, columnsStr);
                            }
                            
                            // เก็บผลลัพธ์
                            result.put(varName, rowsData);
                        } catch (IllegalArgumentException e) {
                            // กรณีไม่พบชีทหรือข้อผิดพลาดอื่นๆ
                            result.put(varName, "ข้อผิดพลาด: " + e.getMessage());
                        }
                    }
                    else {
                        // กรณีที่ค่าไม่ใช่ String ให้เก็บค่าเดิม
                        result.put(varName, value);
                    }
                }
            } finally {
                workbook.close();
            }
        
        return result;
    }

    
    
    /**
     * อ่านค่าจากเซลล์ที่ระบุในไฟล์ Excel
     * 
     * @param fileName ชื่อไฟล์ Excel
     * @param sheetName ชื่อชีทที่ต้องการอ่าน
     * @param cellReference ตำแหน่งเซลล์ (เช่น A1, B5, Z4)
     * @return ค่าในเซลล์
     * @throws IOException หากมีข้อผิดพลาดในการอ่านไฟล์
     */
    public static Object readCellValue(Workbook workbook, String sheetName, String cellReference) throws IOException {
        // File file = new File(fileName);
        // try (FileInputStream fis = new FileInputStream(file)) {
        //     // ตรวจสอบประเภทไฟล์และสร้าง Workbook
        //     Workbook workbook;
        //     if (fileName.toLowerCase().endsWith(".xlsx")) {
        //         workbook = new XSSFWorkbook(fis);
        //     } else if (fileName.toLowerCase().endsWith(".xls")) {
        //         workbook = new HSSFWorkbook(fis);
        //     } else {
        //         throw new IllegalArgumentException("ไฟล์ไม่ใช่รูปแบบ Excel (.xls หรือ .xlsx)");
        //     }
            
            try {
                // เลือกชีท
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    throw new IllegalArgumentException("ไม่พบชีท '" + sheetName + "' ในไฟล์");
                }
                
                return readCellValue(sheet, cellReference);
            } finally {
                workbook.close();
            }
        // }
    }
    
    /**
     * นับจำนวนเซลล์ที่ไม่ว่างในช่วงที่กำหนดในไฟล์ Excel
     * 
     * @param fileName ชื่อไฟล์ Excel
     * @param sheetName ชื่อชีทที่ต้องการอ่าน
     * @param rangeStr ช่วงเซลล์ที่ต้องการนับ (เช่น B15:EOF)
     * @return จำนวนเซลล์ที่ไม่ว่าง
     * @throws IOException หากมีข้อผิดพลาดในการอ่านไฟล์
     */
    public static int countCells(String fileName, String sheetName, String rangeStr) throws IOException {
        File file = new File(fileName);
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
            
            try {
                // เลือกชีท
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    throw new IllegalArgumentException("ไม่พบชีท '" + sheetName + "' ในไฟล์");
                }
                
                return countNonEmptyCells(sheet, rangeStr);
            } finally {
                workbook.close();
            }
        }
    }
    
    /**
     * ตรวจสอบว่าเป็นการอ้างอิงเซลล์เดียวหรือไม่ (เช่น A1, B2, Z4)
     */
    private static boolean isSingleCellReference(String reference) {
        return reference.matches("[A-Za-z]+[0-9]+");
    }
    
    /**
     * อ่านค่าจากเซลล์ที่ระบุ
     */
    private static Object readCellValue(Sheet sheet, String cellReference) {
        CellReference ref = new CellReference(cellReference);
        Row row = sheet.getRow(ref.getRow());
        if (row == null) {
            return null;
        }
        
        Cell cell = row.getCell(ref.getCol());
        if (cell == null) {
            return null;
        }
        
        return getCellValue(cell);
    }
    
    /**
     * ดึงค่าจากเซลล์ตามประเภทข้อมูล
     */
    private static Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    double numValue = cell.getNumericCellValue();
                    // ถ้าเป็นจำนวนเต็ม ส่งกลับเป็น Integer/Long
                    if (numValue == Math.floor(numValue)) {
                        if (numValue <= Integer.MAX_VALUE && numValue >= Integer.MIN_VALUE) {
                            return (int) numValue;
                        } else {
                            return (long) numValue;
                        }
                    }
                    return numValue;
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                try {
                    return cell.getNumericCellValue();
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception ex) {
                        return cell.getCellFormula();
                    }
                }
            case BLANK:
                return null;
            default:
                return null;
        }
    }
    
    /**
     * นับจำนวนเซลล์ที่ไม่ว่างในช่วงที่กำหนด
     */
    private static int countNonEmptyCells(Sheet sheet, String rangeStr) {
        // แยกช่วงจากสตริง (เช่น "B15:B100000" หรือ "B15:EOF")
        String[] rangeParts = rangeStr.split(":");
        if (rangeParts.length != 2) {
            throw new IllegalArgumentException("รูปแบบช่วงไม่ถูกต้อง: " + rangeStr);
        }
        
        String startCellRef = rangeParts[0].trim();
        String endCellRef = rangeParts[1].trim();
        
        // แยกคอลัมน์และแถวจากเซลล์เริ่มต้น
        Matcher startMatcher = Pattern.compile("([A-Za-z]+)([0-9]+)").matcher(startCellRef);
        if (!startMatcher.matches()) {
            throw new IllegalArgumentException("รูปแบบเซลล์เริ่มต้นไม่ถูกต้อง: " + startCellRef);
        }
        
        String startColStr = startMatcher.group(1);
        int startRow = Integer.parseInt(startMatcher.group(2)) - 1; // แปลงเป็น 0-based
        CellReference startCellReference = new CellReference(startCellRef);
        int startColIdx = startCellReference.getCol();
        
        int endRow;
        if ("EOF".equalsIgnoreCase(endCellRef)) {
            // หาแถวสุดท้ายที่มีข้อมูลในคอลัมน์ที่กำหนด
            endRow = sheet.getLastRowNum();
            boolean foundData = false;
            for (int r = endRow; r >= startRow; r--) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    Cell cell = row.getCell(startColIdx);
                    if (cell != null && !isEmpty(cell)) {
                        endRow = r;
                        foundData = true;
                        break;
                    }
                }
            }
            
            if (!foundData) {
                return 0; // ไม่พบข้อมูลเลย
            }
        } else {
            // ใช้ค่าที่ระบุ
            Matcher endMatcher = Pattern.compile("([A-Za-z]+)([0-9]+)").matcher(endCellRef);
            if (!endMatcher.matches()) {
                throw new IllegalArgumentException("รูปแบบเซลล์สิ้นสุดไม่ถูกต้อง: " + endCellRef);
            }
            endRow = Integer.parseInt(endMatcher.group(2)) - 1; // แปลงเป็น 0-based
        }
        
        // นับจำนวนเซลล์ที่ไม่ว่าง
        int count = 0;
        for (int r = startRow; r <= endRow; r++) {
            Row row = sheet.getRow(r);
            if (row != null) {
                Cell cell = row.getCell(startColIdx);
                if (cell != null && !isEmpty(cell)) {
                    count++;
                }
            }
        }
        
        return count;
    }
    
    /**
     * ตรวจสอบว่าเซลล์ว่างหรือไม่
     */
    private static boolean isEmpty(Cell cell) {
        if (cell == null) {
            return true;
        }
        
        switch (cell.getCellType()) {
            case BLANK:
                return true;
            case STRING:
                return cell.getStringCellValue().trim().isEmpty();
            default:
                return false;
        }
    }
    
    /**
     * ค้นหาแถวที่มีค่าตรงตามเงื่อนไขในคอลัมน์ที่กำหนด และอ่านค่าจากคอลัมน์อื่นๆ ในแถวนั้น
     * 
     * @param fileName ชื่อไฟล์ Excel
     * @param searchSheetName ชื่อชีทที่ต้องการค้นหา
     * @param readSheetName ชื่อชีทที่ต้องการอ่านข้อมูล
     * @param searchColumn คอลัมน์ที่ต้องการค้นหา (เช่น "A")
     * @param searchCondition เงื่อนไขการค้นหา (EQUAL, STARTWITH, ENDWITH, CONTENT)
     * @param searchValue ค่าที่ต้องการค้นหา
     * @param columnRefs รายการคอลัมน์ที่ต้องการอ่านค่า (เช่น ["C?", "F?"])
     * @return List ของค่าที่อ่านได้จากคอลัมน์ที่กำหนด โดยแทนที่ ? ด้วยเลขแถวที่พบ
     * @throws IOException หากมีข้อผิดพลาดในการอ่านไฟล์
     */
    public static List<Object> findRowByConditionAndReadColumns(String fileName, String searchSheetName, String readSheetName,
                                                              String searchColumn, String searchCondition, String searchValue,
                                                              List<String> columnRefs) throws IOException {
        System.out.println("findRowByConditionAndReadColumns Start : "+readSheetName);
        File file = new File(fileName);
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
            
            try {
                // เลือกชีทสำหรับค้นหา
                Sheet searchSheet = workbook.getSheet(searchSheetName);
                if (searchSheet == null) {
                    throw new IllegalArgumentException("ไม่พบชีท '" + searchSheetName + "' ในไฟล์");
                }
                
                // เลือกชีทสำหรับอ่านข้อมูล
                Sheet readSheet = workbook.getSheet(readSheetName);
                if (readSheet == null) {
                    throw new IllegalArgumentException("ไม่พบชีท '" + readSheetName + "' ในไฟล์");
                }
                
                System.out.println("searchColumn ::"+searchColumn + "1");
                // แปลงชื่อคอลัมน์เป็น index
                CellReference searchColRef = new CellReference(searchColumn + "1");
                int searchColIdx = searchColRef.getCol();
                
                // ค้นหาแถวที่ตรงตามเงื่อนไข
                int foundRowIdx = -1;
                int foundRowNum = -1;
                
                for (int rowIdx = 0; rowIdx <= searchSheet.getLastRowNum(); rowIdx++) {
                    System.out.println("Loop search data :: Row = "+rowIdx+" valus "+searchValue);
                    Row row = searchSheet.getRow(rowIdx);
                    if (row == null) continue;
                    
                    Cell cell = row.getCell(searchColIdx);
                    if (cell == null || isEmpty(cell)) continue;
                    
                    Object cellValue = getCellValue(cell);
                    String cellStrValue = cellValue != null ? cellValue.toString().trim() : "";
                    
                    boolean match = false;
                    
                    // ตรวจสอบเงื่อนไข
                    switch (searchCondition.toUpperCase()) {
                        case "EQUAL":
                            match = cellStrValue.equals(searchValue);
                            break;
                        case "STARTWITH":
                            match = cellStrValue.startsWith(searchValue);
                            break;
                        case "ENDWITH":
                            match = cellStrValue.endsWith(searchValue);
                            break;
                        case "CONTENT":
                            match = cellStrValue.contains(searchValue);
                            break;
                        default:
                            throw new IllegalArgumentException("เงื่อนไขไม่ถูกต้อง: " + searchCondition);
                    }
                    
                    if (match) {
                        foundRowIdx = rowIdx;
                        foundRowNum = rowIdx + 1; // แปลงเป็น 1-based
                        break;
                    }
                }
                
                if (foundRowIdx == -1) {
                    // ไม่พบแถวที่ตรงตามเงื่อนไข
                    return new ArrayList<>();
                }
                System.out.println("====================== foundRowNum:"+foundRowNum);
                // อ่านค่าจากคอลัมน์ที่กำหนด
                List<Object> result = new ArrayList<>();
                
                for (String columnRef : columnRefs) {
                    String columnRefNameStr = extractColumnRef(columnRef.toString().trim());
                    // if (columnRefRead != null) {
                        // columnRefs.add(currentColumnRef.toString().trim());
                    // }
                    System.out.println("=================columnRef::"+columnRef);
                    System.out.println("=================columnRef::"+extractFirstPart(columnRef));
                    System.out.println("=================columnRefNameStr::"+columnRefNameStr);
                    // แทนที่ ? ด้วยเลขแถวที่พบ
                    String actualColumnRef = columnRefNameStr.replace("?", String.valueOf(foundRowNum));
                    System.out.println("====================== actualColumnRef:"+actualColumnRef);
                
                    // แยกชื่อคอลัมน์และเลขแถว
                    Matcher matcher = Pattern.compile("([A-Za-z]+)([0-9]+)").matcher(actualColumnRef);
                    if (!matcher.matches()) {
                        throw new IllegalArgumentException("รูปแบบคอลัมน์ไม่ถูกต้อง: " + actualColumnRef);
                    }
                    
                    String colName = matcher.group(1);
                    int rowNum = Integer.parseInt(matcher.group(2)) - 1; // แปลงเป็น 0-based
                    
                    CellReference cellRef = new CellReference(actualColumnRef);
                    int colIdx = cellRef.getCol();
                    
                    // อ่านค่าจากเซลล์
                    Row row = readSheet.getRow(rowNum);
                    Object value = null;
                    
                    if (row != null) {
                        Cell cell = row.getCell(colIdx);
                        value = getCellValue(cell);
                        System.out.println("Read value :: "+value);
                    }
                    
                    result.add(value);
                }
                
                return result;
            } finally {
                workbook.close();
            }
        }
    }
    
    /**
     * แยกส่วนแรกจากสตริงรูปแบบ COLUMN("part1", "part2", ...)
     * @param input สตริงที่ต้องการแยก
     * @return ส่วนแรกที่แยกได้
     */
    public static String extractFirstPart(String input) {
        if (input == null || !input.startsWith("COLUMN(") || !input.endsWith(")")) {
            return null;
        }
        
        // ลบ COLUMN( จากด้านหน้าและ ) จากด้านหลัง
        String content = input.substring(7, input.length() - 1);
        
        // แยกด้วย comma
        String[] parts = content.split(",\\s*");
        
        if (parts.length > 0) {
            // เอาเครื่องหมาย " ออกจาก part[0]
            return parts[0].replaceAll("\"", "");
        }
        
        return null;
    }

    /**
     * อ่านข้อมูลจากแถวและคอลัมน์ที่กำหนด
     * 
     * @param fileName ชื่อไฟล์ Excel
     * @param sheetName ชื่อชีทที่ต้องการอ่าน
     * @param rowRangeStr ช่วงแถวที่ต้องการอ่าน (เช่น "5:9")
     * @param columnsStr คอลัมน์ที่ต้องการอ่าน (เช่น "[A,F,I]")
     * @return List ของ Map ที่เก็บข้อมูลแต่ละแถว โดยแต่ละ Map มี key เป็นชื่อคอลัมน์ และ value เป็นค่าในเซลล์
     * @throws IOException หากมีข้อผิดพลาดในการอ่านไฟล์
     */
    public static List<Map<String, Object>> readRowsAndColumns(String fileName, String sheetName, 
                                                             String rowRangeStr, String columnsStr) throws IOException {
        File file = new File(fileName);
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
            
            try {
                // เลือกชีท
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    throw new IllegalArgumentException("ไม่พบชีท '" + sheetName + "' ในไฟล์");
                }
                
                // แยกช่วงแถวจากสตริง (เช่น "5:9")
                String[] rowRangeParts = rowRangeStr.split(":");
                if (rowRangeParts.length != 2) {
                    throw new IllegalArgumentException("รูปแบบช่วงแถวไม่ถูกต้อง: " + rowRangeStr);
                }
                
                int startRow = Integer.parseInt(rowRangeParts[0]) - 1; // แปลงเป็น 0-based
                int endRow = Integer.parseInt(rowRangeParts[1]) - 1; // แปลงเป็น 0-based
                
                // แยกคอลัมน์จากสตริง (เช่น "[A,F,I]")
                if (!columnsStr.startsWith("[") || !columnsStr.endsWith("]")) {
                    throw new IllegalArgumentException("รูปแบบคอลัมน์ไม่ถูกต้อง: " + columnsStr);
                }
                
                String columnsContent = columnsStr.substring(1, columnsStr.length() - 1);
                String[] columnParts = columnsContent.split(",");
                
                // แปลงชื่อคอลัมน์เป็น index
                int[] columnIndices = new int[columnParts.length];
                for (int i = 0; i < columnParts.length; i++) {
                    String colName = columnParts[i].trim();
                    CellReference cellReference = new CellReference(colName + "1");
                    columnIndices[i] = cellReference.getCol();
                }
                
                // อ่านข้อมูลจากแถวและคอลัมน์ที่กำหนด
                List<Map<String, Object>> result = new ArrayList<>();
                
                for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
                    Row row = sheet.getRow(rowIdx);
                    if (row == null) continue;
                    
                    Map<String, Object> rowData = new HashMap<>();
                    
                    for (int i = 0; i < columnIndices.length; i++) {
                        int colIdx = columnIndices[i];
                        Cell cell = row.getCell(colIdx);
                        
                        // ใช้ชื่อคอลัมน์เป็น key
                        String colName = columnParts[i].trim();
                        
                        // อ่านค่าจากเซลล์
                        Object cellValue = getCellValue(cell);
                        
                        // เก็บค่าลงใน Map
                        rowData.put(colName, cellValue);
                    }
                    
                    // เพิ่ม Map ของแถวนี้เข้าไปใน List ผลลัพธ์
                    result.add(rowData);
                }
                
                return result;
            } finally {
                workbook.close();
            }
        }
    }
    
    /**
     * อ่านข้อมูลจากแถวและคอลัมน์ที่กำหนด โดยใช้ mapping ในการแปลงชื่อคอลัมน์เป็นชื่อตัวแปร
     * 
     * @param fileName ชื่อไฟล์ Excel
     * @param sheetName ชื่อชีทที่ต้องการอ่าน
     * @param rowRangeStr ช่วงแถวที่ต้องการอ่าน (เช่น "5:9")
     * @param columnsStr คอลัมน์ที่ต้องการอ่าน (เช่น "[A,F,I]")
     * @param mappingStr การแปลงชื่อคอลัมน์เป็นชื่อตัวแปร (เช่น "[\"ITEM\",\"QTY\",\"AMOUNT\"]")
     * @return List ของ Map ที่เก็บข้อมูลแต่ละแถว โดยแต่ละ Map มี key เป็นชื่อตัวแปรตาม mapping และ value เป็นค่าในเซลล์
     * @throws IOException หากมีข้อผิดพลาดในการอ่านไฟล์
     */
    public static List<Map<String, Object>> readRowsAndColumnsWithMapping(String fileName, String sheetName, 
                                                                        String rowRangeStr, String columnsStr, String mappingStr) throws IOException {
        File file = new File(fileName);
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
            
            try {
                // เลือกชีท
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    throw new IllegalArgumentException("ไม่พบชีท '" + sheetName + "' ในไฟล์");
                }
                
                // แยกช่วงแถวจากสตริง (เช่น "5:9")
                String[] rowRangeParts = rowRangeStr.split(":");
                if (rowRangeParts.length != 2) {
                    throw new IllegalArgumentException("รูปแบบช่วงแถวไม่ถูกต้อง: " + rowRangeStr);
                }
                
                int startRow = Integer.parseInt(rowRangeParts[0]) - 1; // แปลงเป็น 0-based
                int endRow = Integer.parseInt(rowRangeParts[1]) - 1; // แปลงเป็น 0-based
                
                // แยกคอลัมน์จากสตริง (เช่น "[A,F,I]")
                if (!columnsStr.startsWith("[") || !columnsStr.endsWith("]")) {
                    throw new IllegalArgumentException("รูปแบบคอลัมน์ไม่ถูกต้อง: " + columnsStr);
                }
                
                String columnsContent = columnsStr.substring(1, columnsStr.length() - 1);
                String[] columnParts = columnsContent.split(",");
                
                // แยก mapping จากสตริง (เช่น "[\"ITEM\",\"QTY\",\"AMOUNT\"]")
                if (!mappingStr.startsWith("[") || !mappingStr.endsWith("]")) {
                    throw new IllegalArgumentException("รูปแบบ mapping ไม่ถูกต้อง: " + mappingStr);
                }
                
                String mappingContent = mappingStr.substring(1, mappingStr.length() - 1);
                String[] mappingParts = mappingContent.split(",");
                
                // ตรวจสอบว่าจำนวนคอลัมน์และ mapping ตรงกัน
                if (columnParts.length != mappingParts.length) {
                    throw new IllegalArgumentException("จำนวนคอลัมน์และ mapping ไม่ตรงกัน: " + columnParts.length + " vs " + mappingParts.length);
                }
                
                // แปลงชื่อคอลัมน์เป็น index
                int[] columnIndices = new int[columnParts.length];
                for (int i = 0; i < columnParts.length; i++) {
                    String colName = columnParts[i].trim();
                    CellReference cellReference = new CellReference(colName + "1");
                    columnIndices[i] = cellReference.getCol();
                }
                
                // แปลง mapping เป็นชื่อตัวแปร
                String[] variableNames = new String[mappingParts.length];
                for (int i = 0; i < mappingParts.length; i++) {
                    String varName = mappingParts[i].trim();
                    if (varName.startsWith("\"") && varName.endsWith("\"")) {
                        varName = varName.substring(1, varName.length() - 1);
                    }
                    variableNames[i] = varName;
                }
                
                // อ่านข้อมูลจากแถวและคอลัมน์ที่กำหนด
                List<Map<String, Object>> result = new ArrayList<>();
                
                for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
                    Row row = sheet.getRow(rowIdx);
                    if (row == null) continue;
                    
                    Map<String, Object> rowData = new HashMap<>();
                    
                    for (int i = 0; i < columnIndices.length; i++) {
                        int colIdx = columnIndices[i];
                        Cell cell = row.getCell(colIdx);
                        
                        // ใช้ชื่อตัวแปรจาก mapping เป็น key
                        String varName = variableNames[i];
                        
                        // อ่านค่าจากเซลล์
                        Object cellValue = getCellValue(cell);
                        
                        // เก็บค่าลงใน Map
                        rowData.put(varName, cellValue);
                    }
                    
                    // เพิ่ม Map ของแถวนี้เข้าไปใน List ผลลัพธ์
                    result.add(rowData);
                }
                
                return result;
            } finally {
                workbook.close();
            }
        }
    }
    
    /**
     * ตรวจสอบค่าซ้ำในช่วงเซลล์ที่กำหนด โดยสามารถระบุค่าที่ยกเว้นไม่ต้องตรวจสอบได้
     * 
     * @param fileName ชื่อไฟล์ Excel
     * @param sheetName ชื่อชีทที่ต้องการตรวจสอบ
     * @param rangeStr ช่วงเซลล์ที่ต้องการตรวจสอบ (เช่น "A1:A100" หรือ "B5:D20")
     * @param exceptValues ค่าที่ยกเว้นไม่ต้องตรวจสอบการซ้ำ (สามารถเป็น null ถ้าไม่มีค่ายกเว้น)
     * @return Map ที่มีข้อมูลเกี่ยวกับค่าซ้ำที่พบ โดยมี key เป็นค่าที่ซ้ำ และ value เป็น List ของตำแหน่งเซลล์ที่มีค่าซ้ำนั้น
     * @throws IOException หากมีข้อผิดพลาดในการอ่านไฟล์
     */
    public static Map<Object, List<String>> checkDuplicateValuesInRange(Workbook workbook , String sheetName, 
                                                                       String rangeStr, Set<Object> exceptValues) throws IOException {
            try {
                // เลือกชีท
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    throw new IllegalArgumentException("ไม่พบชีท '" + sheetName + "' ในไฟล์");
                }
                
                return checkDuplicateValuesInRange(sheet, rangeStr, exceptValues);
            } finally {
                workbook.close();
            }
    }
    
    /**
     * อ่านค่าทั้งหมดในช่วงเซลล์ที่กำหนด โดยสามารถระบุค่าที่ยกเว้นได้
     * 
     * @param fileName ชื่อไฟล์ Excel
     * @param sheetName ชื่อชีทที่ต้องการอ่าน
     * @param rangeStr ช่วงเซลล์ที่ต้องการอ่าน (เช่น "A1:A100" หรือ "B5:D20")
     * @param exceptValues ค่าที่ยกเว้นไม่ต้องอ่าน (สามารถเป็น null ถ้าไม่มีค่ายกเว้น)
     * @return Map ที่มีข้อมูลเกี่ยวกับค่าทั้งหมดที่อ่านได้ โดยมี key เป็นค่าที่อ่านได้ และ value เป็น List ของตำแหน่งเซลล์ที่มีค่านั้น
     * @throws IOException หากมีข้อผิดพลาดในการอ่านไฟล์
     */
    public static List<String> getAllValuesInRange(Workbook workbook, String sheetName, 
                                                              String rangeStr, Set<Object> exceptValues) throws IOException {
       
            try {
                // เลือกชีท
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    throw new IllegalArgumentException("ไม่พบชีท '" + sheetName + "' ในไฟล์");
                }
                
                // แยกช่วงจากสตริง (เช่น "A1:A100")
                String[] rangeParts = rangeStr.split(":");
                if (rangeParts.length != 2) {
                    throw new IllegalArgumentException("รูปแบบช่วงไม่ถูกต้อง: " + rangeStr);
                }
                
                String startCellRef = rangeParts[0].trim();
                String endCellRef = rangeParts[1].trim();
                
                // แยกคอลัมน์และแถวจากเซลล์เริ่มต้น
                Matcher startMatcher = Pattern.compile("([A-Za-z]+)([0-9]+)").matcher(startCellRef);
                if (!startMatcher.matches()) {
                    throw new IllegalArgumentException("รูปแบบเซลล์เริ่มต้นไม่ถูกต้อง: " + startCellRef);
                }
                
                String startColStr = startMatcher.group(1);
                int startRow = Integer.parseInt(startMatcher.group(2)) - 1; // แปลงเป็น 0-based
                CellReference startCellReference = new CellReference(startCellRef);
                int startColIdx = startCellReference.getCol();
                
                int endRow;
                int endColIdx;
                
                if ("EOF".equalsIgnoreCase(endCellRef)) {
                    // หาแถวสุดท้ายที่มีข้อมูลในคอลัมน์ที่กำหนด
                    endRow = sheet.getLastRowNum();
                    endColIdx = startColIdx; // ใช้คอลัมน์เดียวกับจุดเริ่มต้น
                    
                    boolean foundData = false;
                    for (int r = endRow; r >= startRow; r--) {
                        Row row = sheet.getRow(r);
                        if (row != null) {
                            Cell cell = row.getCell(startColIdx);
                            if (cell != null && !isEmpty(cell)) {
                                endRow = r;
                                foundData = true;
                                break;
                            }
                        }
                    }
                    
                    if (!foundData) {
                        return new ArrayList<>(); // ไม่พบข้อมูลเลย ส่งคืน Map ว่าง
                    }
                } else {
                    // ใช้ค่าที่ระบุ
                    Matcher endMatcher = Pattern.compile("([A-Za-z]+)([0-9]+)").matcher(endCellRef);
                    if (!endMatcher.matches()) {
                        throw new IllegalArgumentException("รูปแบบเซลล์สิ้นสุดไม่ถูกต้อง: " + endCellRef);
                    }
                    
                    String endColStr = endMatcher.group(1);
                    endRow = Integer.parseInt(endMatcher.group(2)) - 1; // แปลงเป็น 0-based
                    CellReference endCellReference = new CellReference(endCellRef);
                    endColIdx = endCellReference.getCol();
                }
                
                // เก็บค่าและตำแหน่งของแต่ละค่า
                List<String> valuePositions = new ArrayList<>();
                
                // ตรวจสอบค่าในแต่ละเซลล์ในช่วงที่กำหนด
                for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
                    Row row = sheet.getRow(rowIdx);
                    if (row == null) continue;
                    
                    for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
                        Cell cell = row.getCell(colIdx);
                        if (cell == null || isEmpty(cell)) continue;
                        
                        Object cellValueObject = getCellValue(cell);
                        String cellValue = cellValueObject == null ? "" : cellValueObject.toString().trim();

                        
                        // ข้ามค่าที่อยู่ในรายการยกเว้น
                        if (exceptValues != null && exceptValues.contains(cellValue)) {
                            continue;
                        }
                        // สร้างตำแหน่งเซลล์ (เช่น "A1", "B5")
                        // String cellPosition = CellReference.convertNumToColString(colIdx) + (rowIdx + 1);
                        
                        // เก็บตำแหน่งของแต่ละค่า
                        if (!valuePositions.contains(cellValue)) {
                            valuePositions.add(cellValue);
                        }                        
                        // valuePositions.get(cellValue).add(cellPosition);
                    }
                }
                
                return valuePositions;
            } finally {
                workbook.close();
            }
    }
    
    /**
     * ตรวจสอบค่าซ้ำในช่วงเซลล์ที่กำหนด โดยสามารถระบุค่าที่ยกเว้นไม่ต้องตรวจสอบได้
     * 
     * @param sheet ชีทที่ต้องการตรวจสอบ
     * @param rangeStr ช่วงเซลล์ที่ต้องการตรวจสอบ (เช่น "A1:A100" หรือ "B5:D20")
     * @param exceptValues ค่าที่ยกเว้นไม่ต้องตรวจสอบการซ้ำ (สามารถเป็น null ถ้าไม่มีค่ายกเว้น)
     * @return Map ที่มีข้อมูลเกี่ยวกับค่าซ้ำที่พบ โดยมี key เป็นค่าที่ซ้ำ และ value เป็น List ของตำแหน่งเซลล์ที่มีค่าซ้ำนั้น
     */
    public static Map<Object, List<String>> checkDuplicateValuesInRange(Sheet sheet, String rangeStr, Set<Object> exceptValues) {
        Map<Object, List<String>> duplicates = new HashMap<>();
        Map<Object, List<String>> valuePositions = new HashMap<>();
        
        // แยกช่วงจากสตริง (เช่น "A1:A100")
        String[] rangeParts = rangeStr.split(":");
        if (rangeParts.length != 2) {
            throw new IllegalArgumentException("รูปแบบช่วงไม่ถูกต้อง: " + rangeStr);
        }
        
        String startCellRef = rangeParts[0].trim();
        String endCellRef = rangeParts[1].trim();
        
        // แยกคอลัมน์และแถวจากเซลล์เริ่มต้น
        Matcher startMatcher = Pattern.compile("([A-Za-z]+)([0-9]+)").matcher(startCellRef);
        if (!startMatcher.matches()) {
            throw new IllegalArgumentException("รูปแบบเซลล์เริ่มต้นไม่ถูกต้อง: " + startCellRef);
        }
        
        String startColStr = startMatcher.group(1);
        int startRow = Integer.parseInt(startMatcher.group(2)) - 1; // แปลงเป็น 0-based
        CellReference startCellReference = new CellReference(startCellRef);
        int startColIdx = startCellReference.getCol();
        
        int endRow;
        int endColIdx;
        
        if ("EOF".equalsIgnoreCase(endCellRef)) {
            // หาแถวสุดท้ายที่มีข้อมูลในคอลัมน์ที่กำหนด
            endRow = sheet.getLastRowNum();
            endColIdx = startColIdx; // ใช้คอลัมน์เดียวกับจุดเริ่มต้น
            
            boolean foundData = false;
            for (int r = endRow; r >= startRow; r--) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    Cell cell = row.getCell(startColIdx);
                    if (cell != null && !isEmpty(cell)) {
                        endRow = r;
                        foundData = true;
                        break;
                    }
                }
            }
            
            if (!foundData) {
                return duplicates; // ไม่พบข้อมูลเลย ส่งคืน Map ว่าง
            }
        } else {
            // ใช้ค่าที่ระบุ
            Matcher endMatcher = Pattern.compile("([A-Za-z]+)([0-9]+)").matcher(endCellRef);
            if (!endMatcher.matches()) {
                throw new IllegalArgumentException("รูปแบบเซลล์สิ้นสุดไม่ถูกต้อง: " + endCellRef);
            }
            
            String endColStr = endMatcher.group(1);
            endRow = Integer.parseInt(endMatcher.group(2)) - 1; // แปลงเป็น 0-based
            CellReference endCellReference = new CellReference(endCellRef);
            endColIdx = endCellReference.getCol();
        }
        
        // ตรวจสอบค่าในแต่ละเซลล์ในช่วงที่กำหนด
        for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null) continue;
            
            for (int colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
                Cell cell = row.getCell(colIdx);
                if (cell == null || isEmpty(cell)) continue;
                
                Object cellValue = getCellValue(cell);
                
                // ข้ามค่าที่อยู่ในรายการยกเว้น
                if (exceptValues != null && exceptValues.contains(cellValue)) {
                    continue;
                }
                
                // สร้างตำแหน่งเซลล์ (เช่น "A1", "B5")
                String cellPosition = CellReference.convertNumToColString(colIdx) + (rowIdx + 1);
                
                // เก็บตำแหน่งของแต่ละค่า
                if (!valuePositions.containsKey(cellValue)) {
                    valuePositions.put(cellValue, new ArrayList<>());
                }
                valuePositions.get(cellValue).add(cellPosition);
            }
        }
        
        // ตรวจสอบค่าซ้ำ
        for (Map.Entry<Object, List<String>> entry : valuePositions.entrySet()) {
            if (entry.getValue().size() > 1) {
                // พบค่าซ้ำ
                duplicates.put(entry.getKey(), entry.getValue());
            }
        }
        
        return duplicates;
    }
    
    /**
     * แยกตำแหน่งเซลล์จากคำสั่ง COLUMN
     * เช่น COLUMN("AMOUNT", "C?") จะได้ "C?"
     * 
     * @param columnStr คำสั่ง COLUMN
     * @return ตำแหน่งเซลล์
     */
    private static String extractColumnRef(String columnStr) {
        // ตรวจสอบว่าเป็นคำสั่ง COLUMN หรือไม่
        if (!columnStr.startsWith("COLUMN(")) {
            return columnStr; // ถ้าไม่ใช่ COLUMN ให้ส่งคืนค่าเดิม
        }
        
        // ตัดส่วนหัวและท้ายของคำสั่งออก
        String content = columnStr.substring("COLUMN(".length(), columnStr.length() - 1);
        
        // แยกพารามิเตอร์
        List<String> params = new ArrayList<>();
        StringBuilder currentParam = new StringBuilder();
        boolean inQuotes = false;
        
        for (int i = 0; i < content.length(); i++) {
            char c = content.charAt(i);
            
            if (c == '"') {
                inQuotes = !inQuotes;
                currentParam.append(c);
            } else if (c == ',' && !inQuotes) {
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
        
        // ตรวจสอบว่ามีพารามิเตอร์ครบหรือไม่
        if (params.size() < 2) {
            return null;
        }
        
        // ดึงพารามิเตอร์ที่ 2 (ตำแหน่งเซลล์)
        String cellRef = params.get(1);
        if (cellRef.startsWith("\"") && cellRef.endsWith("\"")) {
            cellRef = cellRef.substring(1, cellRef.length() - 1);
        }
        
        return cellRef;
    }
}
