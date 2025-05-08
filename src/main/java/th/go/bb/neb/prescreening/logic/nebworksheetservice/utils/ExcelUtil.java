package th.go.bb.neb.prescreening.logic.nebworksheetservice.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.JSONObject;
import org.json.JSONArray;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelUtil {    

    private Workbook workbook;
    private String filePath;
    private JSONObject jsonData;

    /**
     * สร้าง ExcelToJsonReader ด้วยไฟล์ Excel ที่ระบุ
     * 
     * @param filePath ที่อยู่ไฟล์ Excel
     * @throws IOException หากไม่สามารถอ่านไฟล์ได้
     */
    public ExcelUtil(String filePath) throws IOException {
        this.filePath = filePath;
        this.jsonData = new JSONObject();
        
        FileInputStream inputStream = new FileInputStream(new File(filePath));
        
        // ตรวจสอบนามสกุลไฟล์เพื่อเลือกประเภท Workbook ที่เหมาะสม
        if (filePath.endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (filePath.endsWith(".xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            inputStream.close();
            throw new IOException("รูปแบบไฟล์ไม่รองรับ ต้องเป็น .xlsx หรือ .xls เท่านั้น");
        }
        
        inputStream.close();
    }

    /**
     * อ่านค่าจากเซลล์ที่ระบุ (เช่น A1, B5, etc.)
     * 
     * @param sheetName ชื่อของชีท
     * @param cellReference ตำแหน่งเซลล์ (เช่น A1, B5)
     * @return ค่าในเซลล์เป็น String
     */
    public String readCellValue(String sheetName, String cellReference) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            return null;
        }
        
        // แยกตัวอักษรคอลัมน์และหมายเลขแถว
        String colString = cellReference.replaceAll("[0-9]", "");
        String rowString = cellReference.replaceAll("[A-Za-z]", "");
        
        // แปลงตัวอักษรคอลัมน์เป็นดัชนี (A=0, B=1, ...)
        int colIndex = 0;
        for (char c : colString.toUpperCase().toCharArray()) {
            colIndex = colIndex * 26 + (c - 'A' + 1);
        }
        colIndex--; // ปรับให้เริ่มที่ 0
        
        // แปลงหมายเลขแถวเป็นดัชนี (แถวใน POI เริ่มที่ 0)
        int rowIndex = Integer.parseInt(rowString) - 1;
        
        // อ่านค่าจากเซลล์
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }
        
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            return null;
        }
        
        return getCellValueAsString(cell);
    }
    
    /**
     * ดึงค่าจากเซลล์และแปลงเป็น String โดยอัตโนมัติตามประเภทข้อมูล
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // ป้องกันการแสดงเป็นรูปแบบวิทยาศาสตร์
                    double value = cell.getNumericCellValue();
                    if (value == Math.floor(value)) {
                        return String.format("%.0f", value);
                    } else {
                        return String.valueOf(value);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception ex) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return "";
        }
    }
    
    /**
     * อ่านหลายเซลล์และเก็บในรูปแบบ JSON ตามที่กำหนด
     * 
     * @param sheetName ชื่อของชีท
     * @param cellMappings Map ที่เก็บความสัมพันธ์ระหว่างคีย์ใน JSON และตำแหน่งเซลล์ใน Excel
     * @return JSONObject ที่มีข้อมูลจากเซลล์ตามที่กำหนด
     */
    public JSONObject readCellsToJson(String sheetName, Map<String, String> cellMappings) {
        JSONObject result = new JSONObject();
        
        for (Map.Entry<String, String> entry : cellMappings.entrySet()) {
            String jsonKey = entry.getKey();
            String cellReference = entry.getValue();
            String cellValue = readCellValue(sheetName, cellReference);
            
            if (cellValue != null) {
                result.put(jsonKey, cellValue);
            }
        }
        
        return result;
    }
    
    /**
     * อ่านข้อมูลแบบ range (ช่วงของเซลล์) และเก็บเป็น JSONArray
     * 
     * @param sheetName ชื่อของชีท
     * @param startCell เซลล์เริ่มต้น (เช่น A1)
     * @param endCell เซลล์สิ้นสุด (เช่น C10)
     * @param hasHeader ระบุว่าแถวแรกเป็นส่วนหัวหรือไม่
     * @return JSONArray ที่มีข้อมูลจากช่วงเซลล์
     */
    public JSONArray readRangeToJson(String sheetName, String startCell, String endCell, boolean hasHeader) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            return new JSONArray();
        }
        
        // แยกคอลัมน์และแถวของเซลล์เริ่มต้น
        String startColString = startCell.replaceAll("[0-9]", "");
        int startRow = Integer.parseInt(startCell.replaceAll("[A-Za-z]", "")) - 1;
        
        // แยกคอลัมน์และแถวของเซลล์สิ้นสุด
        String endColString = endCell.replaceAll("[0-9]", "");
        int endRow = Integer.parseInt(endCell.replaceAll("[A-Za-z]", "")) - 1;
        
        // แปลงตัวอักษรคอลัมน์เป็นดัชนี
        int startCol = 0;
        for (char c : startColString.toUpperCase().toCharArray()) {
            startCol = startCol * 26 + (c - 'A' + 1);
        }
        startCol--;
        
        int endCol = 0;
        for (char c : endColString.toUpperCase().toCharArray()) {
            endCol = endCol * 26 + (c - 'A' + 1);
        }
        endCol--;
        
        JSONArray resultArray = new JSONArray();
        String[] headers = null;
        
        // อ่านส่วนหัว (ถ้ามี)
        if (hasHeader) {
            Row headerRow = sheet.getRow(startRow);
            if (headerRow != null) {
                headers = new String[endCol - startCol + 1];
                for (int col = startCol; col <= endCol; col++) {
                    Cell cell = headerRow.getCell(col);
                    headers[col - startCol] = (cell != null) ? getCellValueAsString(cell) : "Column" + (col + 1);
                }
                startRow++; // ข้ามแถวหัว
            }
        }
        
        // อ่านข้อมูล
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;
            
            JSONObject rowData = new JSONObject();
            boolean hasData = false;
            
            for (int colIndex = startCol; colIndex <= endCol; colIndex++) {
                Cell cell = row.getCell(colIndex);
                String value = getCellValueAsString(cell);
                
                if (!value.isEmpty()) {
                    hasData = true;
                }
                
                if (headers != null) {
                    // ใช้ชื่อคอลัมน์จากส่วนหัว
                    rowData.put(headers[colIndex - startCol], value);
                } else {
                    // ใช้ตัวอักษรคอลัมน์เป็นคีย์
                    char colChar = (char) ('A' + colIndex);
                    rowData.put(String.valueOf(colChar), value);
                }
            }
            
            if (hasData) {
                resultArray.put(rowData);
            }
        }
        
        return resultArray;
    }
    
    /**
     * อ่านข้อมูลทั้งชีทและแปลงเป็น JSON
     * 
     * @param sheetName ชื่อของชีท
     * @param hasHeader ระบุว่าแถวแรกเป็นส่วนหัวหรือไม่
     * @return JSONArray ที่มีข้อมูลทั้งชีท
     */
    public JSONArray readSheetToJson(String sheetName, boolean hasHeader) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            return new JSONArray();
        }
        
        JSONArray resultArray = new JSONArray();
        String[] headers = null;
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        
        // หาจำนวนคอลัมน์ทั้งหมด
        int maxCol = 0;
        for (int i = firstRow; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getLastCellNum() > maxCol) {
                maxCol = row.getLastCellNum();
            }
        }
        
        // อ่านส่วนหัว (ถ้ามี)
        if (hasHeader && firstRow <= lastRow) {
            Row headerRow = sheet.getRow(firstRow);
            if (headerRow != null) {
                headers = new String[maxCol];
                for (int col = 0; col < maxCol; col++) {
                    Cell cell = headerRow.getCell(col);
                    headers[col] = (cell != null) ? getCellValueAsString(cell) : "Column" + (col + 1);
                }
                firstRow++; // ข้ามแถวหัว
            }
        }
        
        // อ่านข้อมูล
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;
            
            JSONObject rowData = new JSONObject();
            boolean hasData = false;
            
            for (int colIndex = 0; colIndex < maxCol; colIndex++) {
                Cell cell = row.getCell(colIndex);
                String value = getCellValueAsString(cell);
                
                if (!value.isEmpty()) {
                    hasData = true;
                }
                
                if (headers != null) {
                    // ใช้ชื่อคอลัมน์จากส่วนหัว
                    rowData.put(headers[colIndex], value);
                } else {
                    // ใช้ตัวอักษรคอลัมน์เป็นคีย์
                    char colChar = (char) ('A' + colIndex);
                    rowData.put(String.valueOf(colChar), value);
                }
            }
            
            if (hasData) {
                resultArray.put(rowData);
            }
        }
        
        return resultArray;
    }
    
    /**
     * ตรวจสอบความซ้ำซ้อนของข้อมูลในคอลัมน์ที่กำหนด
     * 
     * @param sheetName ชื่อของชีท
     * @param columnReference ตัวอักษรคอลัมน์ (เช่น A, B, C)
     * @param startRow แถวเริ่มต้น (เริ่มจาก 1)
     * @param endRow แถวสิ้นสุด
     * @return JSONArray ของแถวที่มีค่าซ้ำกัน
     */
    public JSONArray findDuplicates(String sheetName, String columnReference, int startRow, int endRow) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            return new JSONArray();
        }
        
        // แปลงตัวอักษรคอลัมน์เป็นดัชนี
        int colIndex = 0;
        for (char c : columnReference.toUpperCase().toCharArray()) {
            colIndex = colIndex * 26 + (c - 'A' + 1);
        }
        colIndex--;
        
        // เก็บค่าและแถวที่พบ
        Map<String, JSONArray> valueMap = new HashMap<>();
        
        for (int rowIndex = startRow - 1; rowIndex < endRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;
            
            Cell cell = row.getCell(colIndex);
            if (cell == null) continue;
            
            String value = getCellValueAsString(cell);
            if (value.isEmpty()) continue;
            
            if (!valueMap.containsKey(value)) {
                valueMap.put(value, new JSONArray());
            }
            
            JSONObject rowInfo = new JSONObject();
            rowInfo.put("row", rowIndex + 1); // แถวที่ (1-based)
            rowInfo.put("value", value);
            
            // เก็บข้อมูลเพิ่มเติมจากแถวนี้
            for (int i = 0; i < row.getLastCellNum(); i++) {
                if (i == colIndex) continue; // ข้ามคอลัมน์ที่ใช้ตรวจสอบ
                
                Cell otherCell = row.getCell(i);
                if (otherCell != null) {
                    char colChar = (char) ('A' + i);
                    rowInfo.put(String.valueOf(colChar), getCellValueAsString(otherCell));
                }
            }
            
            valueMap.get(value).put(rowInfo);
        }
        
        // สร้าง JSONArray ของรายการที่ซ้ำกัน
        JSONArray duplicates = new JSONArray();
        for (Map.Entry<String, JSONArray> entry : valueMap.entrySet()) {
            if (entry.getValue().length() > 1) { // มีมากกว่า 1 แถว = ซ้ำกัน
                JSONObject dupGroup = new JSONObject();
                dupGroup.put("value", entry.getKey());
                dupGroup.put("rows", entry.getValue());
                duplicates.put(dupGroup);
            }
        }
        
        return duplicates;
    }
    
    /**
     * นับจำนวนรายการในคอลัมน์ที่ตรงตามเงื่อนไข
     * 
     * @param sheetName ชื่อของชีท
     * @param columnReference ตัวอักษรคอลัมน์ (เช่น A, B, C)
     * @param condition ค่าที่ต้องการนับ (null = นับทั้งหมดที่ไม่ว่าง)
     * @return จำนวนรายการที่พบ
     */
    public int countItems(String sheetName, String columnReference, String condition) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            return 0;
        }
        
        // แปลงตัวอักษรคอลัมน์เป็นดัชนี
        int colIndex = 0;
        for (char c : columnReference.toUpperCase().toCharArray()) {
            colIndex = colIndex * 26 + (c - 'A' + 1);
        }
        colIndex--;
        
        int count = 0;
        for (Row row : sheet) {
            Cell cell = row.getCell(colIndex);
            if (cell == null) continue;
            
            String value = getCellValueAsString(cell);
            if (value.isEmpty()) continue;
            
            if (condition == null || value.equals(condition)) {
                count++;
            }
        }
        
        return count;
    }
    
    /**
     * ปิด workbook
     */
    public void close() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }
    
    /**
     * ตัวอย่างการใช้งาน
     */
    public static void main(String[] args) {
        try {
            ExcelUtil reader = new ExcelUtil("example.xlsx");
            
            // ตัวอย่างที่ 1: อ่านค่าจากเซลล์เฉพาะ
            String cellValue = reader.readCellValue("Sheet1", "A5");
            System.out.println("ค่าในเซลล์ A5: " + cellValue);
            
            // ตัวอย่างที่ 2: อ่านหลายเซลล์เป็น JSON
            Map<String, String> cellMappings = new HashMap<>();
            cellMappings.put("employeeId", "A5");
            cellMappings.put("name", "B5");
            cellMappings.put("position", "C5");
            cellMappings.put("salary", "D5");
            
            JSONObject employee = reader.readCellsToJson("Sheet1", cellMappings);
            System.out.println("ข้อมูลพนักงาน: " + employee.toString(2));
            
            // ตัวอย่างที่ 3: อ่านช่วงเซลล์เป็น JSON Array
            JSONArray employees = reader.readRangeToJson("Sheet1", "A1", "D10", true);
            System.out.println("รายการพนักงาน: " + employees.toString(2));
            
            // ตัวอย่างที่ 4: ตรวจสอบข้อมูลซ้ำ
            JSONArray duplicates = reader.findDuplicates("Sheet1", "A", 1, 20);
            System.out.println("รายการซ้ำ: " + duplicates.toString(2));
            
            // ตัวอย่างที่ 5: นับจำนวนรายการ
            int count = reader.countItems("Sheet1", "C", "Manager");
            System.out.println("จำนวนตำแหน่ง Manager: " + count);
            
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}