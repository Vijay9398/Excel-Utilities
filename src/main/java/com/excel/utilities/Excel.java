package com.excel.utilities;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;

public class Excel {

    private String BLANK = "";

    /**
     *  to get workbook object
     * @param path
     * @return
     */
    public Workbook getBook(String path){
        Workbook book = null;
        try{
            InputStream stream = new FileInputStream(path);
            if(path.endsWith(".xlsx")){
                book = new XSSFWorkbook(stream);
            }else if(path.endsWith(".xls")){
                book = new HSSFWorkbook(stream);
            }else{
                throw new Exception("given file is not compatible");
            }
        }catch(Exception e){
            e.printStackTrace();
        }
        return book;
    }

    /**
     *  converts all cell data into string cell value
     * @param cell
     * @return
     */
    public String getCellData(Cell cell){
        String value = null;
        switch(cell.getCellTypeEnum()) {
            case STRING:
                value = cell.getStringCellValue();
                break;
            case BLANK:
                value = BLANK;
                break;
            case BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case NUMERIC:
                cell.setCellType(CellType.STRING);
                value = cell.getStringCellValue();
                break;
            case _NONE:
            default:
                int row_index = cell.getRowIndex();
                int col_index = cell.getColumnIndex();
                System.err.println("cell at location("+row_index+","+col_index+") in the sheet : "+cell.getSheet().getSheetName()+" contains unreadable, passing null value");
        }
        return value;
    }

    /**
     *  returns sheet data in row-wise in the form of List
     * @param sheet
     * @return
     */
    public List<List<String>> getRows(Sheet sheet){
        Iterator<Row> rows = sheet.rowIterator();
        List<List<String>> rowList = new ArrayList<>();
        while(rows.hasNext()){
            Row row = rows.next();
            List<String> valueList = getColumns(row);
            rowList.add(valueList);
        }
        return rowList;
    }

    /**
     *  returns sheet data in row-wise in the form of Map
     * @param sheet
     * @return
     */
    public List<Map<String,String>> getRowMaps(Sheet sheet){
        List<List<String>> rowsData = getRows(sheet);
        List<String> header = rowsData.get(0);
        List<Map<String,String>> dataMaps = new ArrayList<>();
        for(int i =1;i<rowsData.size();i++){
            Map<String,String> dataMap = new LinkedHashMap<>();
            List<String> row = rowsData.get(i);
            if(header.size() == rowsData.get(i).size()){
                for(int j = 0;j<header.size();j++){
                    String key = header.get(j);
                    String value = row.get(j);
                    dataMap.put(key,value);
                }
                dataMaps.add(dataMap);
            }else{
                System.out.println("mismatch in size , ignoring "+i+"row");
            }
        }
        return dataMaps;
    }

    /**
     *  returns list of column data for a particular row
     * @param row
     * @return
     */
    public List<String> getColumns(Row row){
        Iterator<Cell> cells = row.cellIterator();
        List<String> valueList = new ArrayList<>();
        while(cells.hasNext()){
            Cell cell = cells.next();
            String value = getCellData(cell);
            if(value == null){
                try{
                    int row_index = cell.getRowIndex();
                    int col_index = cell.getColumnIndex();
                    throw new NullPointerException("cell at location("+row_index+","+col_index+") in the sheet : "+cell.getSheet().getSheetName()+" contains unreadable");
                }catch(Exception e){ e.printStackTrace();}
            }
            valueList.add(value);
        }
        return valueList;
    }


    /**
     *  returns two dimensional objects of maps for data provider purpose
     * @param path
     * @param sheetName
     * @return
     */
    public Object[][] getMapObjects(String path,String sheetName){
        Workbook book = getBook(path);
        List<Map<String,String>> rowsData = getRowMaps(book.getSheet(sheetName));
        Object[][] dataObject = new Object[rowsData.size()][1];
        for(int i =0;i<rowsData.size();i++){
            dataObject[i][0] = rowsData.get(i);
        }
        return dataObject;
    }

    /**
     *  returns 2 dimensional objects of lists of data for data provider purpose
     * @param path
     * @param sheetName
     * @return
     */
    public Object[][] getListObjects(String path,String sheetName){
        Workbook book = getBook(path);
        List<List<String>> rowsData = getRows(book.getSheet(sheetName));
        Object[][] dataObject = new Object[rowsData.size()][1];
        for(int i =0;i<rowsData.size();i++){
            dataObject[i][0] = rowsData.get(i);
        }
        return dataObject;
    }

    /**
     *  return data map based on a key search
     * @param file
     * @param sheetName
     * @param key
     * @return
     */
    public Map<String,String> getDataMap(String file,String sheetName,String key){
        Workbook book = getBook(file);
        List<Map<String,String>> dataMaps = getRowMaps(book.getSheet(sheetName));
        Map<String,String> result = new LinkedHashMap<>();
        for(Map<String,String> map: dataMaps){
            if(map.get("header").equals(key)){
                result = map;
                break;
            }
        }
        if(result.isEmpty()){
            System.err.println("unable to find the data row , containing headers as "+key+" so returning empty map");
        }
        return result;
    }


    public static void main(String[] args) {
        Excel excel = new Excel();
        Workbook book = excel.getBook("src/main/resources/Book1.xlsx");
        int sheets = book.getNumberOfSheets();
        for(int i = 0; i<sheets;i++){
            Sheet sheet = book.getSheetAt(i);
            System.out.println("===================================");
            System.out.println("printing data from sheet : "+sheet.getSheetName());
            List<Map<String,String>> maps = excel.getRowMaps(sheet);
            System.out.println(maps);
        }
    }


}
