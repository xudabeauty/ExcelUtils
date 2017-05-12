package com.utils.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.utils.excel.entity.ExcelCell;
import com.utils.excel.entity.ExcelRow;
import com.utils.excel.entity.ExcelSheet;

public class ExcelUtils {
    private static Logger             logger       = LoggerFactory.getLogger(ExcelUtils.class);
    public static final String EMPTY_STRING = "";

    public static SXSSFWorkbook initWork() {
        return new SXSSFWorkbook(100);
    }
    public static void exportMaps(SXSSFWorkbook wb, OutputStream out, List<Map<String, Object>> maps, boolean newSheet, boolean setHeader) {
        Sheet sheet = null;
        int r = 0;
        if (newSheet) {
            sheet = wb.createSheet();
        } else {
            int sheetNum = wb.getNumberOfSheets();
            if (sheetNum == 0) {
                wb.createSheet();
            } else {
                int sheetIndex = wb.getActiveSheetIndex();
                sheet = wb.getSheetAt(sheetIndex);
                r = sheet.getLastRowNum() + 1;
            }
        }
        if (setHeader) {
            Map<String, Object> map = maps.get(0);
            Row row = sheet.createRow(r);
            int i = 0;
            for (String val : map.keySet()) {
                if (val.contains(".")) {
                    val = StringUtils.substringAfterLast(val, ".");
                }
                Cell c = row.createCell(i);
                setValue(wb, c, val);
                i++;
            }
            r++;
            return;
        }
        for (Map<String, Object> map : maps) {
            Row row = sheet.createRow(r);
            int i = 0;
            for (Entry<String, Object> entry : map.entrySet()) {
                Object val = entry.getValue();
                Cell cell = row.createCell(i);
                setValue(wb, cell, val);
                i++;
            }
            r++;
        }
    }
    
    /**
     * 导出map集合到输出流
            * @author: xuhui  
            * @createTime: 2017年5月12日 下午12:12:52  
            * @history:  
            * @param out
            * @param maps
            * @param setHeader void
     */
    public static void exportMaps(OutputStream out,List<Map<String, Object>>maps,boolean setHeader){
        SXSSFWorkbook wb=new SXSSFWorkbook(100);
                Sheet sheet=wb.createSheet();
                int r=0;
                if(setHeader){
                    Map<String, Object> map=maps.get(0);
                    Row row=sheet.createRow(r);
                    int i=0;
                    for(String val:map.keySet()){
                        Cell cell=row.createCell(i);
                        setValue(wb, cell, val);
                        i++;
                    }
                    r++;
                }
                for(Map<String, Object> map:maps){
                   Row row= sheet.createRow(r);
                   int i=0;
                   for(Entry<String, Object> entry:map.entrySet()){
                  Cell cell=  row.createCell(i);
                  Object val=entry.getValue();
                  setValue(wb, cell, val);
                       i++;
                   }
                   r++;
                }
                try {
                    wb.write(out);
                    wb.dispose();
                } catch (IOException e) {
                    e.printStackTrace();
                }
    }
    /**
     * 导出对象集合到输出流
            * @author: xuhui  
            * @createTime: 2017年5月12日 下午1:03:08  
            * @history:  
            * @param out
            * @param headers
            * @param objects
            * @param fields
            * @throws IllegalAccessException
            * @throws IllegalArgumentException
            * @throws InvocationTargetException
            * @throws IOException void
     */
    
    public static void exportObject(OutputStream out,List<String>headers,List<?>objects,List<String>fields) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, IOException{
        SXSSFWorkbook wb=new SXSSFWorkbook();
        Sheet sheet=wb.createSheet();
        int r=0;
        if(CollectionUtils.isNotEmpty(headers)){
            Row row=sheet.createRow(r);
            for (int i = 0; i < headers.size(); i++) {
                String header=headers.get(i);
                row.createCell(i).setCellValue(header);
            }
            r++;
        }
        if(CollectionUtils.isNotEmpty(fields)){
                 for(Object obj:objects){
                   Row row=  sheet.createRow(r);
                     if(isPrimative(obj)){
                         Cell cell=row.createCell(0);
                         setValue(wb, cell, obj);
                     }else{
                         setValue(wb, row, obj, fields);
                     }
                 }
        }
        else {
            for (Object obj : objects) {
                Row row = sheet.createRow(r);
                if (isPrimative(obj)) {
                    Cell c = row.createCell(0);
                    setValue(wb, c, obj);
                } else {
                        setValue(wb, row, obj);
                }
                r++;
            }
        }
        wb.write(out);
        wb.dispose();
    }  
    /**
     * 导出表格到输出流
            * @author: xuhui  
            * @createTime: 2017年5月12日 下午1:07:37  
            * @history:  
            * @param out
            * @param sheets void
     */
    public static void export(OutputStream out, List<ExcelSheet> sheets)  {
        SXSSFWorkbook wb = new SXSSFWorkbook();
        for (int i = 0; i < sheets.size(); i++) {
            ExcelSheet sheet = sheets.get(i);
            Sheet sh = wb.createSheet();
            List<ExcelRow> rows = sheet.getRows();
            for (int j = 0; j < rows.size(); j++) {
                ExcelRow row = rows.get(j);
                Row r = sh.createRow(j);
                List<ExcelCell> cells = row.getCells();
                for (int k = 0; k < cells.size(); k++) {
                    ExcelCell cell = cells.get(k);
                    Object o = cell.getCell();
                    if (isPrimative(o)) {
                        Cell c = r.createCell(k);
                        setValue(wb, c, o);
                    } else {
                        Cell c = r.createCell(k);
                        c.setCellValue(EMPTY_STRING);
                    }
                }
            }
        }
        try {
            wb.write(out);
            out.flush();
            wb.dispose();
        } catch (IOException e) {
            logger.error("ioEception");
        }
    }

    /**
     * 根据不同类型给单元格赋值，Date，Calender ，boolean
            * @author: xuhui  
            * @createTime: 2017年5月12日 上午11:35:58  
            * @history:  
            * @param wb
            * @param cell
            * @param val void
     */
    public static void setValue(SXSSFWorkbook wb, Cell cell, Object val) {
        if (val == null) {
            cell.setCellValue(EMPTY_STRING);
        } else if (val instanceof Date) {
            CreationHelper creterHelper = wb.getCreationHelper();
            CellStyle cellstyle = wb.createCellStyle();
            cellstyle.setDataFormat(creterHelper.createDataFormat().getFormat("yyyy/MM///dd hh:mm:ss"));
            Date vale = (Date) val;
            cell.setCellValue(vale);
            cell.setCellStyle(cellstyle);
        } else if (val instanceof Calendar) {
            CreationHelper createHelper = wb.getCreationHelper();
            CellStyle cellStyle = wb.createCellStyle();
            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/MM/dd hh:mm:ss"));
            Calendar value = (Calendar) val;
            cell.setCellValue(value);
            cell.setCellStyle(cellStyle);
        } else if (val instanceof Boolean) {
            Boolean value = (Boolean) val;
            cell.setCellValue(value);
        }
    }
    public static void setValue(SXSSFWorkbook wb,Row row,Object obj,List<String> feilds) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException{
        Method methods[]=obj.getClass().getMethods();
        for (int i = 0; i < methods.length; i++) {
            Cell cell=row.createCell(i);
            String getter="get"+feilds.get(i);
            String is = "is" + feilds.get(i);
            Object o = EMPTY_STRING;
            for (Method method : methods) {
                String name=method.getName();
                o = method.invoke(obj);
                break;
            }
            setValue(wb, cell, o);
        }
    }
    /**
     * 给行赋值
            * @author: xuhui  
            * @createTime: 2017年5月12日 下午1:01:26  
            * @history:  
            * @param wb
            * @param row
            * @param obj
            * @throws IllegalAccessException
            * @throws IllegalArgumentException
            * @throws InvocationTargetException void
     */
    
    public static void setValue(SXSSFWorkbook wb,Row row,Object obj) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException{
        Method methods[]=obj.getClass().getMethods();
        int i=0;
        for (Method method : methods) {
            String name=method.getName();
            if(name.startsWith("get")||name.startsWith("is")){
               Object object= method.invoke(obj);
           Cell cell=   row.createCell(i);
           setValue(wb, cell, object);
            }
        }
    }
    /**
     * 工作簿需要的基本类型
            * @author: xuhui  
            * @createTime: 2017年5月12日 下午12:31:02  
            * @history:  
            * @param obj
            * @return boolean
     */
    public static boolean isPrimative(Object obj){
        if(obj instanceof Date||obj instanceof Boolean||obj instanceof Double
                ||obj instanceof Integer||obj instanceof Long||obj instanceof String){
            return true;
        }
        return false;
    }
}
