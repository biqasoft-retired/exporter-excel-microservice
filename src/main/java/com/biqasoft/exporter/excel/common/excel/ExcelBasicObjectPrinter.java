/*
 * Copyright (c) 2016. com.biqasoft
 */

package com.biqasoft.exporter.excel.common.excel;

import com.biqasoft.entity.core.BaseClassFilter;
import com.biqasoft.exporter.excel.dataobject.CellProcessingInterceptor;
import com.biqasoft.exporter.excel.dataobject.CellProcessingInterceptorReturnConfig;
import com.biqasoft.exporter.excel.dataobject.LocalizedExcelHeader;
import com.biqasoft.exporter.excel.dataobject.SheetData;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.function.Function;

/**
 * Created by Nikita Bakaev, ya@nbakaev.ru on 3/21/2016.
 * All Rights Reserved
 */
public class ExcelBasicObjectPrinter {

    private XSSFWorkbook workbook;
    private ExcelHelperService excelHelperService;
    private List<Object[]> metaData = new ArrayList<>();
    private List<SheetData> sheets = new ArrayList<>();
    private SheetData mainSheet;

    private String dateFormat = "m/d/yy h:mm";

    // this is function which will be called on creating every cell
    private Function<CellProcessingInterceptor, CellProcessingInterceptorReturnConfig> function = null;

    public ExcelBasicObjectPrinter(String mainSheetName, ExcelHelperService excelHelperService) {
        this.excelHelperService = excelHelperService;
        this.workbook = new XSSFWorkbook();

        // Create main a blank sheet with data
        this.mainSheet = createSheet(mainSheetName);
    }

    public SheetData createSheet(String name) {
        SheetData sheetData = new SheetData();
        sheetData.setName(name);
        sheetData.setSheet(workbook.createSheet(name));

        sheets.add(sheetData);
        return sheetData;
    }

    public <T extends BaseClassFilter> XSSFSheet createMetaSheet(T t, Long requestSize, Date date) {

        SheetData sheetDate = createSheet("Meta info");
        XSSFSheet sheet = sheetDate.getSheet();
        ObjectMapper mapper = new ObjectMapper();

        try {
            metaData.add(new Object[]{"//", "Builder object - you can pass this directly to API  ' ", (t == null ? "NULL" : mapper.writeValueAsString(t))});
            metaData.add(new Object[]{"//", "Date", date.toString()});
            metaData.add(new Object[]{"//", "Objects", String.valueOf(requestSize)});
            metaData.add(new Object[]{"//", "---------------------------"});
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e.getMessage());
        }

        sheetDate.setData(metaData);

        return sheet;
    }

    /**
     * Date formatting for excel
     *
     * @return
     */
    private CellStyle setDateFormat() {
        CellStyle dateCellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(dateFormat));
        return dateCellStyle;
    }

    /**
     * Call this method at the end
     * before you want to export data
     */
    public void build() {
        CellStyle dateCellStyle = setDateFormat();

        for (SheetData sheetData : sheets) {

            List<Object[]> data = sheetData.getData();
            XSSFSheet sheet = sheetData.getSheet();
            LocalizedExcelHeader[] headers = sheetData.getHeaders();

            if (headers != null) {
                data.add(0, headers);
            }

            sheet.setTabColor(IndexedColors.GREEN.getIndex());
//            sheet.setTabColor(new XSSFColor ( new Color(0xB6D7A8)  ).getIndexed() );

            int rownum = 0;
            for (Object[] key : data) {
                Row row = sheet.createRow(rownum++);
                Object[] objArr = key;
                int cellnum = 0;
                for (Object obj : objArr) {
                    Cell cell = row.createCell(cellnum++);

                    CellProcessingInterceptorReturnConfig cellProcessingInterceptorReturnConfig = null;

                    if (function != null) {
                        CellProcessingInterceptor cellProcessingInterceptor = new CellProcessingInterceptor();
                        cellProcessingInterceptor.setCell(cell);
                        cellProcessingInterceptor.setColNum(cellnum);
                        cellProcessingInterceptor.setRowNum(rownum);
                        cellProcessingInterceptor.setObject(obj);
                        cellProcessingInterceptor.setRow(row);
                        cellProcessingInterceptorReturnConfig = function.apply(cellProcessingInterceptor);
                    }

                    // set style for the header row
                    if ((rownum == 1 || rownum > 1 && cellnum == 1) && headers != null) {
                        excelHelperService.fillCellWithHeaderColor(cell, workbook);
                        sheet.createFreezePane(0, 1); // this will freeze first five rows
                        if (obj instanceof LocalizedExcelHeader){
                            LocalizedExcelHeader header = (LocalizedExcelHeader) obj;

                            if (header.isLocalized()){
                                header.setName(LocalizedSpringContextAware.getMessageByLocaleService().getMessage(header.getName()));
                            }
                        }
                    }

                    if (cellProcessingInterceptorReturnConfig != null && cellProcessingInterceptorReturnConfig.isSkipValueWrite()) {
                        continue;
                    }

                    if (obj instanceof LocalizedExcelHeader)
                        cell.setCellValue(((LocalizedExcelHeader) obj).getName());
                    if (obj instanceof String)
                        cell.setCellValue((String) obj);
                    else if (obj instanceof Integer)
                        cell.setCellValue((Integer) obj);
                    else if (obj instanceof Double)
                        cell.setCellValue((Double) obj);
                    else if (obj instanceof Long)
                        cell.setCellValue((Long) obj);
                    else if (obj instanceof Date) {
                        cell.setCellStyle(dateCellStyle);
                        cell.setCellValue((Date) obj);
                    } else if (obj instanceof Boolean) {
                        cell.setCellValue((Boolean) obj);
                    } else if (obj instanceof BigDecimal) {
                        cell.setCellValue(((BigDecimal) obj).doubleValue());
                    }


                }
            }

//        ----------------------------------------------

            if (headers != null) {
                // set auto width for every column
                for (int m = 0; m < headers.length; m++) {
                    try {
                        sheet.autoSizeColumn(m);
                        // TODO: use auto filter https://trello.com/c/aBJ3ToFA/349-excel-export-fitler
//                        sheet.setAutoFilter(new CellRangeAddress(1,1,m,m));
                    } catch (Exception e) {
                        throw new RuntimeException(e.getMessage());
                    }
                }
            }


        }

    }


    public SheetData getMainSheet() {
        return mainSheet;
    }

    public void setMainSheet(SheetData mainSheet) {
        this.mainSheet = mainSheet;
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public ExcelHelperService getExcelHelperService() {
        return excelHelperService;
    }

    public void setExcelHelperService(ExcelHelperService excelHelperService) {
        this.excelHelperService = excelHelperService;
    }

    public List<Object[]> getMetaData() {
        return metaData;
    }

    public void setMetaData(List<Object[]> metaData) {
        this.metaData = metaData;
    }

    public List<SheetData> getSheets() {
        return sheets;
    }

    public void setSheets(List<SheetData> sheets) {
        this.sheets = sheets;
    }

    public Function<CellProcessingInterceptor, CellProcessingInterceptorReturnConfig> getFunction() {
        return function;
    }

    public void setFunction(Function<CellProcessingInterceptor, CellProcessingInterceptorReturnConfig> function) {
        this.function = function;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }
}
