/*
 * Copyright (c) 2016. com.biqasoft
 */

package com.biqasoft.exporter.excel.export;

import com.biqasoft.datasource.kpi.DataSourceValueResolverService;
import com.biqasoft.entity.constants.DATA_SOURCES_MORE;
import com.biqasoft.entity.datasources.DataSource;
import com.biqasoft.entity.datasources.light.Lights;
import com.biqasoft.entity.dto.export.excel.ExportKPIDTO;
import com.biqasoft.exporter.excel.common.excel.ExcelBasicObjectPrinter;
import com.biqasoft.exporter.excel.common.excel.ExcelHelperService;
import com.biqasoft.exporter.excel.dataobject.CellProcessingInterceptorReturnConfig;
import com.biqasoft.exporter.excel.dataobject.LocalizedExcelHeader;
import com.biqasoft.exporter.excel.dataobject.SheetData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import java.io.ByteArrayOutputStream;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * Write KPIs to excel
 */
@Service
public class KPIsExcelService {

    private final ExcelHelperService excelHelperService;
    private final DataSourceValueResolverService dataSourceValueResolverService;

    @Autowired
    public KPIsExcelService(ExcelHelperService excelHelperService, DataSourceValueResolverService dataSourceValueResolverService) {
        this.excelHelperService = excelHelperService;
        this.dataSourceValueResolverService = dataSourceValueResolverService;
    }

    public byte[] getKPisInEXCEL(ExportKPIDTO kpidto) {

        ExcelBasicObjectPrinter basicExcelObjectPrinter = new ExcelBasicObjectPrinter("KPIs", excelHelperService);
        basicExcelObjectPrinter.createMetaSheet(kpidto.getDataSourceFilter(), Integer.toUnsignedLong(kpidto.getResultedObjects().size()), new Date());
        basicExcelObjectPrinter.setDateFormat(kpidto.getDateFormat());

        SheetData sheetData = basicExcelObjectPrinter.getMainSheet();

        XSSFWorkbook workbook = basicExcelObjectPrinter.getWorkbook();

        List<Object[]> data = sheetData.getData();

        int i = 0;

        sheetData.setHeaders(new LocalizedExcelHeader[]{
                LocalizedExcelHeader.of("№", false, false),
                LocalizedExcelHeader.of("ID", false, false),
                LocalizedExcelHeader.of("NAME", false, false),
                LocalizedExcelHeader.of("VALUE", false, false),
                LocalizedExcelHeader.of("LIGHTS", true, false),
                LocalizedExcelHeader.of("RESOLVED", true, false),
                LocalizedExcelHeader.of("SYSTEM_ISSUED", false, false),
                LocalizedExcelHeader.of("RETURN_TYPE", true, false),
                LocalizedExcelHeader.of("LAST_UPDATE", false, false),
                LocalizedExcelHeader.of("ERROR_RESOLVED_MESSAGE", false, false),
                LocalizedExcelHeader.of("CONTROLLED_CLASS", false, false),
        });

        for (DataSource dataSource : kpidto.getResultedObjects()) {
            i++;
            data.add(new Object[]{String.valueOf(i),
                    dataSource.getId(),
                    dataSource.getName(),
                    dataSourceValueResolverService.getResolvedDataSource(dataSource),
                    dataSource,
                    dataSource.isResolved(),
                    dataSource.isSystemIssued(),
                    dataSource.getReturnType(),
                    dataSource.getLastUpdate(),
                    dataSource.getErrorResolvedMessage(),
                    dataSource.getControlledClass(),
            });
        }

        int indexOfLightsCell = Arrays.asList(sheetData.getHeaders()).indexOf(LocalizedExcelHeader.of("LIGHTS", true, false));
        int indexOfValueCell = Arrays.asList(sheetData.getHeaders()).indexOf(LocalizedExcelHeader.of("VALUE", false, false));
        int indexOfResolvedCell = Arrays.asList(sheetData.getHeaders()).indexOf(LocalizedExcelHeader.of("RESOLVED", true, false));

        basicExcelObjectPrinter.setFunction(x -> {

            int rownum = x.getRowNum();
            int cellnum = x.getColNum();
            Object obj = x.getObject();
            Cell cell = x.getCell();
            Row row = x.getRow();

            CellProcessingInterceptorReturnConfig cellProcessingInterceptorReturnConfig = new CellProcessingInterceptorReturnConfig();

            // this is body
            if (rownum > 1) {

                if (cellnum == indexOfResolvedCell + 1) {

                    if (obj instanceof Boolean) {
                        if ((Boolean) obj) {
//                            excelCommonService.fillCellWithSuccessColor(cell, workbook);
                        } else {
                            excelHelperService.fillCellWithErrorColor(cell, workbook);
                        }
                    }
                }

                if (cellnum == indexOfLightsCell + 1) {
                    cellProcessingInterceptorReturnConfig.setSkipValueWrite(true);

                    if (obj instanceof DataSource) {
                        DataSource dataSource = (DataSource) obj;
                        Lights lights = dataSource.getLights();

                        // this is new kpi - we do not resolve it yet
                        if (dataSource.getLastUpdate() == null && !dataSource.isResolved()) {
                            cell.setCellValue("WAIT_PLEASE");
                            excelHelperService.fillCellWithGreyColor(cell, workbook);
                            excelHelperService.fillCellWithGreyColor(row.getCell(indexOfValueCell), workbook);
                        } else {

                            // if we processed it with error
                            if (!dataSource.isResolved()) {
                                cell.setCellValue("ERROR");
                                excelHelperService.fillCellWithErrorColor(cell, workbook);
                                excelHelperService.fillCellWithErrorColor(row.getCell(indexOfValueCell), workbook);
                            }

                            // this is successes resolved but we do not choose color
                            if (dataSource.isResolved() && lights.getCurrentLight() == null) {
                                cell.setCellValue("SUCCESS");
                                excelHelperService.fillCellWithSuccessColor(row.getCell(indexOfValueCell), workbook);
                            } else {

                                if (lights.getCurrentLight() != null) {
                                    switch (lights.getCurrentLight()) {
                                        case DATA_SOURCES_MORE.ERROR:
                                            cell.setCellValue("ERROR");
                                            if (lights.getError() != null && !StringUtils.isEmpty(lights.getError().getColor())) {
                                                excelHelperService.fillCellWithCustomColor(lights.getError().getColor(), row.getCell(indexOfValueCell), workbook);
                                            } else {
                                                excelHelperService.fillCellWithErrorColor(row.getCell(indexOfValueCell), workbook);
                                            }

                                            break;

                                        case DATA_SOURCES_MORE.SUCCESS:
                                            cell.setCellValue("SUCCESS");
                                            if (lights.getSuccess() != null && !StringUtils.isEmpty(lights.getSuccess().getColor())) {
                                                excelHelperService.fillCellWithCustomColor(lights.getSuccess().getColor(), row.getCell(indexOfValueCell), workbook);
                                            } else {
                                                excelHelperService.fillCellWithSuccessColor(row.getCell(indexOfValueCell), workbook);
                                            }

                                            break;

                                        case DATA_SOURCES_MORE.WARNING:
                                            cell.setCellValue("WARNING");
                                            if (lights.getWarning() != null && !StringUtils.isEmpty(lights.getWarning().getColor())) {
                                                excelHelperService.fillCellWithCustomColor(lights.getWarning().getColor(), row.getCell(indexOfValueCell), workbook);
                                            } else {
                                                excelHelperService.fillCellWithWarningColor(row.getCell(indexOfValueCell), workbook);
                                            }

                                            break;
                                    }
                                }
                            }
                        }
                    }
                }

            }
            return cellProcessingInterceptorReturnConfig;
        });

        basicExcelObjectPrinter.build();

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        try {
            workbook.write(outputStream);
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage());
        }
        return outputStream.toByteArray();
    }


}
