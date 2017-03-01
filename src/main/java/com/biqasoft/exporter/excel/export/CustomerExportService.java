/*
 * Copyright 2016 the original author or authors.
 */

package com.biqasoft.exporter.excel.export;

import com.biqasoft.entity.customer.Customer;
import com.biqasoft.entity.dto.export.excel.ExportCustomersDTO;
import com.biqasoft.entity.filters.CustomerFilter;
import com.biqasoft.exporter.excel.common.excel.ExcelBasicObjectPrinter;
import com.biqasoft.exporter.excel.common.excel.ExcelHelperService;
import com.biqasoft.exporter.excel.dataobject.LocalizedExcelHeader;
import com.biqasoft.exporter.excel.dataobject.SheetData;
import com.biqasoft.microservice.i18n.MessageByLocaleService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * @author Nikita Bakaev, ya@nbakaev.ru
 *         Date: 7/11/2016
 *         All Rights Reserved
 */
@Service
public class CustomerExportService {

    private final ExcelHelperService excelHelperService;
    private final MessageByLocaleService messageByLocaleService;

    @Autowired
    public CustomerExportService(ExcelHelperService excelHelperService, MessageByLocaleService messageByLocaleService) {
        this.excelHelperService = excelHelperService;
        this.messageByLocaleService = messageByLocaleService;
    }

    public byte[] getCustomerInEXCEL(ExportCustomersDTO dto) {
        CustomerFilter customerBuilder = dto.getCustomerFilter();
        List<Customer> customers = dto.getResultedObjects();

        ExcelBasicObjectPrinter basicExcelObjectPrinter = new ExcelBasicObjectPrinter(messageByLocaleService.getMessage("excel.export.sheet.main"), excelHelperService);
        basicExcelObjectPrinter.createMetaSheet(customerBuilder, dto.getEntityNumber(), new Date());
        basicExcelObjectPrinter.setDateFormat(dto.getDateFormat());

        SheetData sheetData = basicExcelObjectPrinter.getMainSheet();
        XSSFWorkbook workbook = basicExcelObjectPrinter.getWorkbook();
        List<Object[]> data = sheetData.getData();

        int i = 0;
        sheetData.setHeaders(new LocalizedExcelHeader[]{
                LocalizedExcelHeader.of("№", false, false),
                LocalizedExcelHeader.of("ID", false, false),
                LocalizedExcelHeader.of("excel.export.customer.header.firstname"),
                LocalizedExcelHeader.of("excel.export.customer.header.lastname"),
                LocalizedExcelHeader.of("excel.export.customer.header.patronymic"),
                LocalizedExcelHeader.of("excel.export.customer.header.email"),
                LocalizedExcelHeader.of("excel.export.customer.header.note"),
                LocalizedExcelHeader.of("excel.export.customer.header.telephone"),
                LocalizedExcelHeader.of("excel.export.customer.header.lead"),
                LocalizedExcelHeader.of("excel.export.customer.header.customer"),
                LocalizedExcelHeader.of("excel.export.customer.header.active"),
                LocalizedExcelHeader.of("excel.export.customer.header.important"),
                LocalizedExcelHeader.of("excel.export.customer.header.b2b"),
                LocalizedExcelHeader.of("excel.export.customer.header.sex"),
                LocalizedExcelHeader.of("excel.export.customer.header.address"),
                LocalizedExcelHeader.of("excel.export.customer.header.position"),
                LocalizedExcelHeader.of("excel.export.customer.header.lead_gen_method"),
                LocalizedExcelHeader.of("excel.export.customer.header.lead_gen_project"),
                LocalizedExcelHeader.of("excel.export.customer.header.sales_funnel_status_id"),
                LocalizedExcelHeader.of("excel.export.customer.header.company_id"),
                LocalizedExcelHeader.of("excel.export.customer.header.responsible_manager_id"),
                LocalizedExcelHeader.of("excel.export.customer.header.task_count"),
                LocalizedExcelHeader.of("excel.export.customer.header.deals_count"),
                LocalizedExcelHeader.of("excel.export.customer.header.deals_amount"),
                LocalizedExcelHeader.of("excel.export.customer.header.opportunity_count"),
                LocalizedExcelHeader.of("excel.export.customer.header.opportunity_amount"),
        });

        for (Customer customer : customers) {
            i++;
            data.add(new Object[]{String.valueOf(i),
                    customer.getId(),
                    customer.getFirstName(),
                    customer.getLastName(),
                    customer.getPatronymic(),
                    customer.getEmail(),
                    customer.getDescription(),
                    customer.getTelephone(),
                    customer.isLead() ? messageByLocaleService.getMessage("common.true") : messageByLocaleService.getMessage("common.false"),
                    customer.isCustomer() ? messageByLocaleService.getMessage("common.true") : messageByLocaleService.getMessage("common.false"),
                    customer.isActive() ? messageByLocaleService.getMessage("common.true") : messageByLocaleService.getMessage("common.false"),
                    customer.isImportant() ? messageByLocaleService.getMessage("common.true") : messageByLocaleService.getMessage("common.false"),
                    customer.isB2b() ? messageByLocaleService.getMessage("common.true") : messageByLocaleService.getMessage("common.false"),
                    customer.getSex(),
                    customer.getAddress(),
                    customer.getPosition(),
                    customer.getLeadGenMethod(),
                    customer.getLeadGenProject(),
                    (customer.getSalesFunnelStatus() != null && customer.getSalesFunnelStatus().getId() != null) ? customer.getSalesFunnelStatus().getId() : messageByLocaleService.getMessage("common.error"),
                    (customer.isB2b() && customer.getCompany() != null && customer.getCompany().getId() != null) ? customer.getCompany().getId() : (customer.isB2b() ? messageByLocaleService.getMessage("common.error") : ""),
                    customer.getResponsibleManagerID() != null && !customer.getResponsibleManagerID().equals("") ? customer.getResponsibleManagerID() : messageByLocaleService.getMessage("common.error"),
                    customer.getCustomerOverview().getActiveTaskNumber(),
                    customer.getCustomerOverview().getDealsNumber(),
                    customer.getCustomerOverview().getDealsAmount(),
                    customer.getCustomerOverview().getOpportunityNumber(),
                    customer.getCustomerOverview().getOpportunityAmount(),
            });
        }

        int indexOfTaskCell = Arrays.asList(sheetData.getHeaders()).indexOf(LocalizedExcelHeader.of("excel.export.customer.header.task_count"));
        int indexOfActiveCell = Arrays.asList(sheetData.getHeaders()).indexOf(LocalizedExcelHeader.of("excel.export.customer.header.active"));
        int indexOfDealsNumberCell = Arrays.asList(sheetData.getHeaders()).indexOf(LocalizedExcelHeader.of("excel.export.customer.header.deals_count"));
        int indexOfDealsAmountCell = Arrays.asList(sheetData.getHeaders()).indexOf(LocalizedExcelHeader.of("excel.export.customer.header.deals_amount"));
        int indexOfOpportunityNumberCell = Arrays.asList(sheetData.getHeaders()).indexOf(LocalizedExcelHeader.of("excel.export.customer.header.opportunity_count"));
        int indexOfResponsibleManagerCell = Arrays.asList(sheetData.getHeaders()).indexOf(LocalizedExcelHeader.of("excel.export.customer.header.responsible_manager_id"));
        int indexOfImportantCell = Arrays.asList(sheetData.getHeaders()).indexOf(LocalizedExcelHeader.of("excel.export.customer.header.important"));

        basicExcelObjectPrinter.setFunction(x -> {

            int rownum = x.getRowNum();
            int cellnum = x.getColNum();
            Object obj = x.getObject();
            Cell cell = x.getCell();

            // this is body
            if (rownum > 1) {

                if (cellnum == indexOfTaskCell + 1) {
                    if (obj instanceof Long) {
                        long numberTask = (Long) obj;
                        if (numberTask < 1) {
                            excelHelperService.fillCellWithErrorColor(cell, workbook);
                        }
                    }
                } else if (cellnum == indexOfDealsNumberCell + 1) {
                    if (obj instanceof Long) {
                        long numberTask = (Long) obj;
                        if (numberTask < 1) {
                            excelHelperService.fillCellWithErrorColor(cell, workbook);
                        }
                    }
                } else if (cellnum == indexOfOpportunityNumberCell + 1) {
                    if (obj instanceof Long) {
                        long numberTask = (Long) obj;
                        if (numberTask < 1) {
                            excelHelperService.fillCellWithErrorColor(cell, workbook);
                        }
                    }
                } else if (cellnum == indexOfDealsAmountCell + 1) {
                    if (obj instanceof BigDecimal) {
                        double numberTask = ((BigDecimal) obj).doubleValue();
                        if (numberTask < 1) {
                            excelHelperService.fillCellWithErrorColor(cell, workbook);
                        }
                    }
                } else if (cellnum == indexOfImportantCell + 1) {
                    if (obj instanceof String) {
                        if (obj.equals(messageByLocaleService.getMessage("common.true"))) {
                            excelHelperService.fillCellWithSuccessColor(cell, workbook);
                        }
                    }
                } else if (cellnum == indexOfActiveCell + 1) {
                    if (obj instanceof String) {
                        if (obj.equals(messageByLocaleService.getMessage("common.true"))) {
                            excelHelperService.fillCellWithSuccessColor(cell, workbook);
                        }
                    }
                } else if (cellnum == indexOfResponsibleManagerCell + 1) {
                    if (obj instanceof String) {
                        if (obj.equals(messageByLocaleService.getMessage("common.error"))) {
                            excelHelperService.fillCellWithErrorColor(cell, workbook);
                        }
                    }
                }
            }

            return null;
        });

        basicExcelObjectPrinter.build();
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        try {
            //Write the workbook in file system
//            FileOutputStream out = new FileOutputStream(new File("howtodoinjava_demo.xlsx"));
            workbook.write(outputStream);
//            workbook.
//            out.close();
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage());
        }
        return outputStream.toByteArray();
    }

}
