/*
 * Copyright 2016 the original author or authors.
 */

package com.biqasoft.exporter.excel.export;

import com.biqasoft.entity.customer.LeadGenMethod;
import com.biqasoft.entity.customer.LeadGenProject;
import com.biqasoft.entity.dto.export.excel.ExportLeadGenMethodDTO;
import com.biqasoft.entity.dto.export.excel.ExportLeadGenMethodWithProjects;
import com.biqasoft.entity.filters.LeadGenMethodExcelFilter;
import com.biqasoft.entity.filters.LeadGenProjectFilter;
import com.biqasoft.gateway.indicators.dto.IndicatorsDTO;
import com.biqasoft.exporter.excel.common.excel.ExcelBasicObjectPrinter;
import com.biqasoft.exporter.excel.common.excel.ExcelHelperService;
import com.biqasoft.exporter.excel.dataobject.LocalizedExcelHeader;
import com.biqasoft.exporter.excel.dataobject.SheetData;
import com.biqasoft.microservice.i18n.MessageByLocaleService;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author Nikita Bakaev, ya@nbakaev.ru
 *         Date: 7/11/2016
 *         All Rights Reserved
 */
@Service
public class LeadGenMethodExcelReportService {

    private final ExcelHelperService excelHelperService;
    private final MessageByLocaleService messageByLocaleService;

    @Autowired
    public LeadGenMethodExcelReportService(ExcelHelperService excelHelperService, MessageByLocaleService messageByLocaleService) {
        this.excelHelperService = excelHelperService;
        this.messageByLocaleService = messageByLocaleService;
    }

    /**
     * Process user request for downloading leadGenMethods statistics
     *
     * @param exportLeadGenMethodDTO
     * @return
     */
    public byte[] getResponseEntity(ExportLeadGenMethodDTO exportLeadGenMethodDTO) {
        LeadGenMethodExcelFilter leadGenMethodBuilder = exportLeadGenMethodDTO.getLeadGenMethodBuilder();

        ExcelBasicObjectPrinter basicExcelObjectPrinter = new ExcelBasicObjectPrinter(messageByLocaleService.getMessage("excel.export.leadgenmethod.sheet"), excelHelperService);
        XSSFWorkbook workbook = basicExcelObjectPrinter.getWorkbook();
        basicExcelObjectPrinter.setDateFormat(exportLeadGenMethodDTO.getDateFormat());

        SheetData mainSheet = basicExcelObjectPrinter.getMainSheet();
        List<Object[]> data = mainSheet.getData();
        basicExcelObjectPrinter.createMetaSheet(exportLeadGenMethodDTO.getLeadGenMethodBuilder(), exportLeadGenMethodDTO.getEntityNumber(), new Date());

        mainSheet.setHeaders(new LocalizedExcelHeader[]{
                LocalizedExcelHeader.of("№", false, false),
                LocalizedExcelHeader.of("ID", false, false),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.name"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.active"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.ROI"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.deals_amount"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.costs_amount"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.leads_number"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.customer_number"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.deals_number"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.costs_number"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.conversion_from_lead_to_customer"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.lead_cost"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.average_payment"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.ltv"),
        });

        int i = 0;
        for (ExportLeadGenMethodWithProjects exportLeadGenMethodWithProjects : exportLeadGenMethodDTO.getResultedObjects()) {
            i++;

            LeadGenMethod leadGenMethod = exportLeadGenMethodWithProjects.getLeadGenMethod();
            IndicatorsDTO indicatorsDAO = leadGenMethod.getCachedKPIsData();

            data.add(new Object[]{String.valueOf(i),
                    leadGenMethod.getId(),
                    leadGenMethod.getName(),
                    leadGenMethod.isActive() ? messageByLocaleService.getMessage("common.true") : messageByLocaleService.getMessage("common.false"),

                    indicatorsDAO.getROI().multiply(new BigDecimal("100")),
                    indicatorsDAO.getDealsAmounts(),
                    indicatorsDAO.getCostsAmount(),
                    indicatorsDAO.getLeadsNumber(),
                    indicatorsDAO.getCustomersNumber(),
                    indicatorsDAO.getDealsNumber(),
                    indicatorsDAO.getCostsNumber(),
                    indicatorsDAO.getConversionFromLeadToCustomer(),
                    indicatorsDAO.getLeadCost(),
                    indicatorsDAO.getAveragePayment(),
                    indicatorsDAO.getLtv()
            });
        }

        // lead gen method and projects info
        if (leadGenMethodBuilder.isDevideProjectsPerSheets()) {
            for (ExportLeadGenMethodWithProjects exportLeadGenMethodWithProjects : exportLeadGenMethodDTO.getResultedObjects()) {
                LeadGenMethod leadGenMethod = exportLeadGenMethodWithProjects.getLeadGenMethod();

                LeadGenProjectFilter filter = new LeadGenProjectFilter();
                filter.setLeadGenMethodID(leadGenMethod.getId());

                List<LeadGenProject> leadGenProjects = exportLeadGenMethodWithProjects.getLeadGenProjects();

                String sheetName;
                sheetName = leadGenMethod.getName() == null ? leadGenMethod.getId() : leadGenMethod.getName();
                SheetData sheetData = basicExcelObjectPrinter.createSheet(sheetName);
                excelAddLeadGenProjectSheetToWorkBook(sheetData, leadGenProjects);
            }
        } else {
            // all projects on one sheet
            List<LeadGenProject> leadGenProjects = new ArrayList<>();

            for (ExportLeadGenMethodWithProjects exportLeadGenMethodWithProjects : exportLeadGenMethodDTO.getResultedObjects()) {
                leadGenProjects.addAll(exportLeadGenMethodWithProjects.getLeadGenProjects());
            }

            SheetData sheetData = basicExcelObjectPrinter.createSheet(messageByLocaleService.getMessage("excel.export.leadgenmethod.project.sheet"));
            excelAddLeadGenProjectSheetToWorkBook(sheetData, leadGenProjects);
        }
        basicExcelObjectPrinter.build();

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        try {
            workbook.write(outputStream);
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage());
        }

        return outputStream.toByteArray();
    }

    private void excelAddLeadGenProjectSheetToWorkBook(SheetData sheetData, List<LeadGenProject> leadGenProjects) {
        List<Object[]> data = sheetData.getData();

        sheetData.setHeaders(new LocalizedExcelHeader[]{
                LocalizedExcelHeader.of("№", false,false),
                LocalizedExcelHeader.of("ID", false,false),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.leadgenmethodid"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.name"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.active"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.ROI"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.deals_amount"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.costs_amount"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.leads_number"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.customer_number"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.deals_number"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.costs_number"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.conversion_from_lead_to_customer"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.lead_cost"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.average_payment"),
                LocalizedExcelHeader.of("excel.export.leadgenmethod.header.ltv")
        });

        int i = 0;
        for (LeadGenProject leadGenMethod : leadGenProjects) {
            i++;
            IndicatorsDTO indicatorsDAO = leadGenMethod.getCachedKPIsData();

            data.add(new Object[]{
                    String.valueOf(i), //index
                    leadGenMethod.getId(),
                    leadGenMethod.getLeadGenMethodId(),
                    leadGenMethod.getName(),
                    leadGenMethod.isActive() ? messageByLocaleService.getMessage("common.true") : messageByLocaleService.getMessage("common.false"),

                    indicatorsDAO.getROI().multiply(new BigDecimal("100")),
                    indicatorsDAO.getDealsAmounts(),
                    indicatorsDAO.getCostsAmount(),
                    indicatorsDAO.getLeadsNumber(),
                    indicatorsDAO.getCustomersNumber(),
                    indicatorsDAO.getDealsNumber(),
                    indicatorsDAO.getCostsNumber(),
                    indicatorsDAO.getConversionFromLeadToCustomer(),
                    indicatorsDAO.getLeadCost(),
                    indicatorsDAO.getAveragePayment(),
                    indicatorsDAO.getLtv()
            });
        }

    }
}
