/*
 * Copyright (c) 2016. com.biqasoft
 */

package com.biqasoft.exporter.excel.export;

import com.biqasoft.entity.dto.export.excel.ExportCustomObjectDTO;
import com.biqasoft.entity.dto.export.excel.ExportCustomersDTO;
import com.biqasoft.entity.dto.export.excel.ExportKPIDTO;
import com.biqasoft.entity.dto.export.excel.ExportLeadGenMethodDTO;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;

import static com.biqasoft.entity.constants.SYSTEM_CONSTS.EXCEL_MIME_TYPE;

/**
 * Export objects to excel
 * <p>
 * Created by ya_000 on 10/5/2015.
 */
@RestController
@RequestMapping("/export")
public class ExportExcelController {

    private final CustomerExportService customerExportService;
    private final CustomObjectDataExcelService customObjectDataExcelService;
    private final KPIsExcelService kpIsExcelService;
    private final LeadGenMethodExcelReportService leadGenMethodExcelReportService;

    @Autowired
    public ExportExcelController(CustomerExportService customerExportService, CustomObjectDataExcelService customObjectDataExcelService,
                                 KPIsExcelService kpIsExcelService, LeadGenMethodExcelReportService leadGenMethodExcelReportService) {
        this.customerExportService = customerExportService;
        this.customObjectDataExcelService = customObjectDataExcelService;
        this.kpIsExcelService = kpIsExcelService;
        this.leadGenMethodExcelReportService = leadGenMethodExcelReportService;
    }

    @RequestMapping(value = "customer", method = RequestMethod.POST)
    ResponseEntity<byte[]> exportCustomers(@RequestBody ExportCustomersDTO exportCustomersDTO, HttpServletRequest request) throws Exception {
        byte[] bytes;
        HttpHeaders headers;
        ResponseEntity<byte[]> responseEntity;

        bytes = customerExportService.getCustomerInEXCEL(exportCustomersDTO);
        headers = new HttpHeaders();

        headers.setContentType(MediaType.parseMediaType(EXCEL_MIME_TYPE));

        responseEntity = new ResponseEntity<>(bytes, headers, HttpStatus.ACCEPTED);
        return responseEntity;
    }

    @RequestMapping(value = "custom_object", method = RequestMethod.POST)
    ResponseEntity<byte[]> exportCustomObject(@RequestBody ExportCustomObjectDTO exportCustomObjectDTO) throws Exception {
        byte[] bytes;
        HttpHeaders headers;
        ResponseEntity<byte[]> responseEntity;

        bytes = customObjectDataExcelService.printExcel(exportCustomObjectDTO);
        headers = new HttpHeaders();

        headers.setContentType(MediaType.parseMediaType(EXCEL_MIME_TYPE));

        responseEntity = new ResponseEntity<>(bytes, headers, HttpStatus.ACCEPTED);
        return responseEntity;
    }

    @RequestMapping(value = "kpi", method = RequestMethod.POST)
    ResponseEntity<byte[]> exportKPI(@RequestBody ExportKPIDTO biqaPaginationResultList ) throws Exception {
        byte[] bytes;
        HttpHeaders headers;
        ResponseEntity<byte[]> responseEntity;

        bytes = kpIsExcelService.getKPisInEXCEL(biqaPaginationResultList);
        headers = new HttpHeaders();

        headers.setContentType(MediaType.parseMediaType(EXCEL_MIME_TYPE));

        responseEntity = new ResponseEntity<>(bytes, headers, HttpStatus.ACCEPTED);
        return responseEntity;
    }

    @RequestMapping(value = "lead_gen_method", method = RequestMethod.POST)
    ResponseEntity<byte[]> exportLeadGenMethod(@RequestBody ExportLeadGenMethodDTO exportLeadGenMethodDTO) throws Exception {
        byte[] bytes;
        HttpHeaders headers;
        ResponseEntity<byte[]> responseEntity;

        bytes = leadGenMethodExcelReportService.getResponseEntity(exportLeadGenMethodDTO);
        headers = new HttpHeaders();

        headers.setContentType(MediaType.parseMediaType(EXCEL_MIME_TYPE));

        responseEntity = new ResponseEntity<>(bytes, headers, HttpStatus.ACCEPTED);
        return responseEntity;
    }

}