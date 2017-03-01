/*
 * Copyright (c) 2016. com.biqasoft
 */

package com.biqasoft.exporter.excel.export;

import com.biqasoft.entity.constants.CUSTOM_FIELD_TYPES;
import com.biqasoft.entity.dto.export.excel.ExportCustomObjectDTO;
import com.biqasoft.entity.filters.CustomObjectsDataFilter;
import com.biqasoft.entity.core.objects.CustomObjectData;
import com.biqasoft.entity.objects.CustomObjectTemplate;
import com.biqasoft.exporter.excel.common.excel.ExcelBasicObjectPrinter;
import com.biqasoft.exporter.excel.common.excel.ExcelHelperService;
import com.biqasoft.exporter.excel.dataobject.LocalizedExcelHeader;
import com.biqasoft.exporter.excel.dataobject.SheetData;
import com.biqasoft.microservice.i18n.MessageByLocaleService;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Created by Nikita Bakaev, ya@nbakaev.ru on 1/29/2016.
 * All Rights Reserved
 */
@Service
public class CustomObjectDataExcelService {

    private final ExcelHelperService excelHelperService;
    private final MessageByLocaleService messageByLocaleService;

    @Autowired
    public CustomObjectDataExcelService(ExcelHelperService excelHelperService, MessageByLocaleService messageByLocaleService) {
        this.excelHelperService = excelHelperService;
        this.messageByLocaleService = messageByLocaleService;
    }

    public byte[] printExcel(ExportCustomObjectDTO exportCustomObjectDTO) {
        CustomObjectTemplate customObjectTemplate = exportCustomObjectDTO.getCustomObjectTemplate();
        CustomObjectsDataFilter builder = exportCustomObjectDTO.getBuilder();
        List<CustomObjectData> list = exportCustomObjectDTO.getList();

        ExcelBasicObjectPrinter basicExcelObjectPrinter = new ExcelBasicObjectPrinter(customObjectTemplate.getName(), excelHelperService);
        XSSFWorkbook workbook = basicExcelObjectPrinter.getWorkbook();
        basicExcelObjectPrinter.setDateFormat(exportCustomObjectDTO.getDateFormat());

        //Create a blank sheet with customers
        SheetData mainSheet = basicExcelObjectPrinter.getMainSheet();

        int i = 0;
        List<List<Object>> data = new ArrayList<>();

        ArrayList<LocalizedExcelHeader> headers = new ArrayList<>();
        headers.add(LocalizedExcelHeader.of("№", false,false));
        headers.add(LocalizedExcelHeader.of("ID", false,false));
        headers.add(LocalizedExcelHeader.of("excel.export.customobject.header.name"));
        headers.add(LocalizedExcelHeader.of("excel.export.customobject.header.description"));

        customObjectTemplate.getCustomFields().forEach(y -> headers.add(LocalizedExcelHeader.of(y.getName(),false, false)));

        headers.add(LocalizedExcelHeader.of("excel.export.customobject.header.created.userid"));
        headers.add(LocalizedExcelHeader.of("excel.export.customobject.header.created.data"));

        builder.setUsePagination(false);
        basicExcelObjectPrinter.createMetaSheet(builder, Integer.toUnsignedLong(list.size()), new Date());

        for (CustomObjectData customObjectData : list) {
            i++;

            List<Object> objects = new ArrayList<>();

            objects.add(String.valueOf(i));
            objects.add(customObjectData.getId());
            objects.add(customObjectData.getName());
            objects.add(customObjectData.getDescription());

            customObjectData.getCustomFields().forEach(y -> {
                        switch (y.getType()) {
                            case CUSTOM_FIELD_TYPES.STRING:
                                objects.add(y.getValue().getStringVal());
                                break;
                            case CUSTOM_FIELD_TYPES.INTEGER:
                                objects.add(y.getValue().getIntVal());
                                break;
                            case CUSTOM_FIELD_TYPES.DOUBLE:
                                objects.add(y.getValue().getDoubleVal());
                                break;
                            case CUSTOM_FIELD_TYPES.DATE:
                                objects.add(y.getValue().getDateVal());
                                break;
                            case CUSTOM_FIELD_TYPES.DICTIONARY:
                                objects.add(y.getValue().getDictVal().getValue().getName());
                                break;
                            case CUSTOM_FIELD_TYPES.BOOLEAN:
                                objects.add(y.getValue().getBoolVal());
                                break;
                            default:
                                objects.add(messageByLocaleService.getMessage("common.error"));
                        }
                    }
            );

            objects.add(customObjectData.getCreatedInfo().getCreatedById());
            objects.add(customObjectData.getCreatedInfo().getCreatedDate());

            data.add(objects);
        }

        ////////////////////////////////////////////////////////////////////
        mainSheet.setHeaders(headers.toArray(new LocalizedExcelHeader[headers.size()]));

        List<Object[]> aVoid = data.stream().map(List::toArray).collect(Collectors.toList());
        mainSheet.setData(aVoid);

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
