package com.biqasoft.exporter.excel.common.excel;

import com.biqasoft.microservice.i18n.MessageByLocaleService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

/**
 * Created by Nikita on 9/22/2016.
 */
@Service
public class LocalizedSpringContextAware {

    private static MessageByLocaleService messageByLocaleService;

    public static MessageByLocaleService getMessageByLocaleService() {
        return messageByLocaleService;
    }

    @Autowired
    public void setMessageByLocaleService(MessageByLocaleService messageByLocaleService) {
        LocalizedSpringContextAware.messageByLocaleService = messageByLocaleService;
    }
}
