/*
 * Copyright 2016 the original author or authors.
 */

package com.biqasoft.exporter.excel.configs;

import com.biqasoft.microservice.configs.LoggerConfigHelper;
import org.slf4j.MDC;
import org.springframework.stereotype.Component;

import javax.servlet.ServletRequestEvent;
import javax.servlet.ServletRequestListener;
import javax.servlet.http.HttpServletRequest;

import static com.biqasoft.microservice.configs.LoggerConfigHelper.REQUEST_ID_LOGGER;

/**
 * @author Nikita Bakaev, ya@nbakaev.ru
 *         Date: 7/20/2016
 *         All Rights Reserved
 */
@Component
public class LoggerConfig implements ServletRequestListener {

    public final static String MICROSERVICE_REQUEST_ID_HEADER = "X-Biqa-Request-Id";

    /**
     * REQUEST INITIALIZED
     * @param servletRequestEvent
     */
    @Override
    public void requestInitialized(ServletRequestEvent servletRequestEvent) {
        HttpServletRequest request = (HttpServletRequest) servletRequestEvent.getServletRequest();

        String requestId = request.getHeader(MICROSERVICE_REQUEST_ID_HEADER);
        if (requestId == null){
            requestId = LoggerConfigHelper.generateRequestId();
        }

        MDC.put(REQUEST_ID_LOGGER, requestId);
    }

    /**
     * REQUEST DESTROYED
     * @param servletRequestEvent
     */
    @Override
    public void requestDestroyed(ServletRequestEvent servletRequestEvent) {
        MDC.clear();
    }
}
