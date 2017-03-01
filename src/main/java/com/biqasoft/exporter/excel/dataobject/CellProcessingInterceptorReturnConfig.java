/*
 * Copyright (c) 2016. com.biqasoft
 */

package com.biqasoft.exporter.excel.dataobject;

/**
 * Created by Nikita Bakaev, ya@nbakaev.ru on 4/11/2016.
 * All Rights Reserved
 */
public class CellProcessingInterceptorReturnConfig {

    // use when you write to cell manually instead of auto write
    // depend on Object type
    private boolean skipValueWrite;


    public boolean isSkipValueWrite() {
        return skipValueWrite;
    }

    public void setSkipValueWrite(boolean skipValueWrite) {
        this.skipValueWrite = skipValueWrite;
    }
}
