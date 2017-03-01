package com.biqasoft.exporter.excel.dataobject;

/**
 * Created by Nikita on 9/22/2016.
 */
public class LocalizedExcelHeader {

    // localized id
    private String name;

    // todo: https://trello.com/c/aBJ3ToFA/349-excel-export-fitler
    private boolean useAutoFilter = false;

    // true if use messageByLocale or if false - just write this value as is
    private boolean isLocalized = true;

    public LocalizedExcelHeader() {
    }

    public LocalizedExcelHeader(String name, boolean useAutoFilter) {
        this.name = name;
        this.useAutoFilter = useAutoFilter;
    }

    public LocalizedExcelHeader(String name) {
        this.name = name;
    }

    public LocalizedExcelHeader(String name, boolean useAutoFilter, boolean isLocalized) {
        this.name = name;
        this.useAutoFilter = useAutoFilter;
        this.isLocalized = isLocalized;
    }

    public static LocalizedExcelHeader of(String name){
        return new LocalizedExcelHeader(name);
    }

    public static LocalizedExcelHeader of(String name, boolean useAutoFilter){
        return new LocalizedExcelHeader(name, useAutoFilter);
    }

    public static LocalizedExcelHeader of(String name, boolean useAutoFilter, boolean isLocalized){
        return new LocalizedExcelHeader(name, useAutoFilter, isLocalized);
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public boolean isUseAutoFilter() {
        return useAutoFilter;
    }

    public void setUseAutoFilter(boolean useAutoFilter) {
        this.useAutoFilter = useAutoFilter;
    }

    public boolean isLocalized() {
        return isLocalized;
    }

    public void setLocalized(boolean localized) {
        isLocalized = localized;
    }


    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        LocalizedExcelHeader header = (LocalizedExcelHeader) o;

        if (useAutoFilter != header.useAutoFilter) return false;
        if (isLocalized != header.isLocalized) return false;
        return name != null ? name.equals(header.name) : header.name == null;

    }

    @Override
    public int hashCode() {
        int result = name != null ? name.hashCode() : 0;
        result = 31 * result + (useAutoFilter ? 1 : 0);
        result = 31 * result + (isLocalized ? 1 : 0);
        return result;
    }
}
