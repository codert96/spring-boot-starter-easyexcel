package com.github.codert96.excel.bean;

import com.github.codert96.excel.annotations.RequestExcel;
import lombok.Getter;
import org.springframework.util.StringUtils;
import org.springframework.validation.FieldError;

@Getter
public class ExcelFieldError extends FieldError {
    private final int row;
    private final RequestExcel requestExcel;
    private final Object target;

    public int getSheetIndex() {
        return requestExcel.sheetIndex();
    }

    public String getSheetName() {
        return requestExcel.sheetName();
    }

    public ExcelFieldError(String objectName,
                           String field,
                           Object rejectedValue,
                           boolean bindingFailure,
                           String[] codes,
                           Object[] arguments,
                           String defaultMessage,
                           int row,
                           RequestExcel requestExcel,
                           Object target
    ) {
        super(objectName, field, rejectedValue, bindingFailure, codes, arguments, defaultMessage);
        this.row = row;
        this.requestExcel = requestExcel;
        this.target = target;
    }


    public String format() {
        StringBuilder stringBuilder = new StringBuilder();
        if (getSheetIndex() >= 0) {
            stringBuilder.append("Sheet");
            stringBuilder.append(getSheetIndex() + 1);
        } else if (StringUtils.hasText(getSheetName())) {
            stringBuilder.append(getSheetName());
        }
        stringBuilder.append("第");
        stringBuilder.append(row);
        stringBuilder.append("行");
        return stringBuilder.toString();
    }

}