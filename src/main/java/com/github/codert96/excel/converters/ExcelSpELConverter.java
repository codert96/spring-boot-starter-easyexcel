package com.github.codert96.excel.converters;

import com.alibaba.excel.converters.string.StringStringConverter;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.github.codert96.excel.annotations.ExcelSpEL;
import com.github.codert96.excel.annotations.ExcelSpEL.Variables;
import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.context.expression.BeanFactoryResolver;
import org.springframework.expression.spel.SpelParserConfiguration;
import org.springframework.expression.spel.standard.SpelExpressionParser;
import org.springframework.expression.spel.support.StandardEvaluationContext;
import org.springframework.lang.NonNull;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;
import java.util.Objects;

public class ExcelSpELConverter extends StringStringConverter implements ApplicationContextAware {
    private final SpelExpressionParser expressionParser = new SpelExpressionParser(new SpelParserConfiguration(true, true));
    private ApplicationContext applicationContext;

    @Override
    public String convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) {
        if (Objects.nonNull(contentProperty)) {
            Field field = contentProperty.getField();
            ExcelSpEL excelSpEL = field.getAnnotation(ExcelSpEL.class);
            if (Objects.nonNull(excelSpEL)) {
                String stringValue = cellData.getStringValue();

                String deserialize = excelSpEL.deserialize();

                StandardEvaluationContext standardEvaluationContext = new StandardEvaluationContext();
                standardEvaluationContext.setBeanResolver(new BeanFactoryResolver(applicationContext));

                standardEvaluationContext.setVariable(Variables.celldata.name(), cellData.clone());
                standardEvaluationContext.setVariable(Variables.rowindex.name().toLowerCase(), cellData.getRowIndex() + 1);
                standardEvaluationContext.setVariable(Variables.cellindex.name().toLowerCase(), cellData.getColumnIndex() + 1);
                standardEvaluationContext.setVariable(Variables.value.name().toLowerCase(), cellData.getStringValue());

                String value = expressionParser.parseExpression(deserialize).getValue(standardEvaluationContext, String.class);

                return StringUtils.hasText(value) ? value : stringValue;
            }
        }

        return super.convertToJavaData(cellData, contentProperty, globalConfiguration);
    }

    @Override
    public WriteCellData<?> convertToExcelData(String value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) {
        if (Objects.nonNull(contentProperty) && Objects.nonNull(contentProperty.getField())) {
            Field field = contentProperty.getField();
            ExcelSpEL excelSpEL = field.getAnnotation(ExcelSpEL.class);
            if (Objects.nonNull(excelSpEL)) {
                String deserialize = excelSpEL.serialize();
                StandardEvaluationContext standardEvaluationContext = new StandardEvaluationContext();
                standardEvaluationContext.setBeanResolver(new BeanFactoryResolver(applicationContext));
                standardEvaluationContext.setVariable(Variables.value.name().toLowerCase(), value);
                String result = expressionParser.parseExpression(deserialize).getValue(standardEvaluationContext, String.class);
                return new WriteCellData<>(Objects.nonNull(result) ? result : value);
            }
        }
        return super.convertToExcelData(value, contentProperty, globalConfiguration);
    }

    @Override
    public void setApplicationContext(@NonNull ApplicationContext applicationContext) throws BeansException {
        this.applicationContext = applicationContext;
    }
}
