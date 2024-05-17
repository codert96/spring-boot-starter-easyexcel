package com.github.codert96.excel.config;

import com.alibaba.excel.converters.Converter;
import com.github.codert96.excel.converters.ExcelSpELConverter;
import com.github.codert96.excel.handler.ExcelRequestResponseResolverHandler;
import org.springframework.beans.BeansException;
import org.springframework.beans.factory.InitializingBean;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.context.annotation.Configuration;
import org.springframework.lang.NonNull;
import org.springframework.validation.SmartValidator;
import org.springframework.web.servlet.mvc.method.annotation.RequestMappingHandlerAdapter;
import org.springframework.web.servlet.mvc.method.annotation.RequestResponseBodyMethodProcessor;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Supplier;

@Configuration(proxyBeanMethods = false)
public class EasyExcelConfig implements ApplicationContextAware, InitializingBean {
    private static final List<Class<? extends Converter<?>>> CONVERTER_LIST = new ArrayList<>();
    private ApplicationContext applicationContext;

    public static void register(Class<? extends Converter<?>> converter) {
        CONVERTER_LIST.add(converter);
    }

    @Override
    public void setApplicationContext(@NonNull ApplicationContext applicationContext) throws BeansException {
        this.applicationContext = applicationContext;
    }

    private <T> void set(Supplier<List<T>> supplier, T t, Consumer<List<T>> consumer) {
        List<T> arrayList = new ArrayList<>();
        supplier.get().forEach(o -> {
            if (o instanceof RequestResponseBodyMethodProcessor) {
                arrayList.add(t);
            }
            arrayList.add(o);
        });
        consumer.accept(Collections.unmodifiableList(arrayList));
    }

    @Override
    public void afterPropertiesSet() {
        EasyExcelConfig.register(ExcelSpELConverter.class);
        SmartValidator smartValidator = applicationContext.getBean(SmartValidator.class);

        ExcelRequestResponseResolverHandler excelRequestResolverHandler = new ExcelRequestResponseResolverHandler(applicationContext, CONVERTER_LIST, smartValidator);

        RequestMappingHandlerAdapter requestMappingHandlerAdapter = applicationContext.getBean(RequestMappingHandlerAdapter.class);

        set(requestMappingHandlerAdapter::getArgumentResolvers, excelRequestResolverHandler, requestMappingHandlerAdapter::setArgumentResolvers);
        set(requestMappingHandlerAdapter::getReturnValueHandlers, excelRequestResolverHandler, requestMappingHandlerAdapter::setReturnValueHandlers);
    }
}
