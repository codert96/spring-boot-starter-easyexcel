package com.github.codert96.excel.handler;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.github.codert96.excel.annotations.RequestExcel;
import com.github.codert96.excel.annotations.ResponseExcel;
import com.github.codert96.excel.bean.ExcelFieldError;
import com.github.codert96.excel.exceptions.ExcelNotValidException;
import com.github.codert96.excel.exceptions.IllegalExcelException;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.BeanUtils;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.context.ApplicationListener;
import org.springframework.context.event.ApplicationEventMulticaster;
import org.springframework.context.expression.BeanFactoryResolver;
import org.springframework.core.MethodParameter;
import org.springframework.core.ResolvableType;
import org.springframework.expression.spel.SpelParserConfiguration;
import org.springframework.expression.spel.standard.SpelExpressionParser;
import org.springframework.expression.spel.support.StandardEvaluationContext;
import org.springframework.http.CacheControl;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.lang.NonNull;
import org.springframework.util.ResourceUtils;
import org.springframework.util.StreamUtils;
import org.springframework.util.StringUtils;
import org.springframework.validation.BeanPropertyBindingResult;
import org.springframework.validation.BindingResult;
import org.springframework.validation.FieldError;
import org.springframework.validation.SmartValidator;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.support.WebDataBinderFactory;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.support.RequestHandledEvent;
import org.springframework.web.method.support.HandlerMethodArgumentResolver;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.method.support.ModelAndViewContainer;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartRequest;

import java.io.*;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.stream.StreamSupport;

@Slf4j
@SuppressWarnings("unused")
public class ExcelRequestResponseResolverHandler implements HandlerMethodArgumentResolver, HandlerMethodReturnValueHandler, ApplicationListener<RequestHandledEvent> {

    private static final String PROCESSOR_KEY_PREFIX = ExcelRequestResponseResolverHandler.class.getName() + "_";
    private final ApplicationContext applicationContext;
    private final List<Class<? extends Converter<?>>> converters;
    private final SmartValidator smartValidator;
    private final SpelExpressionParser expressionParser = new SpelExpressionParser(new SpelParserConfiguration(true, true));
    private final ThreadLocal<List<Path>> tempFiles = ThreadLocal.withInitial(() -> Collections.synchronizedList(new ArrayList<>()));

    public ExcelRequestResponseResolverHandler(ApplicationContext applicationContext, List<Class<? extends Converter<?>>> converters, SmartValidator smartValidator) {
        this.applicationContext = applicationContext;
        this.converters = converters;
        this.smartValidator = smartValidator;

        applicationContext.getBean(ApplicationEventMulticaster.class).addApplicationListener(this);
    }

    private static String excelProperty(Map<String, String> map, Object o, String field) {
        return map.computeIfAbsent(o.getClass() + field, s -> {
            try {
                Field declaredField = o.getClass().getDeclaredField(field);
                ExcelProperty property = declaredField.getAnnotation(ExcelProperty.class);
                String cell = "";
                if (Objects.nonNull(property)) {
                    String[] value = property.value();
                    if (value.length != 0) {
                        cell = value[value.length - 1];
                    }
                }
                return cell;
            } catch (ReflectiveOperationException e) {
                throw new RuntimeException(e);
            }
        });
    }

    private Processor getProcessor() {
        return Optional.ofNullable(RequestContextHolder.getRequestAttributes())
                .map(requestAttributes -> (Processor) requestAttributes.getAttribute(PROCESSOR_KEY_PREFIX, RequestAttributes.SCOPE_REQUEST))
                .orElse(null);
    }

    public static void setProcessor(Processor processor) {
        Optional.ofNullable(RequestContextHolder.getRequestAttributes())
                .ifPresent(requestAttributes -> requestAttributes.setAttribute(PROCESSOR_KEY_PREFIX, processor, RequestAttributes.SCOPE_REQUEST));
    }

    @Override
    public boolean supportsParameter(@NonNull MethodParameter parameter) {
        return parameter.hasParameterAnnotation(RequestExcel.class) && List.class.isAssignableFrom(parameter.getParameterType());
    }

    @Override
    public Object resolveArgument(@NonNull MethodParameter parameter, ModelAndViewContainer mavContainer, @NonNull NativeWebRequest webRequest, WebDataBinderFactory binderFactory) throws Exception {
        parameter = parameter.nestedIfOptional();
        HttpServletRequest httpServletRequest = Objects.requireNonNull(webRequest.getNativeRequest(HttpServletRequest.class));

        RequestExcel requestExcel = Objects.requireNonNull(parameter.getParameterAnnotation(RequestExcel.class));
        String tempKey = "tempFile_" + requestExcel.value();
        ExcelReaderBuilder builder = EasyExcel.read();
        Object attribute = httpServletRequest.getAttribute(tempKey);

        Path tempFile = (Path) attribute;
        if (Objects.isNull(attribute)) {
            tempFile = createTempFile();
        }
        httpServletRequest.setAttribute(tempKey, tempFile);
        if (Objects.isNull(attribute)) {
            try {
                if (httpServletRequest instanceof MultipartRequest multipartRequest) {
                    MultipartFile file = Objects.requireNonNull(multipartRequest.getFile(requestExcel.value()));
                    file.transferTo(tempFile.toFile());
                } else {
                    Files.copy(httpServletRequest.getInputStream(), tempFile, StandardCopyOption.REPLACE_EXISTING);
                }
            } catch (Exception e) {
                log.error(e.getMessage(), e);
            }
        }
        final Path checkFile = tempFile;
        Arrays.stream(requestExcel.checker())
                .map(BeanUtils::instantiateClass)
                .peek(this::setApplicationContext)
                .forEach(excelChecker -> {
                    try {
                        if (!excelChecker.check(checkFile)) {
                            throw new IllegalExcelException();
                        }
                    } catch (Exception e) {
                        throw new IllegalExcelException(e);
                    }
                });
        builder.file(Files.newInputStream(tempFile));


        builder.ignoreEmptyRow(requestExcel.ignoreEmptyRow())
                .head(ResolvableType.forMethodParameter(parameter).getGeneric(0).resolve())
                .headRowNumber(requestExcel.headRowNumber())
                .autoCloseStream(true);

        String password = requestExcel.password();
        if (StringUtils.hasText(password)) {
            builder.password(password);
        }
        converters.stream()
                .map(BeanUtils::instantiateClass)
                .peek(this::setApplicationContext)
                .forEach(builder::registerConverter);
        List<Object> list = new ArrayList<>();
        if (requestExcel.sheetIndex() >= 0) {
            list.addAll(
                    builder.sheet(requestExcel.sheetIndex())
                            .doReadSync()
            );
        } else if (StringUtils.hasText(requestExcel.sheetName())) {
            list.addAll(
                    builder.sheet(requestExcel.sheetName())
                            .doReadSync()
            );
        } else {
            list.addAll(
                    builder.doReadAllSync()
            );
        }
        if (requestExcel.validate()) {
            Map<String, String> cacheMap = new HashMap<>();
            Object[] validateGroups = requestExcel.validateGroups();

            String bindingResultKey = BindingResult.MODEL_KEY_PREFIX + "bindingResult";
            BindingResult bindingResult = Optional.ofNullable(
                            httpServletRequest.getAttribute(bindingResultKey)
                    )
                    .map(BindingResult.class::cast)
                    .map(tmp -> {
                        //noinspection unchecked
                        Optional.ofNullable(tmp.getTarget())
                                .map(List.class::cast)
                                .ifPresent(oldList -> oldList.addAll(list));
                        return tmp;
                    })
                    .orElseGet(() -> {
                        try {
                            WebDataBinder dataBinder = binderFactory.createBinder(webRequest, list, "bindingResult");
                            return dataBinder.getBindingResult();
                        } catch (Exception e) {
                            throw new RuntimeException(e);
                        }
                    });

            httpServletRequest.setAttribute(bindingResultKey, bindingResult);
            for (int i = 0; i < list.size(); i++) {
                Object target = list.get(i);
                BeanPropertyBindingResult result = new BeanPropertyBindingResult(target, target.getClass().getName());
                smartValidator.validate(target, result, validateGroups);
                final int rows = i + requestExcel.headRowNumber() + 1;
                result.getAllErrors()
                        .stream()
                        .map(FieldError.class::cast)
                        .map(error ->
                                new ExcelFieldError(
                                        error.getObjectName(),
                                        error.getField(),
                                        error.getRejectedValue(),
                                        error.isBindingFailure(),
                                        error.getCodes(),
                                        error.getArguments(),
                                        excelProperty(cacheMap, target, error.getField()) + error.getDefaultMessage(),
                                        rows,
                                        requestExcel,
                                        target
                                )
                        )
                        .forEach(bindingResult::addError);
            }
            cacheMap.clear();
            if (requestExcel.validateThrow() && bindingResult.getErrorCount() != 0) {
                BeanPropertyBindingResult propertyBindingResult = new BeanPropertyBindingResult(list, "bindingResult");
                bindingResult.getAllErrors()
                        .stream()
                        .map(ExcelFieldError.class::cast)
                        .filter(objectError -> Objects.equals(objectError.getRequestExcel(), requestExcel))
                        .forEach(propertyBindingResult::addError);
                throw new ExcelNotValidException(parameter, propertyBindingResult);
            }
            if (Objects.nonNull(mavContainer)) {
                mavContainer.addAttribute(bindingResultKey, bindingResult);
            }
        }
        return list;
    }

    @Override
    public boolean supportsReturnType(@NonNull MethodParameter returnType) {
        return returnType.hasMethodAnnotation(ResponseExcel.class) && List.class.isAssignableFrom(Objects.requireNonNull(returnType.getMethod()).getReturnType());
    }

    @Override
    public void handleReturnValue(Object returnValue, @NonNull MethodParameter returnType, @NonNull ModelAndViewContainer mavContainer, @NonNull NativeWebRequest webRequest) throws Exception {
        Path tempFile = createTempFile();
        ResponseExcel responseExcel = Objects.requireNonNull(returnType.getMethodAnnotation(ResponseExcel.class));
        Class<?> resolve = responseExcel.headClass();
        if (resolve.equals(Void.class)) {
            resolve = ResolvableType.forMethodReturnType(Objects.requireNonNull(returnType.getMethod())).getGeneric(0).resolve();
        }
        if (Objects.isNull(resolve) && returnValue instanceof List<?> list && !list.isEmpty()) {
            resolve = list.stream().filter(Objects::nonNull).findFirst().map(Object::getClass).orElse(null);
        }
        ExcelWriterBuilder builder = EasyExcel.write(tempFile.toFile())
                .head(resolve)
                .autoCloseStream(true);
        Arrays.stream(responseExcel.writeHandler())
                .map(BeanUtils::instantiateClass)
                .peek(this::setApplicationContext)
                .forEach(builder::registerWriteHandler);

        converters.stream()
                .map(BeanUtils::instantiateClass)
                .peek(this::setApplicationContext)
                .forEach(builder::registerConverter);

        if (StringUtils.hasText(responseExcel.password())) {
            builder.password(responseExcel.password());
        }
        if (StringUtils.hasText(responseExcel.template()) || StringUtils.hasText(responseExcel.templateSpEL())) {
            if (StringUtils.hasText(responseExcel.templateSpEL())) {
                StandardEvaluationContext standardEvaluationContext = new StandardEvaluationContext();
                standardEvaluationContext.setBeanResolver(new BeanFactoryResolver(applicationContext));
                StreamSupport.stream(Spliterators.spliteratorUnknownSize(webRequest.getParameterNames(), 0), false)
                        .forEach(s -> standardEvaluationContext.setVariable(s, webRequest.getParameter(s)));
                Path path = createTempFile();
                try (InputStream inputStream = Objects.requireNonNull(expressionParser.parseExpression(responseExcel.templateSpEL()).getValue(standardEvaluationContext, InputStream.class))) {
                    Files.copy(inputStream, path, StandardCopyOption.REPLACE_EXISTING);
                }
                builder.withTemplate(path.toFile());
            } else {
                try {
                    String template = responseExcel.template();
                    File resource = ResourceUtils.getFile(template);
                    builder.withTemplate(resource);
                } catch (Exception e) {
                    log.error(e.getMessage(), e);
                }
            }

            try (ExcelWriter excelWriter = builder.build()) {
                WriteSheet writeSheet = EasyExcel.writerSheet(responseExcel.sheetIndex(), responseExcel.sheetName()).build();
                Processor processor = getProcessor();
                if (Objects.nonNull(processor)) {
                    processor.exec(responseExcel, excelWriter, writeSheet);
                } else {
                    ResponseExcel.Config config = responseExcel.config();
                    excelWriter.fill(returnValue,
                            FillConfig
                                    .builder()
                                    .direction(config.direction())
                                    .forceNewRow(config.forceNewRow())
                                    .build(),
                            writeSheet
                    );
                }
                excelWriter.finish();
            }
        } else {
            try (ExcelWriter excelWriter = builder.build()) {
                excelWriter.write((List<?>) returnValue, EasyExcel.writerSheet(responseExcel.sheetIndex(), responseExcel.sheetName()).build());
                excelWriter.finish();
            }
        }

        HttpServletResponse nativeResponse = Objects.requireNonNull(webRequest.getNativeResponse(HttpServletResponse.class));

        nativeResponse.setCharacterEncoding(StandardCharsets.UTF_8.name());
        nativeResponse.setContentType(responseExcel.contentType());
        nativeResponse.setHeader(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                .filename("%s%s".formatted(responseExcel.filename(), responseExcel.suffix()), StandardCharsets.UTF_8)
                .toString()
        );
        nativeResponse.setContentLengthLong(Files.size(tempFile));
        nativeResponse.setHeader(HttpHeaders.CACHE_CONTROL, CacheControl.noStore().getHeaderValue());
        mavContainer.setRequestHandled(true);
        try (InputStream inputStream = new BufferedInputStream(Files.newInputStream(tempFile));
             OutputStream outputStream = new BufferedOutputStream(nativeResponse.getOutputStream())
        ) {
            StreamUtils.copy(inputStream, outputStream);
        }
    }

    private void setApplicationContext(Object o) {
        if (o instanceof ApplicationContextAware applicationContextAware) {
            applicationContextAware.setApplicationContext(applicationContext);
        }
    }

    private Path createTempFile() throws IOException {
        Path tempFile = Files.createTempFile("", ".tmp");
        tempFiles.get().add(tempFile);
        log.debug("创建临时文件：{}", tempFile);
        return tempFile;
    }

    @Override
    public void onApplicationEvent(@NonNull RequestHandledEvent event) {
        try {
            List<Path> paths = tempFiles.get();
            for (Path path : paths) {
                if (Files.deleteIfExists(path)) {
                    log.debug("删除临时文件：{}", path);
                }
            }
            paths.clear();
            tempFiles.remove();
        } catch (Exception ignored) {
        }
    }

    @FunctionalInterface
    public interface Processor {
        void exec(ResponseExcel responseExcel, ExcelWriter excelWriter, WriteSheet writeSheet);
    }


}
