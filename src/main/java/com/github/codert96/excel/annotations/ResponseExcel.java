package com.github.codert96.excel.annotations;

import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.write.handler.WriteHandler;

import java.lang.annotation.*;

@Documented
@Target({ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ResponseExcel {

    String filename() default "result";

    String suffix() default ".xlsx";

    int sheetIndex() default 0;

    String sheetName() default "Sheet1";

    String contentType() default "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    String password() default "";

    String template() default "";

    Class<?> headClass() default Void.class;

    /**
     * @return java.io.InputStream
     */
    String templateSpEL() default "";

    Config config() default @Config;

    Class<? extends WriteHandler>[] writeHandler() default {};

    @Documented
    @Retention(RetentionPolicy.RUNTIME)
    @Target({})
    @interface Config {
        boolean forceNewRow() default false;

        WriteDirectionEnum direction() default WriteDirectionEnum.VERTICAL;
    }
}
