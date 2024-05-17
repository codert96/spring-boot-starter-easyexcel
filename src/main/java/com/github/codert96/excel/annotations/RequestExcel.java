package com.github.codert96.excel.annotations;


import com.github.codert96.excel.config.ExcelChecker;

import java.lang.annotation.*;

@Documented
@Target({ElementType.PARAMETER})
@Retention(RetentionPolicy.RUNTIME)
public @interface RequestExcel {

    String value() default "file";

    int sheetIndex() default -1;

    String sheetName() default "Sheet1";

    String password() default "";

    int headRowNumber() default 1;

    boolean ignoreEmptyRow() default false;

    Class<?>[] validateGroups() default {};

    boolean validate() default false;

    boolean validateThrow() default false;

    Class<? extends ExcelChecker>[] checker() default {};

}
