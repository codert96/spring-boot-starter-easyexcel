package com.github.codert96.excel.annotations;

import java.lang.annotation.*;

@Documented
@Target({ElementType.FIELD, ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@SuppressWarnings({"unused", "SpellCheckingInspection"})
public @interface ExcelSpEL {

    /**
     * &#x5199;
     */
    String serialize() default "#value";

    /**
     * &#x8BFB;
     */
    String deserialize() default "#value";

    enum Variables {
        value,
        celldata,
        rowindex,
        cellindex
    }
}
