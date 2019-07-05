package com.au.annotation;

import java.lang.annotation.*;

/**
 * @author:artificialunintelligent
 * @Date:2019-07-05
 * @Time:11:38
 * @Desc: 目前只支持两级标题，建议其他情况按模板定制注解，然后稍加修改
 */
@Inherited
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD })
public @interface ExcelColumn {

    /**
     * column first title name
     * @return
     */
    String firstTitle() default "";

    /**
     * column second title name
     * @return
     */
    String secondTitle() default "";
}
