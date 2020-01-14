package com.tf.print.template.exception;

/**
 * @ClassName TemplateException
 * @Description TODO
 * @Author kyjonny
 * @Date 6/1/2020 6:10 下午
 **/
public class TemplateException extends RuntimeException {
    public TemplateException() {
        super("模版信息异常");
    }

    public TemplateException(String message) {
        super(message);
    }

    public TemplateException(String message, Throwable cause) {
        super(message, cause);
    }

    public TemplateException(Throwable cause) {
        super(cause);
    }

    protected TemplateException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
