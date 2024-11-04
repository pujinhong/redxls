package com.red.redxls.entity;



import lombok.Data;

import java.io.Serializable;

@Data
public class Result implements Serializable {
	
	private static final long serialVersionUID = 1L;
	
	private boolean success;
    private int code;
    private String msg;
    private Object data;
    private long timestamp;
	
	private Result() {}
    
    public static Result error() {
        return error(null);
    }

    public static Result error(String message) {
        return error(null, message);
    }

    public static Result error(Integer code, String message) {
        if(code == null) {
            code = 500;
        }
        if(message == null) {
            message = "服务器内部错误";
        }
        return build(code, false, message);
    }
    
    public static Result ok() {
        return ok(null);
    }

    public static Result ok(String message) {
        return ok(null, message);
    }

    public static Result ok(Integer code, String message) {
        if(code == null) {
            code = 200;
        }
        if(message == null) {
            message = "操作成功";
        }
        return build(code, true, message);
    }
    
    public static Result build(int code, boolean success, String message) {
        return new Result()
                .setCode(code)
                .setSuccess(success)
                .setMessage(message)
                .setTimestamp(System.currentTimeMillis());
    }
    
    public Result setCode(int code) {
        this.code = code;
        return this;
    }

    public Result setSuccess(boolean success) {
        this.success = success;
        return this;
    }
    public Result setMessage(String msg) {
        this.msg = msg;
        return this;
    }

    public Result setData(Object data) {
        this.data = data;
        return this;
    }

    public Result setTimestamp(long timestamp) {
        this.timestamp = timestamp;
        return this;
    }
    
}
