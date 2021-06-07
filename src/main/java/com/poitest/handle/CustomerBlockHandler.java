package com.poitest.handle;

import com.alibaba.csp.sentinel.slots.block.BlockException;

import javax.servlet.http.HttpServletRequest;

/**
 * 异常处理类，兜底类
 *
 * @Author zhangdj
 * @Date 2021/6/4:16:41
 * @Description
 */
public class CustomerBlockHandler {
    /**
     * 异常处理方法
     * (1) blockHandler 函数访问范围需要是 public，返回类型需要与原方法相匹配，
     * (2) 参数类型需要和原方法相匹配并且最后加一个额外的参数，类型为 BlockException。
     * (3) 注意对应的函数必需为 static 函数
     *
     * @param request
     * @param blockException
     * @return
     */
    public static String handlerException(HttpServletRequest request,
                                          BlockException blockException) {
        return "sentinel error";
    }
}