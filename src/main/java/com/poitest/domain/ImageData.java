package com.poitest.domain;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.poitest.utils.easyexcel.CustomStringStringConverter;

import java.io.File;
import java.io.InputStream;
import java.net.URL;
import java.util.Arrays;

/**
 * easyexcel中导出图片的类
 */
@ContentRowHeight(100)
@ColumnWidth(100 / 8)
public class ImageData {
    private File file;
    private InputStream inputStream;
    /**
     * 如果string类型 必须指定转换器，string默认转换成string
     */
    @ExcelProperty(converter = CustomStringStringConverter.class)
    private String string;
    private byte[] byteArray;
    /**
     * 根据url导出
     *
     * @since 2.1.1
     */
    private URL url;

    public File getFile() {
        return file;
    }

    public void setFile(File file) {
        this.file = file;
    }

    public InputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    public String getString() {
        return string;
    }

    public void setString(String string) {
        this.string = string;
    }

    public byte[] getByteArray() {
        return byteArray;
    }

    public void setByteArray(byte[] byteArray) {
        this.byteArray = byteArray;
    }

    public URL getUrl() {
        return url;
    }

    public void setUrl(URL url) {
        this.url = url;
    }

    @Override
    public String toString() {
        return "ImageData{" +
                "file=" + file +
                ", inputStream=" + inputStream +
                ", string='" + string + '\'' +
                ", byteArray=" + Arrays.toString(byteArray) +
                ", url=" + url +
                '}';
    }
}
