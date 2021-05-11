package com.poitest.utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.zip.GZIPInputStream;
import java.util.zip.GZIPOutputStream;

public class ZipUtil {

	public ZipUtil() {
	}

	public static byte[] gZip(byte data[]) throws Exception {
		ByteArrayOutputStream bos = null;
		GZIPOutputStream gzip;
		bos = null;
		gzip = null;
		try {
			bos = new ByteArrayOutputStream();
			gzip = new GZIPOutputStream(bos);
			gzip.write(data);
			gzip.finish();
			return bos.toByteArray();
		} catch (Exception ex) {
			throw ex;
		} finally {
			if (gzip != null) {
				gzip.close();
			}
			if (bos != null) {
				bos.close();
			}
		}
	}

	public static byte[] unGZip(byte data[]) throws Exception {
		byte b[];
		ByteArrayInputStream bis;
		GZIPInputStream gzip;
		ByteArrayOutputStream baos;
		b = null;
		bis = null;
		gzip = null;
		baos = null;
		try {
			bis = new ByteArrayInputStream(data);
			gzip = new GZIPInputStream(bis);
			byte buf[] = new byte[1024];
			int num = -1;
			baos = new ByteArrayOutputStream();
			while ((num = gzip.read(buf, 0, buf.length)) != -1)
				baos.write(buf, 0, num);
			b = baos.toByteArray();
			baos.flush();
		} catch (Exception ex) {
			throw ex;
		} finally {
			if (baos != null) {
				baos.close();
			}
			if (gzip != null) {
				gzip.close();
			}
			if (bis != null) {
				bis.close();
			}
		}
		return b;
	}
}