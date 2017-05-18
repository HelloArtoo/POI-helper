package com.gobue.excel.internal.utils;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author hurong
 * 
 * @description: Add the class description here.
 */
public class DateUtils {

    public static final String DATETIME_FORMAT = "yyyy-MM-dd HH:mm:ss";
    public static final String DATETIME_FORMAT_SEC = "yyyy-MM-dd HH:mm:ss.SSS";
    public static final String DATETIME_FORMAT_MIN = "yyyy-MM-dd HH:mm";
    public static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd";
    public static final String DATE_FORMAT_NOLINE = "yyyyMMdd";
    public static final String DATE_FORMAT_NOLINE2 = "yyyyMMddHHmmssSSSS";

    public static String getDateStr(Date date, String format) {

        SimpleDateFormat sdf = new SimpleDateFormat(format);
        if (null != date) {
            return sdf.format(date);
        }
        return null;
    }
}
