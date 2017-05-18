package com.gobue.excel.internal.utils;

import java.io.File;

/**
 * @author hurong
 * 
 * @description: Add the class description here.
 */
public class FileUtils {

    public static boolean isExists(File file) {

        return file != null && file.exists();
    }
}
