package com.ambh.exceltools;

import org.apache.xmlbeans.SystemProperties;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;

public class App {

    public static void main(String[] args) throws IOException {

        String fileName = "excelTest.xlsx";
        final String userHome = SystemProperties.getProperty("user.home");

        File file = new File(userHome + File.separator + fileName);
        ExcelReaderWriter excelReaderWriter = new ExcelReaderWriter(file);

        excelReaderWriter.process(data -> new ExcelReaderWriter.RowInfoToWrite(0,
                Arrays.stream(data)
                        .map(str -> str + "_1")
                        .toArray(String[]::new))
        );
    }
}
