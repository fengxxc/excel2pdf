package com.github.fengxxc;

import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class Excel2PDFBlankTest {
    @Test
    public void test() {
        try (InputStream is = Excel2PDF.class.getResourceAsStream("/blank.xlsx");
             OutputStream os = new FileOutputStream("C:\\temp\\excel2pdf\\blank_test.pdf")
        ) {
            Excel2PDF.process(is, os);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
