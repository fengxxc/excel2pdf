package com.github.fengxxc;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import static org.junit.Assert.*;

/**
 * @author fengxxc
 * @date 2022-12-08
 */
public class Excel2PDFTest {

    @Before
    public void setUp() throws Exception {
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void process() {
        try (InputStream is = Excel2PDF.class.getResourceAsStream("/mayilong_resume.xlsx");
             OutputStream os = new FileOutputStream("F:\\temp\\excel2pdf\\mayilong_resume.pdf")
        ) {
            Excel2PDF.process(is, os, document -> {

            });
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}