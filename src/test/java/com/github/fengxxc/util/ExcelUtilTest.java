package com.github.fengxxc.util;

import org.junit.Test;

import static org.junit.Assert.*;

public class ExcelUtilTest {

    @Test
    public void replaceLeadingSpaces() {
        String text = " This is a test";
        System.out.println(text);
        String replace = ExcelUtil.replaceLeadingSpaces(text);
        System.out.println(replace);
        assertEquals("\u00A0This is a test", replace);
    }
}