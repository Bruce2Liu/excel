package com.example.excel.easyexcel;

import com.example.excel.util.TestFileUtil;

import java.io.File;
import java.io.IOException;

/**
 * @author liujunhui
 * @date 2020/12/21 20:01
 */
public class Read {

    public static void main(String[] args) throws IOException {
        File file = TestFileUtil.readUserHomeFile("demo" + File.separator + "read.xlsx");
        System.getProperty("user.home");
        System.out.println(1);

        String a = "123";
        String b = a;

        a="1234";
        System.out.println(b);
    }
}

