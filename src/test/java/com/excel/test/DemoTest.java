package com.excel.test;

import com.excel.utilities.Excel;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.util.List;
import java.util.Map;

public class DemoTest {

    String path = "src/main/resources/Book1.xlsx";

    @DataProvider
    public Object[][] mapSet(){
        Excel excel = new Excel();
        return excel.getMapObjects("src/main/resources/Book1.xlsx","Sheet2");
    }

    @DataProvider
    public Object[][] listSet(){
        Excel excel = new Excel();
        return excel.getListObjects(path,"Sheet2");
    }


    @Test(dataProvider = "mapSet")
    public void test1(Map<String,String> data){
        System.out.println(data);
    }

    @Test(dataProvider = "listSet")
    public void test2(List<String> data){
        System.out.println(data);
    }

    @Test
    public void test3(){
        Excel excel = new Excel();
        Map<String,String> data = excel.getDataMap(path,"Sheet3","test1");
        System.out.println(data);
    }





}
