package com.onejane.datautil;

import com.onejane.export.ExportExcelUtil;
import com.onejane.head.TheadColumn;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public class ExportExcelTest {


    public static void main(String[] args) {
        testExportExcel();

    }


    public static void testExportExcel  (){
        //数据
        List<Map<String, Object>> dataList = MySQLData.getDataList();
        //表头列表
        List<TheadColumn> theadColumnList = MySQLData.getTheadColumnList();


        OutputStream outExcel= null;
        try {
            outExcel = new FileOutputStream("F:\\project\\datautil\\src\\main\\resources\\天气数据.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        //导出Excel
        ExportExcelUtil.exportExcel(outExcel,"天气数据" ,theadColumnList ,dataList);


    }

}
