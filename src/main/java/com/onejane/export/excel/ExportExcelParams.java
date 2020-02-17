package com.onejane.export.excel;

import com.onejane.enums.ExcelType;
import com.onejane.head.TheadColumn;
import lombok.*;

import java.util.List;
import java.util.Map;

@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExportExcelParams {

    /**
     * 导出Excel的Sheet名称
     */
    private String sheetName;

    /**
     * 导出Excel的类型
     */
    private ExcelType type= ExcelType.XSSF;

    /**
     * 是否设置Excel样式
     */
    private boolean isSetCellStyle=false;


    /**
     * 原始Excel表头  List 类型
     */
    private List<TheadColumn> theadColumnList;

    /**
     * 导出的Excel数据
     */
    private List<Map<String, Object>> dataList;

    /**
     * 自定义的单元格样式
     */
    private List<ExcelCellStyle> excelCellStyleList;

}
