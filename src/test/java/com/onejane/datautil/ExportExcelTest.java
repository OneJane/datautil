package com.onejane.datautil;

import com.onejane.export.ExportCSVUtil;
import com.onejane.export.ExportExcelUtil;
import com.onejane.head.TheadColumn;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExportExcelTest {


    public static void main(String[] args) {
        testExportExcel();
        testExportExcel2();

    }


    public static void testExportExcel  (){
        //数据
        List<Map<String, Object>> dataList = MySQLData.getDataList();
        //表头列表
        List<TheadColumn> theadColumnList = MySQLData.getTheadColumnList();


        OutputStream outExcel= null;
        OutputStream outCSV= null;
        try {
            outExcel = new FileOutputStream("F:\\project\\datautil\\src\\main\\resources\\天气数据.xlsx");
            outCSV = new FileOutputStream("F:\\project\\datautil\\src\\main\\resources\\天气数据.csv");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        //导出Excel
        ExportExcelUtil.exportExcel(outExcel,"天气数据" ,theadColumnList ,dataList);
        //导出CSV(设置的表头样式失效)
        ExportCSVUtil.exportCSV(outCSV,theadColumnList ,dataList);

    }

    public static void testExportExcel2(){
        //数据
        List<Map<String, Object>> dataList = new ArrayList<>();


        Map<String, Object> map = new HashMap<>();
        map.put("deal_create_time","2020-2-12 20:59:38");
        map.put("store_freight","0");
        map.put("audit_count_fees","1");
        map.put("audit_amount_fees","2");
        map.put("member_count_fees","3");
        map.put("member_amount_fees","4");
        map.put("deposit_count_fees_1","5");
        map.put("deposit_amount_fees_1","6");
        map.put("subsidy_count_fees","7");
        map.put("subsidy_amount_fees","8");
        map.put("deposit_count_fees_2","9");
        map.put("deposit_amount_fees_2","10");
        map.put("serve_count_fees","11");
        map.put("serve_amount_fees","12");
        map.put("payment_count","13");
        map.put("payment_amount","14");
        map.put("store_wechat_freight","15");
        map.put("audit_wechat_fees_count","16");
        map.put("audit_wechat_fees_amount","17");
        map.put("member_wechat_fees_count","18");
        map.put("member_wechat_fees_amount","19");
        map.put("deposit_wechat_fees_count","20");
        map.put("deposit_wechat_fees_amount","21");
        map.put("subsidy_wechat_fees_count","22");
        map.put("subsidy_wechat_fees_amount","23");
        map.put("audit_bank_fees_count","24");
        map.put("audit_bank_fees_amount","25");
        map.put("member_bank_fees_count","26");
        map.put("member_bank_fees_amount","27");
        map.put("deposit_bank_fees_count","28");
        map.put("deposit_bank_fees_amount","29");
        map.put("youzan_payment_count","30");
        map.put("youzan_payment_amount","31");
        dataList.add(map);


        TheadColumn dealCreateTimeTheadColumn = new TheadColumn("deal_create_time", null, "deal_create_time", "现金流发生日期");
//        dealCreateTimeTheadColumn.setDownMergeCells(true);//向下合并单元格值相同的列
        dealCreateTimeTheadColumn.setDataTextAlign(HorizontalAlignment.CENTER.name());

        TheadColumn mddGatherTheadColumn = new TheadColumn("mdd_gather", null, "mdd_gather", "摩多多收款合计");


        TheadColumn gatherCollectTheadColumn = new TheadColumn("gather_collect", "mdd_gather", "gather_collect", "收款汇总");
        TheadColumn storeFreightTheadColumn = new TheadColumn("store_freight", "gather_collect", "store_freight", "积分商城-运费");

        TheadColumn auditFeesTheadColumn = new TheadColumn("audit_fees", "gather_collect", "audit_fees", "审核费");
        TheadColumn auditFeesCountTheadColumn = new TheadColumn("audit_count_fees", "audit_fees", "audit_count_fees", "笔");
        TheadColumn auditFeesAmountTheadColumn = new TheadColumn("audit_amount_fees", "audit_fees", "audit_amount_fees", "金额");


        TheadColumn memberFeesTheadColumn = new TheadColumn("member_fees", "gather_collect", "member_fees", "会员费");
        TheadColumn memberFeesCountTheadColumn = new TheadColumn("member_count_fees", "member_fees", "member_count_fees", "笔");
        TheadColumn memberFeesAmountTheadColumn = new TheadColumn("member_amount_fees", "member_fees", "member_amount_fees", "金额");

        TheadColumn depositFeesTheadColumn1 = new TheadColumn("deposit_fees_1", "gather_collect", "deposit_fees_1", "代收定金（扣当日退款金额）");
        TheadColumn depositFeesCountTheadColumn1 = new TheadColumn("deposit_count_fees_1", "deposit_fees_1", "deposit_count_fees_1", "笔");
        TheadColumn depositFeesAmountTheadColumn1 = new TheadColumn("deposit_amount_fees_1", "deposit_fees_1", "deposit_amount_fees_1", "金额");

        TheadColumn subsidyFeesTheadColumn = new TheadColumn("subsidy_fees", "gather_collect", "subsidy_fees", "付给用户的补贴");
        TheadColumn subsidyFeesCountTheadColumn = new TheadColumn("subsidy_count_fees", "subsidy_fees", "subsidy_count_fees", "笔");
        TheadColumn subsidyFeesAmountTheadColumn = new TheadColumn("subsidy_amount_fees", "subsidy_fees", "subsidy_amount_fees", "金额");

        TheadColumn depositFeesTheadColumn2 = new TheadColumn("deposit_fees_2", "gather_collect", "deposit_fees_2", "代付定金（付经销商结算金额）");
        TheadColumn depositFeesCountTheadColumn2 = new TheadColumn("deposit_count_fees_2", "deposit_fees_2", "deposit_count_fees_2", "笔");
        TheadColumn depositFeesAdmountTheadColumn2 = new TheadColumn("deposit_amount_fees_2", "deposit_fees_2", "deposit_amount_fees_2", "金额");

        TheadColumn serveFeesTheadColumn = new TheadColumn("serve_fees", "gather_collect", "serve_fees", "服务费收入（佣金-补贴，计算字端非现金流入，当领取补贴完成时计算）");
        TheadColumn serveFeesCountTheadColumn = new TheadColumn("serve_count_fees", "serve_fees", "serve_count_fees", "笔");
        TheadColumn serveFeesAdmountTheadColumn = new TheadColumn("serve_amount_fees", "serve_fees", "serve_amount_fees", "金额）");

        TheadColumn paymentTheadColumn = new TheadColumn("payment", "gather_collect", "payment", "货款(有赞：扣除退款)");
        TheadColumn paymentCountTheadColumn = new TheadColumn("payment_count", "payment", "payment_count", "笔");
        TheadColumn paymentAmountTheadColumn = new TheadColumn("payment_amount", "payment", "payment_amount", "金额");

        TheadColumn gatherChannelTheadColumn = new TheadColumn("gather_channel", null, "gather_channel", "收款渠道");

        TheadColumn wechatChannelTheadColumn = new TheadColumn("wechat_channel", "gather_channel", "wechat_channel", "1-微信-1558676781（摩多多)");
        TheadColumn storeFreightTheadColumn2 = new TheadColumn("store_wechat_freight", "wechat_channel", "store_wechat_freight", "积分商城-运费");
        TheadColumn auditFeesTheadColumn2 = new TheadColumn("audit_wechat_fees", "wechat_channel", "audit_wechat_fees", "审核费");
        TheadColumn auditFeesCountTheadColumn2 = new TheadColumn("audit_wechat_fees_count", "audit_wechat_fees", "audit_wechat_fees_count", "笔");
        TheadColumn auditFeesAmountTheadColumn2 = new TheadColumn("audit_wechat_fees_amount", "audit_wechat_fees", "audit_wechat_fees_amount", "金额");

        TheadColumn memberFeesTheadColumn2 = new TheadColumn("member_wechat_fees", "wechat_channel", "member_wechat_fees", "会员费");
        TheadColumn memberWechatFeesCountTheadColumn = new TheadColumn("member_wechat_fees_count", "member_wechat_fees", "member_wechat_fees_count", "笔");
        TheadColumn memberWechatFeesAmountTheadColumn = new TheadColumn("member_wechat_fees_amount", "member_wechat_fees", "member_wechat_fees_amount", "金额");

        TheadColumn depositFeesTheadColumn3 = new TheadColumn("deposit_wechat_fees", "wechat_channel", "deposit_wechat_fees", "代收定金（扣当日退款金额）");
        TheadColumn depositWechatFeesCountTheadColumn = new TheadColumn("deposit_wechat_fees_count", "deposit_wechat_fees", "deposit_wechat_fees_count", "笔");
        TheadColumn depositWechatFeesAmountTheadColumn = new TheadColumn("deposit_wechat_fees_amount", "deposit_wechat_fees", "deposit_wechat_fees_amount", "金额");

        TheadColumn subsidyFeesTheadColumn2 = new TheadColumn("subsidy_wechat_fees", "wechat_channel", "subsidy_wechat_fees", "付给用户的补贴");
        TheadColumn subsidyWechatFeesCountTheadColumn2 = new TheadColumn("subsidy_wechat_fees_count", "subsidy_wechat_fees", "subsidy_wechat_fees_count", "笔");
        TheadColumn subsidyWechatFeesAmountTheadColumn2 = new TheadColumn("subsidy_wechat_fees_amount", "subsidy_wechat_fees", "subsidy_wechat_fees_amount", "金额");


        TheadColumn bankChannelTheadColumn = new TheadColumn("bank_channel", "gather_channel", "bank_channel", "2-银行");
        TheadColumn auditFeesTheadColumn3 = new TheadColumn("audit_bank_fees", "bank_channel", "audit_bank_fees", "审核费");
        TheadColumn auditBankFeesCountTheadColumn = new TheadColumn("audit_bank_fees_count", "audit_bank_fees", "audit_bank_fees_count", "笔");
        TheadColumn auditBankFeesAmountTheadColumn = new TheadColumn("audit_bank_fees_amount", "audit_bank_fees", "audit_bank_fees_amount", "金额");

        TheadColumn memberFeesTheadColumn3 = new TheadColumn("member_bank_fees", "bank_channel", "member_bank_fees", "会员费");
        TheadColumn memberBankFeesCountTheadColumn = new TheadColumn("member_bank_fees_count", "member_bank_fees", "member_bank_fees_count", "笔");
        TheadColumn memberBankFeesAmountTheadColumn = new TheadColumn("member_bank_fees_amount", "member_bank_fees", "member_bank_fees_amount", "金额");

        TheadColumn depositFeesTheadColumn4 = new TheadColumn("deposit_bank_fees", "bank_channel", "deposit_bank_fees", "代付定金（付经销商结算金额）");
        TheadColumn depositBankFeesCountTheadColumn = new TheadColumn("deposit_bank_fees_count", "deposit_bank_fees", "deposit_bank_fees_count", "笔");
        TheadColumn depositBankFeesAmountTheadColumn = new TheadColumn("deposit_bank_fees_amount", "deposit_bank_fees", "deposit_bank_fees_amount", "金额");


        TheadColumn youzanChannelTheadColumn = new TheadColumn("youzan_channel", "gather_channel", "youzan_channel", "3-有赞商城");
        TheadColumn paymentTheadColumn2 = new TheadColumn("youzan_payment", "youzan_channel", "youzan_payment", "货款(有赞：扣除退款)");
        TheadColumn youzanPaymentCountTheadColumn = new TheadColumn("youzan_payment_count", "youzan_payment", "youzan_payment_count", "笔");
        TheadColumn youzanPaymentAmountTheadColumn = new TheadColumn("youzan_payment_amount", "youzan_payment", "youzan_payment_amount", "金额");


        List<TheadColumn> theadColumnList = new ArrayList<>();
        theadColumnList.add(dealCreateTimeTheadColumn);
        theadColumnList.add(mddGatherTheadColumn);
        theadColumnList.add(gatherCollectTheadColumn);
        theadColumnList.add(storeFreightTheadColumn);
        theadColumnList.add(auditFeesTheadColumn);
        theadColumnList.add(auditFeesCountTheadColumn);
        theadColumnList.add(auditFeesAmountTheadColumn);
        theadColumnList.add(memberFeesTheadColumn);
        theadColumnList.add(memberFeesCountTheadColumn);
        theadColumnList.add(memberFeesAmountTheadColumn);
        theadColumnList.add(depositFeesTheadColumn1);
        theadColumnList.add(depositFeesCountTheadColumn1);
        theadColumnList.add(depositFeesAmountTheadColumn1);
        theadColumnList.add(subsidyFeesTheadColumn);
        theadColumnList.add(subsidyFeesCountTheadColumn);
        theadColumnList.add(subsidyFeesAmountTheadColumn);
        theadColumnList.add(depositFeesTheadColumn2);
        theadColumnList.add(depositFeesCountTheadColumn2);
        theadColumnList.add(depositFeesAdmountTheadColumn2);
        theadColumnList.add(serveFeesTheadColumn);
        theadColumnList.add(serveFeesCountTheadColumn);
        theadColumnList.add(serveFeesAdmountTheadColumn);
        theadColumnList.add(paymentTheadColumn);
        theadColumnList.add(paymentCountTheadColumn);
        theadColumnList.add(paymentAmountTheadColumn);

        theadColumnList.add(gatherChannelTheadColumn);
        theadColumnList.add(wechatChannelTheadColumn);
        theadColumnList.add(storeFreightTheadColumn2);
        theadColumnList.add(auditFeesTheadColumn2);
        theadColumnList.add(auditFeesCountTheadColumn2);
        theadColumnList.add(auditFeesAmountTheadColumn2);
        theadColumnList.add(memberFeesTheadColumn2);
        theadColumnList.add(memberWechatFeesCountTheadColumn);
        theadColumnList.add(memberWechatFeesAmountTheadColumn);
        theadColumnList.add(depositFeesTheadColumn3);
        theadColumnList.add(depositWechatFeesCountTheadColumn);
        theadColumnList.add(depositWechatFeesAmountTheadColumn);
        theadColumnList.add(subsidyFeesTheadColumn2);
        theadColumnList.add(subsidyWechatFeesCountTheadColumn2);
        theadColumnList.add(subsidyWechatFeesAmountTheadColumn2);
        theadColumnList.add(bankChannelTheadColumn);
        theadColumnList.add(auditFeesTheadColumn3);
        theadColumnList.add(auditBankFeesCountTheadColumn);
        theadColumnList.add(auditBankFeesAmountTheadColumn);
        theadColumnList.add(memberFeesTheadColumn3);
        theadColumnList.add(memberBankFeesCountTheadColumn);
        theadColumnList.add(memberBankFeesAmountTheadColumn);
        theadColumnList.add(depositFeesTheadColumn4);
        theadColumnList.add(depositBankFeesCountTheadColumn);
        theadColumnList.add(depositBankFeesAmountTheadColumn);

        theadColumnList.add(youzanChannelTheadColumn);
        theadColumnList.add(paymentTheadColumn2);
        theadColumnList.add(youzanPaymentCountTheadColumn);
        theadColumnList.add(youzanPaymentAmountTheadColumn);


        OutputStream outExcel = null;
        try {
            outExcel = new FileOutputStream("F:\\project\\datautil\\src\\main\\resources\\财务报表.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        //导出Excel
        ExportExcelUtil.exportExcel(outExcel, "F:\\project\\datautil\\src\\main\\resources\\财务报表.xlsx", theadColumnList, dataList);

    }

}
