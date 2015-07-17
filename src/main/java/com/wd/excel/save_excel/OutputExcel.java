package com.wd.excel.save_excel;

import com.mongodb.*;
import org.apache.poi.hssf.usermodel.*;

import java.io.*;
import java.net.UnknownHostException;

/**
 * @author wendong
 * @email wendong@juxinli.com
 * @date 2015/7/6
 */
public class OutputExcel {

    public static void main(String[] args) throws IOException {
        System.out.println("helloworld");

        String fileName = "导出Excel.xls";
        fileName = new String(fileName.getBytes("GBK"), "iso8859-1");

        OutputStream output = new FileOutputStream(new File("D:\\" + fileName));
        BufferedOutputStream bufferedOutPut = new BufferedOutputStream(output);
        // 定义单元格报头
//        String worksheetTitle = "Excel导出Student信息";

        HSSFWorkbook wb = new HSSFWorkbook();

        // 创建单元格样式
        HSSFCellStyle cellStyleTitle = wb.createCellStyle();
        // 指定单元格居中对齐
        cellStyleTitle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 指定单元格垂直居中对齐
        cellStyleTitle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 指定当单元格内容显示不下时自动换行
        cellStyleTitle.setWrapText(true);
        // ------------------------------------------------------------------
        HSSFCellStyle cellStyle = wb.createCellStyle();
        // 指定单元格居中对齐
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 指定单元格垂直居中对齐
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 指定当单元格内容显示不下时自动换行
        cellStyle.setWrapText(true);
        // ------------------------------------------------------------------
        // 设置单元格字体
        HSSFFont font = wb.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontName("宋体");
        font.setFontHeight((short) 200);
        cellStyleTitle.setFont(font);

        // 工作表名
        String name = "name";
        String idCard = "idcard";
        String address = "address";
        String overdueDays = "overdueDays";
        String sex = "sex";
        String borrowingBalance = "borrowingBalance";
        String mobiles = "mobiles";
        String categories = "categories";
        String debt = "debt";


        HSSFSheet sheet = wb.createSheet();

//        ExportExcel exportExcel = new ExportExcel(wb, sheet);
        // 创建报表头部
//        exportExcel.createNormalHead(worksheetTitle, 9);
        // 定义第一行
        HSSFRow row1 = sheet.createRow(0);
        HSSFCell cell1 = row1.createCell(0);

        //第一行第一列

        cell1.setCellStyle(cellStyleTitle);
        cell1.setCellValue(new HSSFRichTextString(name));
        //第一行第er列
        cell1 = row1.createCell(1);
        cell1.setCellStyle(cellStyleTitle);
        cell1.setCellValue(new HSSFRichTextString(idCard));

        //第一行第san列
        cell1 = row1.createCell(2);
        cell1.setCellStyle(cellStyleTitle);
        cell1.setCellValue(new HSSFRichTextString(address));

        //第一行第si列
        cell1 = row1.createCell(3);
        cell1.setCellStyle(cellStyleTitle);
        cell1.setCellValue(new HSSFRichTextString(overdueDays));

        //第一行第wu列
        cell1 = row1.createCell(4);
        cell1.setCellStyle(cellStyleTitle);
        cell1.setCellValue(new HSSFRichTextString(sex));

        //第一行第liu列
        cell1 = row1.createCell(5);
        cell1.setCellStyle(cellStyleTitle);
        cell1.setCellValue(new HSSFRichTextString(borrowingBalance));

        //第一行第qi列
        cell1 = row1.createCell(6);
        cell1.setCellStyle(cellStyleTitle);
        cell1.setCellValue(new HSSFRichTextString(mobiles));

        //第一行第qi列
        cell1 = row1.createCell(7);
        cell1.setCellStyle(cellStyleTitle);
        cell1.setCellValue(new HSSFRichTextString(categories));

        //第一行第qi列
        cell1 = row1.createCell(8);
        cell1.setCellStyle(cellStyleTitle);
        cell1.setCellValue(new HSSFRichTextString(debt));


        OutputExcel oe = new OutputExcel();
        DB bak_blacklist_prd = oe.conn("localhost", 27017, "bak_blacklist_prd");
        DBCollection person = bak_blacklist_prd.getCollection("person");

        DBCursor cursor = person.find();

        //定义第二行
        HSSFRow row = sheet.createRow(1);
        HSSFCell cell = row.createCell(1);

        int i = 0;
        int n = 0;
        while (cursor.hasNext()) {

            DBObject next = cursor.next();

            if (n == 65535)
                break;

            row = sheet.createRow(i + 1);

            cell = row.createCell(0);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(new HSSFRichTextString(next.get("name").toString()));

            cell = row.createCell(1);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(new HSSFRichTextString(next.get("id_card").toString()));

            if (next.get("others") != null) {
                cell = row.createCell(2);
                cell.setCellStyle(cellStyle);
                Object addr = (((BasicDBObject) next.get("others")).get("地址"));
                if (addr != null)
                    cell.setCellValue(new HSSFRichTextString(addr.toString()));

                cell = row.createCell(3);
                cell.setCellStyle(cellStyle);
                Object days = ((BasicDBObject) next.get("others")).get("最大逾期天数");
                if (days != null)
                    cell.setCellValue(new HSSFRichTextString(days.toString()));

                cell = row.createCell(4);
                cell.setCellStyle(cellStyle);
                Object se = ((BasicDBObject) next.get("others")).get("性别");
                if (se != null)
                    cell.setCellValue(new HSSFRichTextString(se.toString()));

                cell = row.createCell(5);
                cell.setCellStyle(cellStyle);
                Object money = ((BasicDBObject) next.get("others")).get("累计借入本金");
                if (money != null)
                    cell.setCellValue(new HSSFRichTextString(money.toString()));
            }

            cell = row.createCell(6);
            cell.setCellStyle(cellStyle);
            Object mobiles1 = next.get("mobiles");
            if (mobiles1 != null)
                cell.setCellValue(new HSSFRichTextString(((BasicDBList) mobiles1).get(0).toString()));

            cell = row.createCell(7);
            Object categories1 = next.get("categories");
            if (categories1 != null) {
                String str = "";
                for (Object o : ((BasicDBList) categories1)) {
                    str += o + ",";
                }
                if (str.endsWith(","))
                    str = str.substring(0, str.length() - 1);
                cell.setCellValue(new HSSFRichTextString(str));
            }
            cell.setCellStyle(cellStyle);

            cell = row.createCell(8);
            Object debt1 = next.get("debt");
            if (debt1 != null)
                cell.setCellValue(new HSSFRichTextString(debt1.toString()));
            cell.setCellStyle(cellStyle);

            i++;
            n++;
        }
        try {
            bufferedOutPut.flush();
            wb.write(bufferedOutPut);
            bufferedOutPut.close();
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Output   is   closed ");
        } finally {
            System.out.println("OK!");
        }
    }

    private static final Integer soTimeOut = 300000;
    private static final Integer connectionsPerHost = 500;
    private static final Integer threadsAllowedToBlockForConnectionMultiplier = 500;

    private DB conn(String host, int port, String database) {
        DB db;
        try {
            MongoClient mongoClient = new MongoClient(new ServerAddress(host, port), new MongoClientOptions.Builder()
                    .socketTimeout(soTimeOut)
                    .connectionsPerHost(connectionsPerHost)
                    .threadsAllowedToBlockForConnectionMultiplier(threadsAllowedToBlockForConnectionMultiplier)
                    .socketKeepAlive(true)
                    .build()
            );
            db = mongoClient.getDB(database);
            return db;
        } catch (UnknownHostException e) {
            e.printStackTrace();
        }
        return null;
    }

}
