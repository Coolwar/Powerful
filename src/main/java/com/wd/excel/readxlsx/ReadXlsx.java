package com.wd.excel.readxlsx;

import com.mongodb.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * ����Excel���Ĺ�����
 */
public class ReadXlsx {

    public static void main(String[] args) throws IOException {
        List<DBObject> list = new ReadXlsx().readXls();

        MongoClient db_conn_credit = new MongoClient("192.168.200.53", 27017);
//        MongoClient db_conn_credit = new MongoClient("localhost", 27017);
        DB db = db_conn_credit.getDB("data_interface");
        DBCollection coll = db.getCollection("bankCard");
        for (DBObject xlsDto : list) {
            coll.insert(xlsDto);

        }

    }

    /**
     * ��ȡxls�ļ�����
     *
     * @return List <XlsDto>����
     * @throws IOException ����/���(i/o)�쳣
     */
    private List<DBObject> readXls() throws IOException {
        InputStream is = new FileInputStream("C:\\Users\\juxinli01\\Desktop\\bank-bin\\bin-test.xlsx");
        XSSFWorkbook hssfWorkbook = new XSSFWorkbook(is);
        DBObject xlsDto = null;
        List<DBObject> list = new ArrayList<DBObject>();
        // ѭ��������Sheet
//        for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
            XSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
//            if (hssfSheet == null) {
//                continue;
//            }
            // ѭ����Row
            for (int rowNum = 1; rowNum < hssfSheet.getLastRowNum(); rowNum++) {
                XSSFRow hssfRow = hssfSheet.getRow(rowNum);
                xlsDto = new BasicDBObject();
                // ѭ����Cell
                // for (int cellNum = 0; cellNum <=4; cellNum++) {
                XSSFCell xh = hssfRow.getCell(0);

                xlsDto.put("cardSix", getValue(xh));

                XSSFCell xm = hssfRow.getCell(1);

                xlsDto.put("cardBin", getValue(xm));

                XSSFCell yxsmc = hssfRow.getCell(2);

                xlsDto.put("issuingBank", getValue(yxsmc));
                XSSFCell kcm = hssfRow.getCell(3);

                xlsDto.put("companyCode", getValue(kcm));

                XSSFCell cj = hssfRow.getCell(4);

                xlsDto.put("bankName", getValue(cj));

                XSSFCell cj5 = hssfRow.getCell(5);

                xlsDto.put("state", getValue(cj5));

                XSSFCell cj6 = hssfRow.getCell(6);

                xlsDto.put("province", getValue(cj6));

                XSSFCell cj7 = hssfRow.getCell(7);

                xlsDto.put("location", getValue(cj7));

                XSSFCell cj8 = hssfRow.getCell(8);

                xlsDto.put("cardName", getValue(cj8));

                XSSFCell cj9 = hssfRow.getCell(9);

                xlsDto.put("cardType", getValue(cj9));

                XSSFCell cj10 = hssfRow.getCell(10);

                xlsDto.put("cardCategory", getValue(cj10));

                XSSFCell cj11 = hssfRow.getCell(11);

                xlsDto.put("qlty", getValue(cj11));

                XSSFCell cj12 = hssfRow.getCell(12);

                xlsDto.put("brand", getValue(cj12));

                XSSFCell cj13 = hssfRow.getCell(13);
                xlsDto.put("product", getValue(cj13));

                XSSFCell cj14 = hssfRow.getCell(14);
                xlsDto.put("lv", getValue(cj14));

                XSSFCell cj15 = hssfRow.getCell(15);
                xlsDto.put("lvNumber", getValue(cj15));

                XSSFCell cj16 = hssfRow.getCell(16);
                xlsDto.put("puka", getValue(cj16));

                XSSFCell cj17 = hssfRow.getCell(17);
                xlsDto.put("silverCard", getValue(cj17));

                XSSFCell cj18 = hssfRow.getCell(18);
                xlsDto.put("goldCard", getValue(cj18));

                XSSFCell cj19 = hssfRow.getCell(19);
                xlsDto.put("platinumCard", getValue(cj19));

                XSSFCell cj20 = hssfRow.getCell(20);
                xlsDto.put("diamondCard", getValue(cj20));

                XSSFCell cj21 = hssfRow.getCell(21);
                xlsDto.put("otherCard", getValue(cj21));

                XSSFCell cj22 = hssfRow.getCell(22);
                xlsDto.put("distinguish", getValue(cj22));

                xlsDto.put("priority", 2);
                list.add(xlsDto);

            }
//        }
        return list;
    }

    /**
     * �õ�Excel���е�ֵ
     *
     * @param hssfCell Excel�е�ÿһ������
     * @return Excel��ÿһ�������е�ֵ
     */
    @SuppressWarnings("static-access")
    private String getValue(XSSFCell hssfCell) {
        if (hssfCell == null)
            return "";
        if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
            // ���ز������͵�ֵ
            return String.valueOf(hssfCell.getBooleanCellValue());
        } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
            // ������ֵ���͵�ֵ
            return String.valueOf(hssfCell.getNumericCellValue());
        } else {
            // �����ַ������͵�ֵ
            return String.valueOf(hssfCell.getStringCellValue());
        }
    }

}
