package com.wd.excel.readxls;

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
 * 操作Excel表格的功能类
 */
public class ExcelReader {

    public static void main(String[] args) throws IOException {
        List<DBObject> list = new ExcelReader().readXls();

//		MongoClient db_conn_credit=new MongoClient("192.168.200.53", 27017);
        MongoClient db_conn_credit = new MongoClient("localhost", 27017);
        DB db = db_conn_credit.getDB("blacklist");
        DBCollection coll = db.getCollection("person");
        for (DBObject xlsDto : list) {
            coll.insert(xlsDto);

        }

    }

    /**
     * 读取xls文件内容
     *
     * @return List <XlsDto>对象
     * @throws IOException 输入/输出(i/o)异常
     */
    private List<DBObject> readXls() throws IOException {
        InputStream is = new FileInputStream("C:\\Users\\juxinli01\\Desktop\\blacklist\\p2p.xlsx");
        XSSFWorkbook hssfWorkbook = new XSSFWorkbook(is);
        DBObject xlsDto = null;
        List<DBObject> list = new ArrayList<DBObject>();
        // 循环工作表Sheet
        for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
            XSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
            if (hssfSheet == null) {
                continue;
            }
            // 循环行Row
            for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                XSSFRow hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow == null) {
                    continue;
                }
                xlsDto = new BasicDBObject();
                // 循环列Cell
                // for (int cellNum = 0; cellNum <=4; cellNum++) {
                XSSFCell xh = hssfRow.getCell(0);
                if (xh == null) {
                    continue;
                }
                xlsDto.put("cardSix", getValue(xh));

                XSSFCell xm = hssfRow.getCell(1);
                if (xm == null) {
                    continue;
                }
                xlsDto.put("cardBin", getValue(xm));

                XSSFCell yxsmc = hssfRow.getCell(2);
                if (yxsmc == null) {
                    continue;
                }
                xlsDto.put("issuingBank", getValue(yxsmc));
                XSSFCell kcm = hssfRow.getCell(3);
                if (kcm == null) {
                    continue;
                }
                xlsDto.put("companyCode", getValue(kcm));

                XSSFCell cj = hssfRow.getCell(4);
                if (cj == null) {
                    continue;
                }
                xlsDto.put("bankName", getValue(cj));

                XSSFCell cj5 = hssfRow.getCell(5);
                if (cj5 == null) {
                    continue;
                }
                xlsDto.put("state", getValue(cj5));

                XSSFCell cj6 = hssfRow.getCell(6);
                if (cj6 == null) {
                    continue;
                }
                xlsDto.put("province", getValue(cj6));

                XSSFCell cj7 = hssfRow.getCell(7);
                if (cj7 == null) {
                    continue;
                }
                xlsDto.put("location", getValue(cj7));

                XSSFCell cj8 = hssfRow.getCell(8);
                if (cj8 == null) {
                    continue;
                }
                xlsDto.put("cardName", getValue(cj8));

                XSSFCell cj9 = hssfRow.getCell(9);
                if (cj9 == null) {
                    continue;
                }
                xlsDto.put("cardType", getValue(cj9));

                XSSFCell cj10 = hssfRow.getCell(10);
                if (cj10 == null) {
                    continue;
                }
                xlsDto.put("cardCategory", getValue(cj10));

                XSSFCell cj11 = hssfRow.getCell(11);
                if (cj11 == null) {
                    continue;
                }
                xlsDto.put("qlty", getValue(cj11));

                XSSFCell cj12 = hssfRow.getCell(12);
                if (cj12 == null) {
                    continue;
                }
                xlsDto.put("brand", getValue(cj12));

                XSSFCell cj13 = hssfRow.getCell(13);
                if (cj13 == null) {
                    continue;
                }
                xlsDto.put("product", getValue(cj13));

                XSSFCell cj14 = hssfRow.getCell(14);
                if (cj14 == null) {
                    continue;
                }
                xlsDto.put("lv", getValue(cj14));

                XSSFCell cj15 = hssfRow.getCell(15);
                if (cj15 == null) {
                    continue;
                }
                xlsDto.put("lvNumber", getValue(cj15));

                XSSFCell cj16 = hssfRow.getCell(16);
                if (cj16 == null) {
                    continue;
                }
                xlsDto.put("puka", getValue(cj16));

                XSSFCell cj17 = hssfRow.getCell(17);
                if (cj17 == null) {
                    continue;
                }
                xlsDto.put("silverCard", getValue(cj17));

                XSSFCell cj18 = hssfRow.getCell(18);
                if (cj18 == null) {
                    continue;
                }
                xlsDto.put("goldCard", getValue(cj18));

                XSSFCell cj19 = hssfRow.getCell(19);
                if (cj19 == null) {
                    continue;
                }
                xlsDto.put("platinumCard", getValue(cj19));

                XSSFCell cj20 = hssfRow.getCell(20);
                if (cj20 == null) {
                    continue;
                }
                xlsDto.put("diamondCard", getValue(cj20));

                XSSFCell cj21 = hssfRow.getCell(21);
                if (cj21 == null) {
                    continue;
                }
                xlsDto.put("otherCard", getValue(cj21));

                XSSFCell cj22 = hssfRow.getCell(22);
                if (cj22 == null) {
                    continue;
                }
                xlsDto.put("distinguish", getValue(cj22));
                list.add(xlsDto);

            }
        }
        return list;
    }

    /**
     * 得到Excel表中的值
     *
     * @param hssfCell Excel中的每一个格子
     * @return Excel中每一个格子中的值
     */
    @SuppressWarnings("static-access")
    private String getValue(XSSFCell hssfCell) {
        if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
            // 返回布尔类型的值
            return String.valueOf(hssfCell.getBooleanCellValue());
        } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
            // 返回数值类型的值
            return String.valueOf(hssfCell.getNumericCellValue());
        } else {
            // 返回字符串类型的值
            return String.valueOf(hssfCell.getStringCellValue());
        }
    }

}
