package com.alan;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.lang.String;
import java.io.*;

public class MyWorkbook {

    public static void main(String[] args) throws IOException {
        MyWorkbook myWorkbook = new MyWorkbook();

        //  将数据写入Excel表
//        List<User> userList=myWorkbook.getUserList();
//        myWorkbook.creatWorkbook(userList);

        // 读取Excel 表数据
        myWorkbook.readWorkbook();

    }

    /**
     * 将集合数据写人Excel表
     * @param userList
     * @throws IOException
     */
    public void creatWorkbook( List<User> userList) throws IOException {
        //创建Excel对象, 03Excel为 HSSFWorkbook，07Excel为 XSSFWorkbook
        HSSFWorkbook wb = new HSSFWorkbook();
        //通过Excel对象创建字体对象,设置字体样式
        HSSFFont font = wb.createFont();
        font.setFontName("宋体");//字体样式
        font.setFontHeightInPoints((short)18);//字体高度，大小
        font.setColor(HSSFColor.RED.index);//颜色
        font.setBold(true);//粗体

        //通过Excel对象创建单元格式
        HSSFCellStyle cellStyle = wb.createCellStyle();
        //水平对齐
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //垂直对齐
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //将字体样式设置到单元格样式
        cellStyle.setFont(font);

        //通过Excel对象创建一张表
        HSSFSheet sheet1 = wb.createSheet("sheet1");
        //设置表行高
        sheet1.setDefaultRowHeightInPoints(20);
        //设置表列宽
        sheet1.setDefaultColumnWidth(30);

        //将单元格样式（包含字体样式）应用于第几列
        sheet1.setDefaultColumnStyle(0,cellStyle);

        //新建第一行
        HSSFRow row0 = sheet1.createRow(0);
        //新建第一行的第一个单元格
        HSSFCell cell0 = row0.createCell(0);
        //设置单元格内容
        cell0.setCellValue("用户信息表");
        //合并单元格
        sheet1.addMergedRegion(new CellRangeAddress(0,0,0,3));
        //将样式应用于单元格
        cell0.setCellStyle(cellStyle);

        //第二行
        HSSFRow row = sheet1.createRow(1);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("userid");
        cell.setCellStyle(cellStyle);
        HSSFCell cell1 = row.createCell(1);
        cell1.setCellValue("username");
        cell1.setCellStyle(cellStyle);
        HSSFCell cell2 = row.createCell(2);
        cell2.setCellValue("password");
        cell2.setCellStyle(cellStyle);
        HSSFCell cell3 = row.createCell(3);
        cell3.setCellValue("hobby");
        cell3.setCellStyle(cellStyle);

        //将用户集合写入表
        for (int i=0;i < userList.size();i++){
            HSSFRow rows = sheet1.createRow(i+2);
            for (int j=0;j < 4 ;j++){
                HSSFCell cells = rows.createCell(j);
                if (j==0){
                    cells.setCellValue(userList.get(i).userid);
                    cells.setCellStyle(cellStyle);
                    continue;
                }
                if (j==1){
                    cells.setCellValue(userList.get(i).username);
                    cells.setCellStyle(cellStyle);
                    continue;
                }
                if (j==2){
                    cells.setCellValue(userList.get(i).password);
                    cells.setCellStyle(cellStyle);
                    continue;
                }
                if (j==3){
                    cells.setCellValue(userList.get(i).hobby);
                    cells.setCellStyle(cellStyle);
                    continue;
                }
            }
            System.out.println(userList.get(i));
        }
        //输出表
        FileOutputStream output=new FileOutputStream("f:\\workbook.xls");
        wb.write(output);
        output.flush();
        output.close();
    }

    /**
     * 读取表数据
     * @throws IOException
     */
    public void readWorkbook() throws IOException {
        //创建输入流
//        FileInputStream fis = new FileInputStream("E:\JAVA资料\上课\阶段三\刘振兴个人分享\demoWorkbook.xls");
        FileInputStream fis = new FileInputStream("E:\\JAVA资料\\上课\\阶段三\\刘振兴个人分享\\demoWorkbook.xlsx");
        //新建新的Excel，从输入流中读取数据
//        HSSFWorkbook sheets = new HSSFWorkbook(fis);
       XSSFWorkbook sheets=new XSSFWorkbook(fis);
        //从Excel表中拿出第一张表
        XSSFSheet sheet = sheets.getSheetAt(0);
//        HSSFSheet sheet = sheets.getSheetAt(0);
        //循环读取每行的数据
        for (Row r:sheet){
            if (r.getRowNum()<2){
                continue;
            }
            User user = new User();
            //将一行中的某单元格数据赋值给用户。
            user.setUserid((int)r.getCell(0).getNumericCellValue());
            user.setUsername(r.getCell(1).getStringCellValue());
            user.setPassword(r.getCell(2).getStringCellValue());
            user.setHobby(r.getCell(3).getStringCellValue());
            System.out.println(user);
        }
    }

    /**
     * 模拟数据库获取数据
     * @return
     */
    public  List<User> getUserList(){
        List<User> userList =new ArrayList<User>();
        for (int i=0;i<30;i++){
            User user = new User();
            user.setUserid(i);
            user.setUsername("zs"+i);
            user.setPassword("123"+i);
            user.setHobby("敲代码"+i);
            userList.add(user);
        }
        return userList;
    }
}
