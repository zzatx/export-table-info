package cn.zzh;

import com.mchange.v2.c3p0.ComboPooledDataSource;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.*;

import javax.sql.DataSource;
import java.awt.*;
import java.io.FileOutputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws Exception {
        //声明需要导出的数据库
        String dbName = "dsjjmpt_doris";
        //声明book
        XSSFWorkbook book = new XSSFWorkbook();
        //获取Connection,获取db的元数据
        DataSource dataSource=new ComboPooledDataSource();
        Connection con = dataSource.getConnection();
        //声明statemen
        Statement st = con.createStatement();
        //st.execute("use "+dbName);
        DatabaseMetaData dmd = con.getMetaData();
        //获取数据库有多少表
        ResultSet rs = dmd.getTables(dbName,dbName,null,new String[]{"TABLE"});
        //获取所有表名　－　就是一个sheet
        List<String> tables = new ArrayList<String>();
        while(rs.next()){
            String tableName = rs.getString("TABLE_NAME");
            tables.add(tableName);
        }

        for(String tableName:tables){
            XSSFSheet sheet = book.createSheet(tableName);
            //获取所有列名
            //创建第一行
            XSSFRow row = sheet.createRow(0);
            XSSFCell cell1 = row.createCell(0);
            cell1.setCellValue("序号");
            XSSFCell cell2 = row.createCell(1);
            cell2.setCellValue("名称");
            XSSFCell cell3 = row.createCell(2);
            cell3.setCellValue("数据类型");
            XSSFCell cell4 = row.createCell(3);
            cell4.setCellValue("是否为空");
            XSSFCell cell5 = row.createCell(4);
            cell5.setCellValue("注释");


            ResultSet resultSet = st.executeQuery("show full columns from "+dbName+"."+tableName);
            int index=0;
            while (resultSet.next()) {
                index++;
                String field = resultSet.getString("Field");
                String type = resultSet.getString("Type");
                String comment = resultSet.getString("Comment");
                String isnull = resultSet.getString("Null");

                XSSFRow valueRow = sheet.createRow(index);
                //创建一个新的列
                XSSFCell cell00 = valueRow.createCell(0);
                cell00.setCellValue(index);
                //写入列名
                XSSFCell cell01 = valueRow.createCell(1);
                cell01.setCellValue(field);
                XSSFCell cell02 = valueRow.createCell(2);
                cell02.setCellValue(type);
                XSSFCell cell03 = valueRow.createCell(3);
                cell03.setCellValue(isnull);
                XSSFCell cell04 = valueRow.createCell(4);
                cell04.setCellValue(comment);
            }
        }
        con.close();
        book.write(new FileOutputStream("C:\\Users\\ZZH\\Desktop\\"+dbName+".xlsx"));

    }

}