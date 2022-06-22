package com.mysqlToExecl;


import com.roadjava.util.DBUtile;
import com.spire.xls.CellRange;
import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;

public class SBToExe {

    public static void main(String[] args) {


        List pointName = SBName.tests();
        String[] strings = {"序号","点号", "户名", "户室名",  "经度", "纬度","普查情况", "备注"};

        for (int i = 0; i < pointName.size(); i++) {

            System.out.println(pointName.get(i));
            Connection conn = null;
            PreparedStatement ps = null;
            ResultSet resultSet = null;
            //定义一个sql 查询出所有数据
            String sql = "SELECT ROW_NUMBER() OVER(PARTITION BY XQNAME ORDER BY USERNO) as 序号,USERNO as 点号,DWNAME as 户名,ADDRESS as 户室名,X as 经度,Y as 纬度,'完成' as 普查情况,'无' as 备注 FROM `txwatermeter` WHERE XQNAME='" + pointName.get(i) + "'";
            System.out.println(sql);

            try {

                conn = DBUtile.getConn();
                if (conn == null) {
                }
                ps = conn.prepareStatement(sql);
                resultSet = ps.executeQuery();

                Workbook wb = new Workbook();
                Workbook wb1 = new Workbook();

                wb.loadFromFile("D://xbcg//111.xlsx");

                Worksheet sheet = wb.getWorksheets().get(0);


                //sheet.setName("Copiedsheet");
                Worksheet emptySheet = wb1.getWorksheets().get(0);
                emptySheet.copyFrom(sheet);
                //设置第一列
                resultSet.last();
                int lastRow = resultSet.getRow();
                if (lastRow < 650){
                    System.out.println("行数： " + lastRow + "-------------------------" );
                    continue;
                }

                CellRange cell = emptySheet.getCellRange(1, 1);
                if (lastRow ==1){
                    cell.setValue(pointName.get(i) + "水表普查自检表(共" + resultSet.getRow() + "点） 抽查（1个点）");
                }else {
                    cell.setValue(pointName.get(i) + "水表普查自检表(共" + resultSet.getRow() + "点） 抽查（"+lastRow/2+"个点）");
                }
                resultSet.absolute(0);
                for (int in = 0; in < strings.length; in++) {
                    CellRange cell1 = emptySheet.getCellRange(2, in + 1);
                    cell1.setValue(strings[in]);
                }
                while (resultSet.next()) {
                    if (lastRow==1){
                        for (int in = 0; in < strings.length; in++) {
//                            System.out.println(resultSet.getRow());

                            CellRange cell1 = emptySheet.getCellRange(resultSet.getRow() + 2, in + 1);

                            cell1.setValue(resultSet.getString(strings[in]));
                        }
                    }
                    if (lastRow/2 >= resultSet.getRow()){
                        for (int in = 0; in < strings.length; in++) {
//                            System.out.println(resultSet.getRow());

                            CellRange cell1 = emptySheet.getCellRange(resultSet.getRow() + 2, in + 1);

                            cell1.setValue(resultSet.getString(strings[in]));
                        }
                    }

                }


                //保存文档
                System.out.println("保存成功++++++++++++++++++++++++++++++++++++++++++");
                wb1.saveToFile("D://xbcg//xiaoqu//" + pointName.get(i) + "-水表普查表.xlsx", FileFormat.Version2013);

            } catch (SQLException throwables) {
                throwables.printStackTrace();
            } finally {
                DBUtile.clossRs(resultSet);
                DBUtile.closeConn(conn);
                DBUtile.clossPs(ps);
            }


        }

    }
}
