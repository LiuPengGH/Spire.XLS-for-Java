package com.mysqlToExecl;


import com.roadjava.util.DBUtile;
import com.spire.xls.*;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;

public class PointToExe {

    public static void main(String[] args) {


        List pointName = PointDLName.tests();
        String[] strings = {"管点编号", "X坐标", "Y坐标", "特征", "附属物", "地面高程", "埋深", "井盖类型", "井盖规格", "井盖材质", "阀门型号", "权属", "备注"};

        for (int i = 0; i < pointName.size(); i++) {

            System.out.println(pointName.get(i));
            Connection conn = null;
            PreparedStatement ps = null;
            ResultSet resultSet = null;
            //定义一个sql 查询出所有数据
            String sql = "SELECT  管点编号,X坐标,Y坐标,特征,附属物,地面高程,井底埋深  as 埋深,井盖类型,井盖规格,井盖材质,'' as 阀门型号,管线范围类型 as 权属,备注 FROM `txjs_point` WHERE 道路名称 IS NOT NULL and" +
                    " 道路名称 = '" + pointName.get(i) + "'";
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

                while (resultSet.next()) {

                    for (int in = 0; in < strings.length; in++) {
//                        System.out.println(resultSet.getRow());
                        CellRange cell1 = emptySheet.getCellRange(resultSet.getRow() + 2, in + 1);

                        cell1.setValue(resultSet.getString(strings[in]));
                    }

                }
                resultSet.last();
                CellRange cell = emptySheet.getCellRange(1, 1);
                cell.setValue(pointName.get(i) + "自来水管线探测成果点表(共" + resultSet.getRow() + "点）");
                resultSet.first();
                for (int in = 0; in < strings.length; in++) {
                    CellRange cell1 = emptySheet.getCellRange(2, in + 1);
                    cell1.setValue(strings[in]);
                }

                //保存文档
                System.out.println("保存成功++++++++++++++++++++++++++++++++++++++++++");
                wb1.saveToFile("D://xbcg//point//" + pointName.get(i) + "--点.xlsx", FileFormat.Version2013);

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
