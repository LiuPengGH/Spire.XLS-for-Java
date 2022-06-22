package com.mysqlToExecl;




import com.roadjava.util.DBUtile;
import com.spire.xls.*;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;

public class LineTableToExe {

    public static void main(String[] args) {



        List tests = LineName.tests();
        String [] strings = {"起点点号","终点点号","起点高程","终点高程","材质","起点X","起点Y","终点X","终点Y","起点埋深","终点埋深","埋设方式","管径","使用状况","道路名称","管线类型","流向","备注"};

        for (int i =0;i < tests.size();i++) {


            System.out.println(tests.get(i));
            Connection conn = null;
            PreparedStatement ps = null;
            ResultSet resultSet = null;
            //定义一个sql 查询出所有数据
            String sql = "SELECT 起点点号,终点点号,起点高程,终点高程,材质,起点X,起点Y,终点X,终点Y,起点埋深,终点埋深,埋设方式,管径,使用状况,道路名称,管线范围类型,REPLACE(REPLACE(lx,'1','逆流'),'0','顺流') as 流向,备注  FROM `txjs_line` WHERE 道路名称 IS NOT NULL and" +
                    " 道路名称 = '" + tests.get(i) + "'";
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
                CellStyle style2 = wb1.getStyles().addStyle("n");
                //设置第一列


                while (resultSet.next()){

                    for (int in = 0;in<strings.length;in ++){
                        System.out.println(resultSet.getRow());
                        CellRange cell1 = emptySheet.getCellRange(resultSet.getRow()+2, in+1);
                        if (strings[in].equals("管线类型")){
                            cell1.setValue(resultSet.getString("管线范围类型"));
                        }else
                            cell1.setValue(resultSet.getString(strings[in]));


                    }

                }


                resultSet.last();
                CellRange cell = emptySheet.getCellRange(1, 1);
                cell.setValue(tests.get(i)+"自来水管线探测成果线表(共"+resultSet.getRow()+"段）");
                resultSet.first();
                for (int in = 0;in<strings.length;in ++){
                    CellRange cell1 = emptySheet.getCellRange(2, in+1);
                    cell1.setValue(strings[in]);

                }

//保存文档

                System.out.println("保存成功++++++++++++++++++++++++++++++++++++++++++");
                wb1.saveToFile("D://xbcg//line//"+tests.get(i)+"--线.xlsx", FileFormat.Version2013);






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
