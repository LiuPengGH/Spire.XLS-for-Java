package com.mysqlToExecl;

import com.roadjava.util.DBUtile;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public class LineName {


    public static List tests() {


        Connection conn = null;
        PreparedStatement ps = null;
        ResultSet resultSet = null;
        //定义一个sql 查询出所有数据
        String sql = "select 道路名称 FROM `txjs_line` WHERE 道路名称 IS NOT NULL group by 道路名称";
        List<String> daoList = new ArrayList<>();

        try {


            conn = com.roadjava.util.DBUtile.getConn();
            if (conn == null) {
            }
            ps = conn.prepareStatement(sql);
            resultSet = ps.executeQuery();
            while (resultSet.next()){
                daoList.add(resultSet.getString("道路名称"));
            }
            return daoList;

        } catch (SQLException throwables) {
            throwables.printStackTrace();
            return null;
        } finally {
            com.roadjava.util.DBUtile.clossRs(resultSet);
            com.roadjava.util.DBUtile.closeConn(conn);
            DBUtile.clossPs(ps);
        }


    }
}
