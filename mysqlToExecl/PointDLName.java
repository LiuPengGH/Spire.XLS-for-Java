package com.mysqlToExecl;

import com.roadjava.util.DBUtile;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public class PointDLName {


    public static List tests() {


        Connection conn = null;
        PreparedStatement ps = null;
        ResultSet resultSet = null;
        //定义一个sql 查询出所有数据
        String sql = "select 道路名称 FROM `txjs_point` WHERE 道路名称 IS NOT NULL group by 道路名称";
        List<String> daoList = new ArrayList<>();

        try {


            conn = DBUtile.getConn();
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
            DBUtile.clossRs(resultSet);
            DBUtile.closeConn(conn);
            DBUtile.clossPs(ps);
        }


    }
}
