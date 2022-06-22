package com.mysqlToExecl;

import com.roadjava.util.DBUtile;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public class SBName {


    public static List tests() {


        Connection conn = null;
        PreparedStatement ps = null;
        ResultSet resultSet = null;
        //定义一个sql 查询出所有数据
        String sql = "SELECT XQNAME FROM `txwatermeter` GROUP BY XQNAME";
        List<String> daoList = new ArrayList<>();

        try {


            conn = DBUtile.getConn();
            if (conn == null) {
            }
            ps = conn.prepareStatement(sql);
            resultSet = ps.executeQuery();
            while (resultSet.next()){
                daoList.add(resultSet.getString("XQNAME"));
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
