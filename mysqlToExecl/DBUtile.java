package com.mysqlToExecl;


import java.sql.*;

public class DBUtile {
    //设置mysql驱动和url
    static final String JDBC_DRIVER = "com.mysql.jdbc.Driver";
    static final String DB_URL = "jdbc:mysql://localhost:3306/sys?useSSL=false&allowPublicKeyRetrieval=true&serverTimezone=UTC";
    //设置用户名和密码
    static final String USER = "root";
    static final String PASS = "root00";


    static {
        try {
            Class.forName(JDBC_DRIVER);
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
    }

    public static Connection getConn(){
        try {
            return DriverManager.getConnection(DB_URL,USER,PASS);
        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }
        return null;
    }
    public static void closeConn(Connection connection){
        if (connection != null){
            try {
                connection.close();
            } catch (SQLException throwables) {
                throwables.printStackTrace();
            }
        }
    }
    public static void clossPs(PreparedStatement preparedStatement){
        try {
            preparedStatement.close();
        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }
    }
    public static void clossRs(ResultSet resultSet){
        try {
            resultSet.close();
        } catch (SQLException throwables) {
            throwables.printStackTrace();
        }
    }
}
