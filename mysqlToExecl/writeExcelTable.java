package com.mysqlToExecl;


import com.roadjava.util.DBUtile;
import jxl.Workbook;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;

import java.io.*;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;

public class writeExcelTable {

    public static void main(String[] args) {


        List tests = LineName.tests();
        String liuxiang = null;

        for (int i =0;i < tests.size();i++){


            Connection conn = null;
            PreparedStatement ps = null;
            ResultSet resultSet = null;
            //定义一个sql 查询出所有数据
            String sql = "SELECT 起点点号,终点点号,起点高程,终点高程,材质,起点X,起点Y,终点X,终点Y,起点埋深,终点埋深,埋设方式,管径,使用状况,道路名称,管线类型,流向,备注 FROM `txjs_line` WHERE 道路名称 IS NOT NULL and " +
                    " 道路名称 = '" + tests.get(i) + "'";

            try {


                conn = DBUtile.getConn();
                if (conn == null) {
                }
                ps = conn.prepareStatement(sql);
                resultSet = ps.executeQuery();

                WritableWorkbook wwb = null;
                //创建可写入的Excel工作薄











                String fileName = "D://xbcg//sa//xian2.xls";
                File file = new File(fileName);


                //以fileName 为文件名来创建一个Workbook
                wwb = Workbook.createWorkbook(file);
                wwb.removeSheet(4);
                //创建工作表
                WritableSheet ws = wwb.getSheet( 0);
                //查询数据库中所有的数据

                resultSet.last();
                Label label = new Label(0, 0, tests.get(i)+"自来水管线探测成果线表(共"+resultSet.getRow()+"段）");

                System.out.println("最后"+ resultSet.getRow());
                resultSet.first();
                System.out.println("最先"+ resultSet.getRow());


                WritableCellFormat wcf = new WritableCellFormat();
                WritableFont wf = new WritableFont(WritableFont.createFont("宋体"),11,WritableFont.BOLD);
                wcf.setFont(wf);
                Label title = new Label(0,0,"title",wcf);
                ws.addCell(title);


                //要插入到的Excel表格的行号，默认从0开始
                Label e0 = new Label(0, 1, "起点点号");
                Label e1 = new Label(1, 1, "终点点号");
                Label e2 = new Label(2, 1, "起点高程");
                Label e3 = new Label(3, 1, "终点高程");
                Label e4=new Label(4, 1, "材质");
                Label e5=new Label(5, 1, "起点X");
                Label e6=new Label(6, 1, "起点Y");
                Label e7=new Label(7, 1, "终点X");
                Label e8=new Label(8, 1, "终点Y");
                Label e9=new Label(9, 1, "起点埋深");
                Label e10=new Label(10, 1, "终点埋深");
                Label e11=new Label(11, 1, "埋设方式");
                Label e12=new Label(12, 1, "管径");
                Label e13=new Label(13, 1, "使用状况");
                Label e14=new Label(14, 1, "道路名称");
                Label e15=new Label(15, 1, "管线类型");
                Label e16=new Label(16, 1, "流向");
                Label e17=new Label(17, 1, "备注");

                ws.addCell(e0);
                ws.addCell(e1);
                ws.addCell(e2);
                ws.addCell(e3);
                ws.addCell(e4);
                ws.addCell(e5);
                ws.addCell(e6);
                ws.addCell(e7);
                ws.addCell(e8);
                ws.addCell(e9);
                ws.addCell(e10);
                ws.addCell(e11);
                ws.addCell(e12);
                ws.addCell(e13);
                ws.addCell(e14);
                ws.addCell(e15);
                ws.addCell(e16);
                ws.addCell(e17);
                ws.addCell(label);

                while (resultSet.next()) {

                    String name = resultSet.getString("道路名称");
                    System.out.println(name);
                    int row = resultSet.getRow();
                    System.out.println("当前行数：  " + row);
//                    if (resultSet.getString("流向").equals("0")|| resultSet.getString("流向").isEmpty()){
//                        liuxiang = "逆流";
//                    }else liuxiang = "顺流";


                    Label l0 = new Label(0, row, resultSet.getString("起点点号"));
                    Label l1 = new Label(1, row, resultSet.getString("终点点号"));
                    Label l2 = new Label(2, row, resultSet.getString("起点高程"));
                    Label l3 = new Label(3, row, resultSet.getString("终点高程"));
                    Label l4 = new Label(4, row, resultSet.getString("材质"));
                    Label l5 = new Label(5, row, resultSet.getString("起点X"));
                    Label l6 = new Label(6, row, resultSet.getString("起点Y"));
                    Label l7 = new Label(7, row, resultSet.getString("终点X"));
                    Label l8 = new Label(8, row, resultSet.getString("终点Y"));
                    Label l9 = new Label(9, row, resultSet.getString("起点埋深"));
                    Label l10 = new Label(10, row, resultSet.getString("终点埋深"));
                    Label l11 = new Label(11, row, resultSet.getString("埋设方式"));
                    Label l12 = new Label(12, row, resultSet.getString("管径"));
                    Label l13 = new Label(13, row, resultSet.getString("使用状况"));
                    Label l14 = new Label(14, row, resultSet.getString("道路名称"));
                    Label l15 = new Label(15, row, resultSet.getString("管线类型"));
                    Label l16 = new Label(16, row, resultSet.getString("流向"));
                    Label l17 = new Label(17, row, resultSet.getString("备注"));


                    ws.addCell(l0);
                    ws.addCell(l1);
                    ws.addCell(l2);
                    ws.addCell(l3);
                    ws.addCell(l4);
                    ws.addCell(l5);
                    ws.addCell(l6);
                    ws.addCell(l7);
                    ws.addCell(l8);
                    ws.addCell(l9);
                    ws.addCell(l10);
                    ws.addCell(l11);
                    ws.addCell(l12);
                    ws.addCell(l13);
                    ws.addCell(l14);
                    ws.addCell(l15);
                    ws.addCell(l16);
                    ws.addCell(l17);

                }
                //写进文档
                wwb.write();
                System.out.println("数据写入成功");

                //关闭Excel工作簿对象
                wwb.close();


            } catch (SQLException throwables) {
                throwables.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (RowsExceededException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            } finally {
                DBUtile.clossRs(resultSet);
                DBUtile.closeConn(conn);
                DBUtile.clossPs(ps);
            }


        }


//
//        try {
//            WritableWorkbook wwb=null;
//            //创建可写入的Excel工作薄
//
//            String fileName="D://book.xls";
//            File file=new File(fileName);
//
//            if(!file.exists()) {
//                file.createNewFile();
//            }
//            //以fileName 为文件名来创建一个Workbook
//            wwb= Workbook.createWorkbook(file);
//            //创建工作表
//            WritableSheet ws=wwb.createSheet("Test Shee 1", 0);
//
//            //查询数据库中所有的数据
//            List<StuEntity> list=StuService.getAllByDb();
//            //要插入到的Excel表格的行号，默认从0开始
//
//            Label labelId=new Label(0, 0, "编号（id）");
//            Label labelName=new Label(1, 0, "姓名（name）");
//            Label labelSex=new Label(2, 0, "性别(sex)");
//            Label labelNum=new Label(3, 0, "薪水（num）");
//
//            ws.addCell(labelId);
//            ws.addCell(labelName);
//            ws.addCell(labelSex);
//            ws.addCell(labelNum);
//
//            for (int i = 0; i < list.size(); i++) {
//                Label labelId_i=new Label(0, i+1, list.get(i).getId()+"");
//                Label labelName_i=new Label(1, i+1, list.get(i).getName());
//                Label labelSex_i=new Label(2, i+1, list.get(i).getSex());
//                Label labelNum_i=new Label(3, i+1, list.get(i).getNum()+"");
//                ws.addCell(labelId_i);
//                ws.addCell(labelName_i);
//                ws.addCell(labelSex_i);
//                ws.addCell(labelNum_i);
//            }
//
//            //写进文档
//            wwb.write();
//            System.out.println("数据写入成功");
//            //关闭Excel工作簿对象
//
//            wwb.close();
//
//        } catch (Exception e) {
//
//            System.out.println("数据写入失败");
//            e.printStackTrace();
//        }
    }
    public static File copyFile(File source,String dest )throws IOException{
        //创建目的地文件夹
        File destfile = new File(dest);
        if(!destfile.exists()){
            destfile.mkdir();
        }
        //如果source是文件夹，则在目的地址中创建新的文件夹
        if(source.isDirectory()){
            File file = new File(dest+"\\"+source.getName());//用目的地址加上source的文件夹名称，创建新的文件夹
            file.mkdir();
            //得到source文件夹的所有文件及目录
            File[] files = source.listFiles();
            if(files.length==0){
                return file;
            }else{
                for(int i = 0 ;i<files.length;i++){
                    copyFile(files[i],file.getPath());
                }
            }
            return file;
        }
        //source是文件，则用字节输入输出流复制文件
        else if(source.isFile()){
            FileInputStream fis = new FileInputStream(source);
            //创建新的文件，保存复制内容，文件名称与源文件名称一致
            File dfile = new File(dest+"\\"+source.getName());
            if(!dfile.exists()){
                dfile.createNewFile();
            }

            FileOutputStream fos = new FileOutputStream(dfile);
            // 读写数据
            // 定义数组
            byte[] b = new byte[1024];
            // 定义长度
            int len;
            // 循环读取
            while ((len = fis.read(b))!=-1) {
                // 写出数据
                fos.write(b, 0 , len);
            }

            //关闭资源
            fos.close();
            fis.close();
            return dfile;
        }
        return null;
    }



}
