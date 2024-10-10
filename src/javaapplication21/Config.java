/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package javaapplication21;

import java.sql.Connection;
import java.sql.DriverManager;
import javax.swing.JOptionPane;
import java.sql.*;

/**
 *
 * @author Dell
 */
public class Config {

    Connection conn;

    public Config() {
        try {

//            Class.forName("com.mysql.cj.jdbc.Driver");
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
            conn = DriverManager.getConnection("jdbc:sqlserver://DESKTOP-Q1R10TN:1433;databaseName=एमपीआर2;user=sa;password=admin@123456789;encrypt=true;trustServerCertificate=true");

//            conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/एमपीआर","root","admin@123456789");
            System.out.println("Successfully connected");
        } catch (Exception ex) {
            System.out.println(ex.toString());
            JOptionPane.showMessageDialog(null, "Database exception occured" + ex.toString(), "Pension Record System", JOptionPane.ERROR_MESSAGE);
        }
    }

}
