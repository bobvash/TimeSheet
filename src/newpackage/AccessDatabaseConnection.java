/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package newpackage;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
 
/**
 *
 * @author Jsupport
 */
public class AccessDatabaseConnection {
 
    public static Connection connect() {
 
        Connection connection = null;
        try {
            Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
            String path = "C:\\Documents and Settings\\sesa245615\\My Documents\\Dropbox\\NetBeansProjects\\TimeSheet\\";
            String database = "jdbc:odbc:Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="+path+"Projects.accdb";
//          connection = DriverManager.getConnection("jdbc:odbc:Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="+path+"Projects.accdb","","");
            connection = DriverManager.getConnection( database ,"",""); 
 
        } catch (Exception ex) {
            System.out.println(ex);
            return null;
        }
        return connection;
    }

    static void AddWorkedHours(Date stopTime, Object ProjName, Object EsNum, Double diff) {
        try {
            Connection connection = connect();
            Statement stmt = connection.createStatement();
            ResultSet rset = stmt.executeQuery("SELECT * FROM HoursWorked");

            if (rset.next()) {
                    System.out.println(rset.getObject("DateWorked"));                   
            }
 
        } catch (SQLException ex) {
             System.out.println(ex);
        }
        
        try {
            Connection connection = connect();
//            Statement stmt = connection.createStatement();
//            stmt.executeUpdate("INSERT INTO HoursWorked (DateWorked, Hours) VALUES ('" + stopTime.toString() + "', " + diff + ")");

            Calendar beginTime = Calendar.getInstance();
            beginTime.setTime(stopTime);
            beginTime.set(Calendar.HOUR_OF_DAY,0);
            beginTime.set(Calendar.MINUTE, 0);
            beginTime.set(Calendar.SECOND, 0);
            beginTime.set(Calendar.MILLISECOND, 0);
            
            Calendar endTime = Calendar.getInstance();
            endTime.setTime(stopTime);
            endTime.set(Calendar.HOUR_OF_DAY,0);
            endTime.set(Calendar.MINUTE, 0);
            endTime.set(Calendar.SECOND, 0);
            endTime.set(Calendar.MILLISECOND, 0);
            endTime.add(Calendar.HOUR_OF_DAY, 24);
            
            PreparedStatement checkExist = connection.prepareStatement("SELECT * FROM HoursWorked WHERE ES =? AND DateWorked >= ? AND DateWorked < ?");
            checkExist.setObject(1, EsNum);
//            exist.setTimestamp(2, new java.sql.Timestamp(beginTime.
//                    getTimeInMillis()));
//            exist.setTimestamp(3, new java.sql.Timestamp(endTime.
//                    getTimeInMillis()));
            checkExist.setDate(2, new java.sql.Date(beginTime.
                    getTimeInMillis()));
            checkExist.setDate(3, new java.sql.Date(endTime.
                    getTimeInMillis()));
            ResultSet rset = checkExist.executeQuery();
            
            if (rset.next()) {
//                Double temp = rset.getDouble("Hours");
//                rset.updateDouble("Hours", temp+diff);
//                
                Double rowHours = rset.getDouble("Hours");
                Integer rowId = rset.getInt("Id");
                PreparedStatement exist = connection.prepareStatement("UPDATE HoursWorked SET Hours = ? WHERE Id = ?");
                exist.setDouble(1, rowHours + diff);
                exist.setInt(2, rowId);
                
                exist.executeQuery();                
            }
            else {
                PreparedStatement nonExist = connection.prepareStatement("INSERT INTO HoursWorked (DateWorked, Hours, ES, ProjectName) VALUES (?,?,?,?)");
                nonExist.setDate(1, new java.sql.Date(stopTime.getTime()));
                nonExist.setDouble(2, diff);
                nonExist.setObject(3, EsNum);
                nonExist.setObject(4, ProjName);

                nonExist.executeQuery();
            }

        } catch (SQLException ex) {
             System.out.println(ex);
        }
    }


    public Object[][] getData(){
        ArrayList<Object[]> dataRows = new ArrayList<Object[]>();
        try {
            Connection connection = connect();
            Statement stmt = connection.createStatement();
            ResultSet rset = stmt.executeQuery("SELECT * FROM project");

            if (rset.next()) {
                Object[] row = {
                    rset.getString("Q2C"),
                    rset.getString("ES"),
                    rset.getString("ProjectName"),
                    rset.getInt("HoursWorked"),
                    rset.getInt("HoursQuoted"),
                    rset.getString("Status"),
                    createScopeList(rset.getBoolean("isSC"), rset.getBoolean("isTCC"), rset.getBoolean("isAF"),rset.getBoolean("isAFBdyLbl"))
                };
                
                dataRows.add(row);
            }
 
        } catch (SQLException ex) {
             System.out.println(ex);
        }
        return dataRows.toArray(new Object[dataRows.size()][]);
    }
 
    public static void main(String[] args) {
        new AccessDatabaseConnection().getData();
    }
    
// Function to create a list of scopes in comma separated list

    private String createScopeList(boolean isSC, boolean isTCC, boolean isAF, boolean isAFBdyLbl) {
        String scope = "";
        if (isSC) {
                scope = scope+ "SC, ";
        }
        
        if (isTCC) {
                scope = scope + "TCC, ";
        }
        
        if (isAF) {
                scope = scope+ "AF, ";
        }
        
        if (isAFBdyLbl) {
                scope = scope + "AF Bdy Lbl, ";
        }
        
        scope = scope.substring(0,scope.length()-2);
        
        return scope;
    }

}


