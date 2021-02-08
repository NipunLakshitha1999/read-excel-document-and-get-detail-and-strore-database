import java.sql.Connection;
import java.sql.Driver;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DBConnection {

    private static DBConnection dbConnection;
    private Connection connection;

    private DBConnection() throws ClassNotFoundException, SQLException{
        Class.forName("com.mysql.cj.jdbc.Driver");
         connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/excelData", "root", "root");
    }

    public static DBConnection getDbConnection() throws ClassNotFoundException,SQLException{
        if (dbConnection==null){
            dbConnection=new DBConnection();
        }
        return dbConnection;
    }
    public Connection getConnection(){
        return connection;
    }
}
