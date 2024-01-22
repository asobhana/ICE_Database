import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

/**
 * @author ASOBHANA
 */





    public class ParseData {
        public static void main(String[] args) {
            String jdbcUrl = "jdbc:mysql://localhost:3306/DB1;databaseName=mydatabase";
            String username = "asobhana";
            String password = "asobhana";

            try (Connection connection = DriverManager.getConnection(jdbcUrl, username, password)) {
                System.out.println("Connected to the database!");
                // Add your database operations here

            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }




