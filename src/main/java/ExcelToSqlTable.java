import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class ExcelToSqlTable {

    public static void main(String[] args) {
        String excelFilePath = "src/main/data/ICE.xlsx";
        String jdbcUrl = "jdbc:sqlserver://DESKTOP-3MO16VJ//SQLEXPRESS";
        String username = "sa";
        String password = "asobhana";

        try (Connection connection = DriverManager.getConnection(jdbcUrl, username, password);
             FileInputStream excelFile = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(excelFile)) {
             Sheet sheet = workbook.getSheetAt(0);
             createTable(connection);
             insertData(connection, sheet);
             System.out.println("Data imported successfully.");

        } catch ( Exception e) {
            e.printStackTrace();
        }
    }

    private static void createTable(Connection connection) throws SQLException {
        String createTableSQL = "CREATE TABLE IF NOT EXISTS your_table ("
                + "TRADE_DATE VARCHAR(255),"
                + "HUB VARCHAR(255),"
                + "PRODUCT VARCHAR(255),"
                + "STRIP VARCHAR(255),"
                + "CONTRACT VARCHAR(255),"
                + "CONTRACT_TYPE VARCHAR(255),"
                + "STRIKE VARCHAR(255),"
                + "SETTLEMENT_PRICE VARCHAR(255),"
                + "NET_CHANGE VARCHAR(255),"
                + "EXPIRATION_DATE VARCHAR(255),"
                + "PRODUCT_ID VARCHAR(255),"
                + ")";
        try (PreparedStatement preparedStatement = connection.prepareStatement(createTableSQL)) {
            preparedStatement.executeUpdate();
        }
    }

    private static void insertData(Connection connection, Sheet sheet) throws SQLException {
        String insertSQL = "INSERT INTO your_table (TRADE_DATE, HUB, PRODUCT, STRIP, CONTRACT, CONTRACT_TYPE, STRIKE, SETTLEMENT_PRICE, NET_CHANGE, EXPIRATION_DATE, PRODUCT_ID) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

        try (PreparedStatement preparedStatement = connection.prepareStatement(insertSQL)) {
            for (Row row : sheet) {
                // Skip header row
                if (row.getRowNum() == 0) {
                    continue;
                }

                preparedStatement.setString(1, getStringValue(row.getCell(0)));
                preparedStatement.setString(2, getStringValue(row.getCell(1)));
                preparedStatement.setString(3, getStringValue(row.getCell(2)));
                preparedStatement.setString(4, getStringValue(row.getCell(3)));
                preparedStatement.setString(5, getStringValue(row.getCell(4)));
                preparedStatement.setString(6, getStringValue(row.getCell(5)));
                preparedStatement.setString(7, getStringValue(row.getCell(6)));
                preparedStatement.setString(8, getStringValue(row.getCell(7)));
                preparedStatement.setString(9, getStringValue(row.getCell(8)));
                preparedStatement.setString(10, getStringValue(row.getCell(9)));
                preparedStatement.setString(11, getStringValue(row.getCell(10)));
                preparedStatement.setString(12, getStringValue(row.getCell(11)));
                preparedStatement.executeUpdate();
            }
        }
    }

    private static String getStringValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return null;
        }
    }
}
