package ru.statics;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.jdbc.core.JdbcTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.Scanner;

public class Profile {


    private final JdbcTemplate jdbc;

    public Profile(JdbcTemplate jdbc) {
        this.jdbc = jdbc;
    }

    public static void main(String[] args) throws SQLException {

        String sql = "trunc(cir.c_push_time, 'hh') time, ca.c_code abonent, cb_code bo, cb.c_name, cir.c_st, count(1) Количество сообщений\n" +
                "from z#cit_in_request cir\n" +
                "join z#cit_abonent ca on cir.c_abonent=ca.id\n" +
                "join z#cit_bo  cb on cir.c_bo=cb.id\n" +
                "where cir.c_push_time between to_date ('13/11/2023', 'dd/mm/yyyy' hh24) and\n" +
                "to_date ('13/12/2023', 'dd/mm/yyyy' hh24)\n" +
                "group by ca.c_code, cb.c_code, cb.c_name, trunc(cir.c_push_time, 'hh'), cir.c_st\n" +
                "order by time";
        String excelPath = "C:\\Профиль\\output.xlsx";
        String sqlExamle = "select * from idea where time between to_timestamp(?, 'YY-MM-DD HH24') and to_timestamp(?, 'YY-MM-DD HH24')";

        String url = "jdbc:postgresql://127.0.0.1:5432/idea";
        String username = "postgres";
        String password = "postgres";
        Scanner scanner = new Scanner(System.in);
        System.out.println("Введите начальную дату (YY-MM-DD HH24)");
        String startDate = scanner.nextLine();
        System.out.println("Введите конечную дату (YY-MM-DD HH24)");
        String endDate = scanner.nextLine();
        try (Connection connection = DriverManager.getConnection(url, username, password);
             PreparedStatement statement = connection.prepareStatement(sqlExamle)) {
            statement.setString(1, startDate);
            statement.setString(2, endDate);

            ResultSet resultSet = statement.executeQuery();
            writeExcel(resultSet, excelPath);
        }
    }

    private static void writeExcel(ResultSet resultSet, String excelPath) throws SQLException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Data");
        writeHeaderLine(resultSet, sheet);
        writeDataLines(resultSet, workbook, sheet);
        try (FileOutputStream out = new FileOutputStream(excelPath)) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void writeHeaderLine(ResultSet resultSet, Sheet sheet) throws SQLException {
        ResultSetMetaData metaData = resultSet.getMetaData();
        int numerOfColumns = metaData.getColumnCount();
        Row headerRow = sheet.createRow(0);
        for (int i = 1; i <= numerOfColumns; i++) {
            Cell headerCell = headerRow.createCell(i - 1);
            headerCell.setCellValue(metaData.getColumnName(i));
        }
    }

    private static void writeDataLines(ResultSet resultSet, Workbook workbook, Sheet sheet) throws SQLException {
        int rowCount = 1;
        while (resultSet.next()) {
            Row row = sheet.createRow(rowCount++);
            for (int i = 1; i <= resultSet.getMetaData()
                                          .getColumnCount(); i++) {
                Cell cell = row.createCell(i - 1);
                switch (resultSet.getMetaData()
                                 .getColumnType(i)) {
                    case Types.VARCHAR:
                        cell.setCellValue(resultSet.getString(i));
                        break;
                    case Types.INTEGER:
                        cell.setCellValue(resultSet.getInt(i));
                        break;
                    case Types.TIMESTAMP:
                        cell.setCellValue(resultSet.getTimestamp(i)
                                                   .toString());
                        break;
                    default:
                        cell.setCellValue(resultSet.getString(i));
                        break;
                }

            }


        }


    }
}
