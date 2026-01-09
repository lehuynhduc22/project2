package main;

import com.sun.net.httpserver.*;
import org.apache.commons.csv.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.net.InetSocketAddress;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;

public class Main {

    private static final Path UPLOAD_DIR = Paths.get("uploads");
    private static final Path OUTPUT_DIR = Paths.get("outputs");

    private static final String SUP_ID1 = "Sub_id1";
    private static final String SUP_ID2 = "Sub_id2";
    private static final String SUP_ID3 = "Sub_id3";
    private static final String SUP_ID4 = "Sub_id4";
    private static final String COMMISSION_COL = "Tổng hoa hồng đơn hàng(₫)";

    public static void main(String[] args) throws Exception {

        Files.createDirectories(UPLOAD_DIR);
        Files.createDirectories(OUTPUT_DIR);

        HttpServer server = HttpServer.create(new InetSocketAddress(8080), 0);

        // ===== HOME PAGE =====
        server.createContext("/", exchange -> {
            String html =
                "<html>" +
                "<body>" +
                "<h2>Upload CSV</h2>" +
                "<form method='post' enctype='multipart/form-data' action='/upload'>" +
                "<input type='file' name='file'/>" +
                "<br/><br/>" +
                "<button>Process</button>" +
                "</form>" +
                "</body>" +
                "</html>";

            byte[] bytes = html.getBytes("UTF-8");
            exchange.sendResponseHeaders(200, bytes.length);
            exchange.getResponseBody().write(bytes);
            exchange.close();
        });

        // ===== UPLOAD HANDLER =====
        server.createContext("/upload", exchange -> {
            if (!exchange.getRequestMethod().equalsIgnoreCase("POST")) {
                exchange.sendResponseHeaders(405, -1);
                return;
            }

            Path csvFile = UPLOAD_DIR.resolve("input.csv");

            try (InputStream is = exchange.getRequestBody()) {
                Files.copy(is, csvFile, StandardCopyOption.REPLACE_EXISTING);
            }

            try {
				process(csvFile);
			} catch (Exception e) {
				e.printStackTrace();
			}

            String resp =
            		"<html>\r\n"
            		+ "                <body>\r\n"
            		+ "                    <h3>DONE</h3>\r\n"
            		+ "                    <a href=\"/download\">Download Result</a>\r\n"
            		+ "                </body>\r\n"
            		+ "                </html>";

            exchange.sendResponseHeaders(200, resp.getBytes().length);
            exchange.getResponseBody().write(resp.getBytes());
            exchange.close();
        });

        // ===== DOWNLOAD =====
        server.createContext("/download", exchange -> {
            Path file = OUTPUT_DIR.resolve("TONG_HOA_HONG_ALL_SUPID2.xlsx");
            byte[] data = Files.readAllBytes(file);

            exchange.getResponseHeaders().add("Content-Disposition",
                    "attachment; filename=" + file.getFileName());
            exchange.sendResponseHeaders(200, data.length);
            exchange.getResponseBody().write(data);
            exchange.close();
        });

        server.start();
        System.out.println("Server running on port 8080");
    }

    // =====================================================
    // PROCESS CSV → EXCEL (LOGIC GỐC CỦA BẠN)
    // =====================================================
    public static void process(Path inputCsv) throws Exception {

        Map<String, Map<String, Stat>> data = new HashMap<>();

        try (
                Reader reader = Files.newBufferedReader(inputCsv, StandardCharsets.UTF_8);
                CSVParser parser = CSVFormat.DEFAULT
                        .withFirstRecordAsHeader()
                        .withIgnoreHeaderCase()
                        .withTrim()
                        .parse(reader)
        ) {
            for (CSVRecord r : parser) {
                String supId2 = safe(r.get(SUP_ID2));
                String supId3 = safe(r.get(SUP_ID4));

                if ("TranChauDuongDen".contains(supId2)) {
                    supId2 = safe(r.get(SUP_ID1));
                    supId3 = safe(r.get(SUP_ID3));
                }

                BigDecimal money = parseMoney(r.get(COMMISSION_COL));

                data.computeIfAbsent(supId2.toLowerCase(), k -> new HashMap<>())
                        .computeIfAbsent(supId3, k -> new Stat())
                        .add(money);
            }
        }

        exportSummaryAllSupId2(data);
    }

    // =====================================================
    // EXPORT SUMMARY FILE
    // =====================================================
    private static void exportSummaryAllSupId2(Map<String, Map<String, Stat>> data) throws Exception {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Tong_Hop_SupId2");

        Row h = sheet.createRow(0);
        h.createCell(0).setCellValue("Sup_id2");
        h.createCell(1).setCellValue("Tổng số đơn");
        h.createCell(2).setCellValue("Tổng hoa hồng");

        int rowIdx = 1;
        BigDecimal grandTotal = BigDecimal.ZERO;
        int grandOrders = 0;

        for (String supId2 : data.keySet()) {
            int orders = 0;
            BigDecimal total = BigDecimal.ZERO;

            for (Stat s : data.get(supId2).values()) {
                orders += s.count;
                total = total.add(s.total);
            }

            Row r = sheet.createRow(rowIdx++);
            r.createCell(0).setCellValue(supId2);
            r.createCell(1).setCellValue(orders);
            r.createCell(2).setCellValue(formatMoney(total));

            grandOrders += orders;
            grandTotal = grandTotal.add(total);
        }

        Row totalRow = sheet.createRow(rowIdx);
        totalRow.createCell(0).setCellValue("TỔNG TẤT CẢ");
        totalRow.createCell(1).setCellValue(grandOrders);
        totalRow.createCell(2).setCellValue(formatMoney(grandTotal));

        for (int i = 0; i < 3; i++) sheet.autoSizeColumn(i);

        try (OutputStream os = Files.newOutputStream(
                OUTPUT_DIR.resolve("TONG_HOA_HONG_ALL_SUPID2.xlsx"))) {
            wb.write(os);
        }
        wb.close();
    }

    // =====================================================
    // UTIL
    // =====================================================
    private static BigDecimal parseMoney(String raw) {
        if (raw == null || raw.isBlank()) return BigDecimal.ZERO;
        return new BigDecimal(raw.replace("đ", "").replace(",", "").trim());
    }

    private static String formatMoney(BigDecimal v) {
        return String.format("%,.0f đ", v);
    }

    private static String safe(String v) {
        return (v == null || v.isBlank()) ? "UNKNOWN" : v.trim();
    }

    static class Stat {
        int count = 0;
        BigDecimal total = BigDecimal.ZERO;

        void add(BigDecimal money) {
            count++;
            total = total.add(money);
        }
    }
}
