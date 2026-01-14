package main;

import org.apache.commons.csv.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;

public class Main {

    private static final String INPUT_FILE = "D://tool/05012026/Acc-Manh05012026.csv";
    private static final String OUTPUT_DIR = "D://tool/05012026/";
    private static final String SUP_ID1 = "Sub_id1";
    private static final String SUP_ID2 = "Sub_id2";
    private static final String SUP_ID3 = "Sub_id3";
    private static final String SUP_ID4 = "Sub_id4";

    private static final String COMMISSION_COL = "Tổng hoa hồng đơn hàng(₫)";

    public static void main(String[] args) throws Exception {
        process();
        System.out.println("DONE!");
    }

    public static void process() throws Exception {

        Files.createDirectories(Paths.get(OUTPUT_DIR));

        // supId2 -> (supId3 -> stat)
        Map<String, Map<String, Stat>> data = new HashMap<>();

        try (
                Reader reader = Files.newBufferedReader(Paths.get(INPUT_FILE), StandardCharsets.UTF_8);
                CSVParser parser = CSVFormat.DEFAULT
                        .withFirstRecordAsHeader()
                        .withIgnoreHeaderCase()
                        .withTrim()
                        .parse(reader)
        ) {
            for (CSVRecord r : parser) {
                String supId2 = safe(r.get(SUP_ID2)).toLowerCase();
                String supId3 = safe(r.get(SUP_ID4)).toLowerCase();
                if ("TranChauDuongDen".contains(supId2)) {
                	supId2 =  safe(r.get(SUP_ID1)).toLowerCase();
                	supId3 =  safe(r.get(SUP_ID3)).toLowerCase();
				}
                BigDecimal money = parseMoney(r.get(COMMISSION_COL));

                data
                        .computeIfAbsent(supId2.toLowerCase(), k -> new HashMap<>())
                        .computeIfAbsent(supId3, k -> new Stat())
                        .add(money);
            }
        }

        // xuất từng file supId2
        for (String supId2 : data.keySet()) {
            exportExcelBySupId2(supId2, data.get(supId2));
        }

        // xuất file tổng
        exportSummaryAllSupId2(data);
    }

    // =====================================================
    // FILE THEO TỪNG SUP_ID2
    // =====================================================
    private static void exportExcelBySupId2(String supId2, Map<String, Stat> map) throws Exception {

        Workbook wb = new XSSFWorkbook();

        // Sheet 1: Tổng hoa hồng
        Sheet totalSheet = wb.createSheet("Tong_HoaHong");
        totalSheet.createRow(0).createCell(0).setCellValue("Nhóm 2 | " + supId2);

        BigDecimal totalMoney = BigDecimal.ZERO;
        int totalOrders = 0;

        for (Stat s : map.values()) {
            totalMoney = totalMoney.add(s.total);
            totalOrders += s.count;
        }

        totalSheet.createRow(2).createCell(0).setCellValue("Tổng số đơn");
        totalSheet.getRow(2).createCell(1).setCellValue(totalOrders);

        totalSheet.createRow(3).createCell(0).setCellValue("Tổng hoa hồng (VNĐ)");
        totalSheet.getRow(3).createCell(1).setCellValue(formatMoney(totalMoney));

        // Sheet 2: Chi tiết
        Sheet detail = wb.createSheet("HoaHong_Theo_Sub_id3");

        Row h = detail.createRow(0);
        h.createCell(0).setCellValue("Sub_id3");
        h.createCell(1).setCellValue("Số đơn");
        h.createCell(2).setCellValue("Hoa hồng");

        List<Map.Entry<String, Stat>> list = new ArrayList<>(map.entrySet());
        list.sort((a, b) -> b.getValue().total.compareTo(a.getValue().total));

        int rowIdx = 1;
        for (Map.Entry<String, Stat> e : list) {
            Row r = detail.createRow(rowIdx++);
            r.createCell(0).setCellValue(e.getKey());
            r.createCell(1).setCellValue(e.getValue().count);
            r.createCell(2).setCellValue(formatMoney(e.getValue().total));
        }

        for (int i = 0; i < 3; i++) detail.autoSizeColumn(i);

        try (OutputStream os = Files.newOutputStream(Paths.get(OUTPUT_DIR, supId2 + ".xlsx"))) {
            wb.write(os);
        }
        wb.close();
    }

    // =====================================================
    // FILE TỔNG TẤT CẢ SUP_ID2 (CÓ GRAND TOTAL CUỐI FILE)
    // =====================================================
    private static void exportSummaryAllSupId2(Map<String, Map<String, Stat>> data) throws Exception {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Tong_Hop_SupId2");

        Row h = sheet.createRow(0);
        h.createCell(0).setCellValue("Sup_id2");
        h.createCell(1).setCellValue("Tổng số đơn");
        h.createCell(2).setCellValue("Tổng hoa hồng");

        List<RowSummary> list = new ArrayList<>();

        BigDecimal grandTotalMoney = BigDecimal.ZERO;
        int grandTotalOrders = 0;

        for (String supId2 : data.keySet()) {
            int orders = 0;
            BigDecimal money = BigDecimal.ZERO;

            for (Stat s : data.get(supId2).values()) {
                orders += s.count;
                money = money.add(s.total);
            }

            grandTotalOrders += orders;
            grandTotalMoney = grandTotalMoney.add(money);

            list.add(new RowSummary(supId2, orders, money));
        }

        // sort giảm dần theo hoa hồng
        list.sort((a, b) -> b.total.compareTo(a.total));

        int rowIdx = 1;
        for (RowSummary r : list) {
            Row row = sheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(r.supId2);
            row.createCell(1).setCellValue(r.orders);
            row.createCell(2).setCellValue(formatMoney(r.total));
        }

        // ===== DÒNG GRAND TOTAL =====
        Row totalRow = sheet.createRow(rowIdx);
        totalRow.createCell(0).setCellValue("TỔNG TẤT CẢ");
        totalRow.createCell(1).setCellValue(grandTotalOrders);
        totalRow.createCell(2).setCellValue(formatMoney(grandTotalMoney));

        for (int i = 0; i < 3; i++) sheet.autoSizeColumn(i);

        try (OutputStream os = Files.newOutputStream(Paths.get(OUTPUT_DIR, "TONG_HOA_HONG_ALL_SUPID2.xlsx"))) {
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

    // =====================================================
    // MODEL
    // =====================================================
    static class Stat {
        int count = 0;
        BigDecimal total = BigDecimal.ZERO;

        void add(BigDecimal money) {
            count++;
            total = total.add(money);
        }
    }

    static class RowSummary {
        String supId2;
        int orders;
        BigDecimal total;

        RowSummary(String supId2, int orders, BigDecimal total) {
            this.supId2 = supId2;
            this.orders = orders;
            this.total = total;
        }
    }
}
