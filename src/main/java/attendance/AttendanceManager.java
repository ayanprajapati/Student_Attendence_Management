package attendance;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * All Excel (.xlsx) work lives here.
 *
 * Workbook layout
 * ---------------
 * "Students"  : ID | Name | Branch                                  (master register)
 *
 * DAY-WISE tracker, one sheet named "Days_45" (for 45 days):
 *     ID | Name | Branch | Day 1 | Day 2 | ... | Day 45 | Attendance %
 *
 * MONTH-WISE tracker, one sheet per month ("Oct-2026", "Nov-2026", ...):
 *     ID | Name | Branch | 01 Thu | 02 Fri | ... | 31 Sat | Attendance %
 *     (one column per calendar date, in date order)
 *
 * Cells hold "P" (present) or "A" (absent); an unmarked day stays empty.
 * Attendance % = present / days marked so far, and is always the LAST column.
 */
public class AttendanceManager {

    private static final String STUDENTS_SHEET = "Students";
    private static final String PCT_HEADER = "Attendance %";
    private static final int FIXED_COLS = 3; // ID, Name, Branch

    private static final DateTimeFormatter MONTH_FMT = DateTimeFormatter.ofPattern("MMM-yyyy", Locale.ENGLISH);
    private static final DateTimeFormatter DAY_FMT = DateTimeFormatter.ofPattern("dd EEE", Locale.ENGLISH);

    private final String filePath;
    private Workbook wb;
    private final DataFormatter fmt = new DataFormatter();
    private CellStyle headerStyle;
    private CellStyle centerStyle;

    // ------------------------------------------------------------------ setup

    public AttendanceManager(String filePath) throws IOException {
        this.filePath = filePath;

        File f = new File(filePath);
        if (f.exists()) {
            try (FileInputStream in = new FileInputStream(f)) {
                wb = new XSSFWorkbook(in);
            }
        } else {
            wb = new XSSFWorkbook();
        }

        createStyles();

        if (wb.getSheet(STUDENTS_SHEET) == null) {
            Sheet sh = wb.createSheet(STUDENTS_SHEET);
            Row hr = sh.createRow(0);
            headerCell(hr, 0, "ID");
            headerCell(hr, 1, "Name");
            headerCell(hr, 2, "Branch");
        }
        save();
    }

    private void createStyles() {
        Font bold = wb.createFont();
        bold.setBold(true);

        headerStyle = wb.createCellStyle();
        headerStyle.setFont(bold);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        centerStyle = wb.createCellStyle();
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
    }

    private void headerCell(Row row, int col, String value) {
        Cell c = row.createCell(col);
        c.setCellValue(value);
        c.setCellStyle(headerStyle);
    }

    // ------------------------------------------------------- creating sheets

    public boolean sheetExists(String name) {
        return wb.getSheet(name) != null;
    }

    /** Day-wise tracker: one sheet with columns Day 1 ... Day N. Returns the sheet name. */
    public String ensureDaysSheet(int days) throws IOException {
        String name = "Days_" + days;
        if (wb.getSheet(name) == null) {
            String[] headers = new String[days];
            for (int i = 0; i < days; i++) {
                headers[i] = "Day " + (i + 1);
            }
            createAttendanceSheet(name, headers);
        }
        syncSheet(name);
        save();
        return name;
    }

    /** Month-wise tracker: one sheet per month, one column per calendar date. */
    public String ensureMonthSheet(YearMonth ym) throws IOException {
        String name = ym.format(MONTH_FMT);
        if (wb.getSheet(name) == null) {
            int len = ym.lengthOfMonth();
            String[] headers = new String[len];
            for (int d = 1; d <= len; d++) {
                headers[d - 1] = ym.atDay(d).format(DAY_FMT);
            }
            createAttendanceSheet(name, headers);
        }
        syncSheet(name);
        save();
        return name;
    }

    private void createAttendanceSheet(String name, String[] dayHeaders) {
        Sheet sh = wb.createSheet(name);
        Row hr = sh.createRow(0);

        headerCell(hr, 0, "ID");
        headerCell(hr, 1, "Name");
        headerCell(hr, 2, "Branch");
        for (int i = 0; i < dayHeaders.length; i++) {
            headerCell(hr, FIXED_COLS + i, dayHeaders[i]);
        }
        headerCell(hr, FIXED_COLS + dayHeaders.length, PCT_HEADER);

        sh.createFreezePane(FIXED_COLS, 1); // keep ID/Name/Branch visible while scrolling
    }

    // ----------------------------------------------------------- registration

    public boolean hasStudents() {
        return !getStudents().isEmpty();
    }

    public boolean idExists(String id) {
        for (Student s : getStudents()) {
            if (s.getId().equalsIgnoreCase(id)) return true;
        }
        return false;
    }

    public List<Student> getStudents() {
        List<Student> list = new ArrayList<>();
        Sheet sh = wb.getSheet(STUDENTS_SHEET);
        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row == null) continue;
            String id = text(row, 0);
            if (id.isEmpty()) continue;
            list.add(new Student(id, text(row, 1), text(row, 2)));
        }
        return list;
    }

    public void registerStudent(Student s) throws IOException {
        Sheet sh = wb.getSheet(STUDENTS_SHEET);
        writeIdentity(sh.createRow(sh.getLastRowNum() + 1), s);
        try {
            for (int c = 0; c < FIXED_COLS; c++) sh.autoSizeColumn(c);
        } catch (Throwable ignored) {
            // auto-sizing needs fonts; skip silently if the system has none
        }
        syncAllSheets();
        save();
    }

    public boolean deleteStudent(String id) throws IOException {
        boolean found = false;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            if (removeRowById(wb.getSheetAt(i), id)) found = true;
        }
        if (found) save();
        return found;
    }

    private boolean removeRowById(Sheet sh, String id) {
        int last = sh.getLastRowNum();
        for (int r = 1; r <= last; r++) {
            Row row = sh.getRow(r);
            if (row != null && text(row, 0).equalsIgnoreCase(id)) {
                if (r < last) {
                    sh.shiftRows(r + 1, last, -1);
                } else {
                    sh.removeRow(row);
                }
                return true;
            }
        }
        return false;
    }

    private void writeIdentity(Row row, Student s) {
        row.createCell(0).setCellValue(s.getId());
        row.createCell(1).setCellValue(s.getName());
        row.createCell(2).setCellValue(s.getBranch());
    }

    private void syncAllSheets() {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetAt(i).getSheetName();
            if (!name.equals(STUDENTS_SHEET)) syncSheet(name);
        }
    }

    /** Copies any student missing from an attendance sheet over from the master register. */
    private void syncSheet(String sheetName) {
        Sheet sh = wb.getSheet(sheetName);

        Set<String> existing = new HashSet<>();
        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row != null) existing.add(text(row, 0).toLowerCase());
        }

        for (Student s : getStudents()) {
            if (!existing.contains(s.getId().toLowerCase())) {
                writeIdentity(sh.createRow(sh.getLastRowNum() + 1), s);
            }
        }
        recalc(sh);
        layout(sh);
    }

    // ------------------------------------------------------------- attendance

    /** Number of day columns (Day 1..N, or dates of the month) in this sheet. */
    public int dayColumnCount(String sheetName) {
        return pctCol(wb.getSheet(sheetName)) - FIXED_COLS;
    }

    public String columnHeader(String sheetName, int dayIdx) {
        return text(wb.getSheet(sheetName).getRow(0), FIXED_COLS + dayIdx);
    }

    /** True if any student already has a mark in that day column. */
    public boolean isMarked(String sheetName, int dayIdx) {
        Sheet sh = wb.getSheet(sheetName);
        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row != null && !isBlank(row.getCell(FIXED_COLS + dayIdx))) return true;
        }
        return false;
    }

    /** First day column nobody has been marked in yet, or -1 if every day is done. */
    public int nextUnmarkedDay(String sheetName) {
        int n = dayColumnCount(sheetName);
        for (int i = 0; i < n; i++) {
            if (!isMarked(sheetName, i)) return i;
        }
        return -1;
    }

    /**
     * Writes one day's attendance.
     *
     * @param dayIdx 0-based day column (Day 1 = 0, or 1st of the month = 0)
     * @param marks  student id -> 1 (present) or 0 (absent)
     */
    public void recordAttendance(String sheetName, int dayIdx, Map<String, Integer> marks) throws IOException {
        syncSheet(sheetName);
        Sheet sh = wb.getSheet(sheetName);
        int col = FIXED_COLS + dayIdx;

        Map<String, Integer> byId = new HashMap<>();
        for (Map.Entry<String, Integer> e : marks.entrySet()) {
            byId.put(e.getKey().toLowerCase(), e.getValue());
        }

        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row == null) continue;

            Cell c = row.getCell(col);
            if (c == null) c = row.createCell(col);

            Integer v = byId.get(text(row, 0).toLowerCase());
            if (v == null) {
                c.setBlank();
            } else {
                c.setCellValue(v == 1 ? "P" : "A");
            }
            c.setCellStyle(centerStyle);
        }

        recalc(sh);
        save();
    }

    // ---------------------------------------------------------------- reports

    /**
     * Attendance of every student, combined over the given sheets
     * (one sheet for day-wise; one or all month sheets for month-wise).
     */
    public List<Student> getAllWithAttendance(List<String> sheetNames) {
        Map<String, Student> map = new LinkedHashMap<>();
        for (Student s : getStudents()) {
            map.put(s.getId().toLowerCase(), s);
        }

        for (String name : sheetNames) {
            Sheet sh = wb.getSheet(name);
            if (sh == null) continue;
            Row hr = sh.getRow(0);
            int pc = pctCol(sh);

            for (int r = 1; r <= sh.getLastRowNum(); r++) {
                Row row = sh.getRow(r);
                if (row == null) continue;
                Student s = map.get(text(row, 0).toLowerCase());
                if (s == null) continue;

                for (int c = FIXED_COLS; c < pc; c++) {
                    Cell cell = row.getCell(c);
                    if (isBlank(cell)) continue;

                    String key = sheetNames.size() > 1 ? name + " | " + text(hr, c) : text(hr, c);
                    int present = presentValue(cell);
                    s.addRecord(key, present == 1 ? "P" : "A", present, 1);
                }
            }
        }
        return new ArrayList<>(map.values());
    }

    public Student findStudent(List<String> sheetNames, String id) {
        for (Student s : getAllWithAttendance(sheetNames)) {
            if (s.getId().equalsIgnoreCase(id)) return s;
        }
        return null;
    }

    // ---------------------------------------------------------------- helpers

    /** Recomputes the "Attendance %" column: present / days marked so far. */
    private void recalc(Sheet sh) {
        int pc = pctCol(sh);

        for (int r = 1; r <= sh.getLastRowNum(); r++) {
            Row row = sh.getRow(r);
            if (row == null) continue;

            int present = 0;
            int total = 0;
            for (int c = FIXED_COLS; c < pc; c++) {
                Cell cell = row.getCell(c);
                if (isBlank(cell)) continue; // day not marked yet
                total++;
                present += presentValue(cell);
            }

            double pct = (total == 0) ? 0.0 : Math.round(present * 10000.0 / total) / 100.0;
            Cell p = row.getCell(pc);
            if (p == null) p = row.createCell(pc);
            p.setCellValue(pct);
            p.setCellStyle(centerStyle);
        }
    }

    private int pctCol(Sheet sh) {
        return sh.getRow(0).getLastCellNum() - 1;
    }

    private boolean isBlank(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) return true;
        return cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty();
    }

    private int presentValue(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            return "P".equalsIgnoreCase(cell.getStringCellValue().trim()) ? 1 : 0;
        }
        return 0;
    }

    private String text(Row row, int col) {
        if (row == null) return "";
        Cell c = row.getCell(col);
        return c == null ? "" : fmt.formatCellValue(c).trim();
    }

    /** ID/Name/Branch/% columns auto-fit; the many day columns get a fixed narrow width. */
    private void layout(Sheet sh) {
        try {
            int pc = pctCol(sh);
            for (int c = 0; c < FIXED_COLS; c++) sh.autoSizeColumn(c);
            for (int c = FIXED_COLS; c < pc; c++) sh.setColumnWidth(c, 2800);
            sh.autoSizeColumn(pc);
        } catch (Throwable ignored) {
            // auto-sizing needs fonts; skip silently if the system has none
        }
    }

    private void save() throws IOException {
        try (FileOutputStream out = new FileOutputStream(filePath)) {
            wb.write(out);
        }
    }
}