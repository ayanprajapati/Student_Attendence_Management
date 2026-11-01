package attendance;

import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

public class Main {

    private static final String FILE_NAME = "attendance.xlsx";
    private static final Scanner in = new Scanner(System.in);

    private static AttendanceManager mgr;

    // Current tracker
    private static boolean daysMode;                          // true = day-wise, false = month-wise
    private static final List<String> sheets = new ArrayList<>();      // sheet name(s) of the tracker
    private static final List<YearMonth> months = new ArrayList<>();   // month-wise only (same order as sheets)

    public static void main(String[] args) {
        System.setProperty("java.awt.headless", "true");

        System.out.println("=====================================");
        System.out.println("   STUDENT ATTENDANCE MANAGEMENT");
        System.out.println("=====================================");

        try {
            mgr = new AttendanceManager(FILE_NAME);
            chooseTracker();
            startupFlow();
            menuLoop();
        } catch (IOException e) {
            System.out.println("\nCould not read/write '" + FILE_NAME + "': " + e.getMessage());
            System.out.println("If the file is open in Excel, close it and run the program again.");
        }
    }

    // ------------------------------------------------------------ main flow

    private static void chooseTracker() throws IOException {
        System.out.println("\nHow do you want to record attendance?");
        System.out.println("  1. Day-wise ");
        System.out.println("  2. Month-wise ");
        int choice = askInt("Choose (1-2): ", 1, 2);

        sheets.clear();
        months.clear();

        if (choice == 1) {
            daysMode = true;
            int days = askInt("For how many days? (1-365): ", 1, 365);
            String name = "Days_" + days;
            boolean existed = mgr.sheetExists(name);
            mgr.ensureDaysSheet(days);
            sheets.add(name);
            System.out.println(existed
                    ? "Found existing sheet '" + name + "' - continuing from it."
                    : "Created sheet '" + name + "' with columns Day 1 to Day " + days + ".");
        } else {
            daysMode = false;
            int count = askInt("For how many months? (1-12): ", 1, 12);
            YearMonth start = askMonth("Starting month (yyyy-MM, Enter = current month): ");
            for (int i = 0; i < count; i++) {
                YearMonth ym = start.plusMonths(i);
                months.add(ym);
                sheets.add(mgr.ensureMonthSheet(ym));
            }
            System.out.println("Sheets ready: " + String.join(", ", sheets));
        }
    }

    private static void startupFlow() throws IOException {
        if (!mgr.hasStudents()) {
            System.out.println("\nNo students registered yet. Let's register them first.");
            do {
                registerOne();
            } while (askYes("Register another student?"));
        } else {
            System.out.println("\n" + mgr.getStudents().size() + " registered student(s) found in " + FILE_NAME + ".");
        }

        if (mgr.hasStudents() && askYes("Mark attendance now?")) {
            markAttendance();
        }
    }

    private static void menuLoop() {
        while (true) {
            String label = daysMode ? "DAY-WISE: " + sheets.get(0) : "MONTH-WISE: " + String.join(", ", sheets);
            System.out.println("\n--------- MENU (" + label + ") ---------");
            System.out.println(" 1. Mark attendance");
            System.out.println(" 2. Register new student");
            System.out.println(" 3. Student info (by ID)");
            System.out.println(" 4. Students below a given attendance %");
            System.out.println(" 5. View all students");
            System.out.println(" 6. Delete a student");
            System.out.println(" 7. Switch tracker (day-wise / month-wise)");
            System.out.println(" 0. Exit");

            int choice = askInt("Choose: ", 0, 7);
            try {
                switch (choice) {
                    case 1: markAttendance(); break;
                    case 2: registerOne(); break;
                    case 3: showStudentInfo(); break;
                    case 4: showBelowThreshold(); break;
                    case 5: viewAll(); break;
                    case 6: deleteStudent(); break;
                    case 7: chooseTracker(); break;
                    default:
                        System.out.println("Goodbye! Data saved in " + FILE_NAME);
                        return;
                }
            } catch (IOException e) {
                System.out.println("File error: " + e.getMessage());
                System.out.println("If the file is open in Excel, close it and try again.");
            }
        }
    }

    // --------------------------------------------------------- registration

    private static void registerOne() throws IOException {
        String id;
        while (true) {
            id = ask("Student ID     : ");
            if (id.isEmpty()) {
                System.out.println("ID cannot be empty.");
            } else if (mgr.idExists(id)) {
                System.out.println("That ID is already registered. Use a different one.");
            } else {
                break;
            }
        }
        String name = askNonEmpty("Student name   : ");
        String branch = askNonEmpty("Branch         : ");

        mgr.registerStudent(new Student(id, name, branch));
        System.out.println("Registered: " + id + " - " + name + " (" + branch + ")");
    }

    private static void deleteStudent() throws IOException {
        String id = ask("ID of student to delete: ");
        if (!mgr.idExists(id)) {
            System.out.println("No student with that ID.");
            return;
        }
        if (askYes("This removes the student and ALL their attendance from every sheet. Sure?")) {
            mgr.deleteStudent(id);
            System.out.println("Deleted.");
        }
    }

    // ----------------------------------------------------------- attendance

    private static void markAttendance() throws IOException {
        List<Student> students = mgr.getStudents();
        if (students.isEmpty()) {
            System.out.println("Register students first.");
            return;
        }

        String sheet;
        int dayIdx;

        if (daysMode) {
            // ---- day-wise: pick Day 1, Day 2, ... ----
            sheet = sheets.get(0);
            int total = mgr.dayColumnCount(sheet);
            int next = mgr.nextUnmarkedDay(sheet);
            if (next == -1) {
                System.out.println("All " + total + " days are already recorded. You can still re-mark a day to correct it.");
            }
            dayIdx = askDayNumber(total, next);
        } else {
            // ---- month-wise: pick a calendar date, it decides the sheet and the column ----
            LocalDate date = askDate();
            int monthIdx = months.indexOf(YearMonth.from(date));
            if (monthIdx < 0) {
                System.out.println("There is no sheet for " + YearMonth.from(date)
                        + ". Your sheets are: " + String.join(", ", sheets));
                return;
            }
            if (date.isAfter(LocalDate.now()) && !askYes(date + " is in the future. Mark anyway?")) {
                return;
            }
            if (date.getDayOfWeek() == DayOfWeek.SUNDAY && !askYes(date + " is a Sunday. Mark anyway?")) {
                return;
            }
            sheet = sheets.get(monthIdx);
            dayIdx = date.getDayOfMonth() - 1;
        }

        String header = mgr.columnHeader(sheet, dayIdx);
        System.out.println("\nMarking attendance for [" + sheet + "] " + header);

        if (mgr.isMarked(sheet, dayIdx) && !askYes("This day is already recorded. Overwrite it?")) {
            return;
        }

        Map<String, Integer> marks = collectMarks(students);

        int absent = 0;
        for (int v : marks.values()) if (v == 0) absent++;
        System.out.println("Present: " + (marks.size() - absent) + "   Absent: " + absent);

        mgr.recordAttendance(sheet, dayIdx, marks);
        System.out.println("Saved to " + FILE_NAME);
    }

    /** Asks P/A for every registered student. Returns id -> 1 (present) / 0 (absent). */
    private static Map<String, Integer> collectMarks(List<Student> students) {
        Map<String, Integer> marks = new LinkedHashMap<>();

        if (askYes("Mark everyone present and only enter the absentees?")) {
            for (Student s : students) marks.put(s.getId(), 1);

            String line = ask("Absent IDs (comma separated, Enter if none): ");
            if (!line.isEmpty()) {
                for (String token : line.split(",")) {
                    String id = token.trim();
                    boolean matched = false;
                    for (Student s : students) {
                        if (s.getId().equalsIgnoreCase(id)) {
                            marks.put(s.getId(), 0);
                            matched = true;
                        }
                    }
                    if (!matched && !id.isEmpty()) {
                        System.out.println("  (ignored unknown ID: " + id + ")");
                    }
                }
            }
        } else {
            for (Student s : students) {
                String prompt = String.format("%-10s %-20s (P/A, Enter = P): ", s.getId(), s.getName());
                String ans;
                while (true) {
                    ans = ask(prompt).toUpperCase();
                    if (ans.isEmpty() || ans.equals("P") || ans.equals("A")) break;
                    System.out.println("  Please type P or A.");
                }
                marks.put(s.getId(), ans.equals("A") ? 0 : 1);
            }
        }
        return marks;
    }

    // -------------------------------------------------------------- reports

    private static void showStudentInfo() {
        String id = ask("Enter student ID: ");
        if (!mgr.idExists(id)) {
            System.out.println("No student found with ID '" + id + "'.");
            return;
        }

        Student overall = mgr.findStudent(sheets, id);

        System.out.println("\n----- Student Details -----");
        System.out.println("ID     : " + overall.getId());
        System.out.println("Name   : " + overall.getName());
        System.out.println("Branch : " + overall.getBranch());

        if (overall.getRecords().isEmpty()) {
            System.out.println("\nNo attendance recorded yet.");
            return;
        }

        // one block per sheet (a single block for day-wise)
        for (String sheet : sheets) {
            Student s = mgr.findStudent(List.of(sheet), id);
            System.out.println("\n[" + sheet + "]");
            if (s == null || s.getRecords().isEmpty()) {
                System.out.println("  nothing recorded yet");
                continue;
            }

            int i = 0;
            for (Map.Entry<String, String> e : s.getRecords().entrySet()) {
                System.out.printf("  %-12s", e.getKey() + " = " + e.getValue());
                if (++i % 5 == 0) System.out.println();
            }
            if (i % 5 != 0) System.out.println();
            System.out.printf("  -> %d of %d days present (%.2f%%)%n", s.getPresent(), s.getTotal(), s.getPercentage());
        }

        System.out.printf("%nOVERALL: %d of %d days present = %.2f%%  %s%n",
                overall.getPresent(), overall.getTotal(), overall.getPercentage(),
                overall.getPercentage() < 75 ? "(LOW - below 75%)" : "");
    }

    private static void showBelowThreshold() {
        List<String> scope = chooseScope();
        double limit = askDouble("Show students with attendance below (%): ", 0, 100);

        List<Student> low = new ArrayList<>();
        for (Student s : mgr.getAllWithAttendance(scope)) {
            if (s.getPercentage() < limit) low.add(s);
        }

        if (low.isEmpty()) {
            System.out.println("No student is below " + limit + "%.");
        } else {
            System.out.println("\nStudents below " + limit + "% (" + low.size() + ") - " + String.join(", ", scope));
            printTable(low);
        }
    }

    private static void viewAll() {
        List<String> scope = chooseScope();
        List<Student> all = mgr.getAllWithAttendance(scope);
        if (all.isEmpty()) {
            System.out.println("No students registered.");
        } else {
            System.out.println("\nAttendance - " + String.join(", ", scope));
            printTable(all);
        }
    }

    /** Month-wise only: all months combined, or one particular month. */
    private static List<String> chooseScope() {
        if (daysMode || sheets.size() == 1) return sheets;

        System.out.println("Which period?");
        System.out.println("  0. All months combined");
        for (int i = 0; i < sheets.size(); i++) {
            System.out.println("  " + (i + 1) + ". " + sheets.get(i));
        }
        int c = askInt("Choose: ", 0, sheets.size());
        return c == 0 ? sheets : List.of(sheets.get(c - 1));
    }

    private static void printTable(List<Student> list) {
        System.out.println();
        System.out.printf("%-10s %-20s %-10s %9s %9s%n", "ID", "Name", "Branch", "Present", "Attend %");
        System.out.println("-".repeat(62));
        for (Student s : list) {
            System.out.printf("%-10s %-20s %-10s %4d/%-4d %8.2f%%%n",
                    s.getId(), s.getName(), s.getBranch(),
                    s.getPresent(), s.getTotal(), s.getPercentage());
        }
    }

    // ------------------------------------------------------- input helpers

    private static String ask(String prompt) {
        System.out.print(prompt);
        return in.nextLine().trim();
    }

    private static String askNonEmpty(String prompt) {
        while (true) {
            String s = ask(prompt);
            if (!s.isEmpty()) return s;
            System.out.println("This field cannot be empty.");
        }
    }

    private static int askInt(String prompt, int min, int max) {
        while (true) {
            try {
                int v = Integer.parseInt(ask(prompt));
                if (v >= min && v <= max) return v;
            } catch (NumberFormatException ignored) {
            }
            System.out.println("Please enter a whole number between " + min + " and " + max + ".");
        }
    }

    /** Asks for Day number; returns 0-based index. Enter picks the next unmarked day if there is one. */
    private static int askDayNumber(int total, int next) {
        while (true) {
            String prompt = next >= 0
                    ? "Day number to mark (1-" + total + ", Enter = Day " + (next + 1) + "): "
                    : "Day number to mark (1-" + total + "): ";
            String s = ask(prompt);
            if (s.isEmpty() && next >= 0) return next;
            try {
                int v = Integer.parseInt(s);
                if (v >= 1 && v <= total) return v - 1;
            } catch (NumberFormatException ignored) {
            }
            System.out.println("Please enter a number between 1 and " + total + ".");
        }
    }

    private static double askDouble(String prompt, double min, double max) {
        while (true) {
            try {
                double v = Double.parseDouble(ask(prompt));
                if (v >= min && v <= max) return v;
            } catch (NumberFormatException ignored) {
            }
            System.out.println("Please enter a number between " + min + " and " + max + ".");
        }
    }

    private static boolean askYes(String prompt) {
        while (true) {
            String s = ask(prompt + " (y/n): ").toLowerCase();
            if (s.equals("y") || s.equals("yes")) return true;
            if (s.equals("n") || s.equals("no")) return false;
        }
    }

    private static LocalDate askDate() {
        while (true) {
            String s = ask("Date to mark (yyyy-MM-dd, Enter = today): ");
            if (s.isEmpty()) return LocalDate.now();
            try {
                return LocalDate.parse(s);
            } catch (DateTimeParseException e) {
                System.out.println("Invalid date. Example: 2026-10-31");
            }
        }
    }

    private static YearMonth askMonth(String prompt) {
        while (true) {
            String s = ask(prompt);
            if (s.isEmpty()) return YearMonth.now();
            try {
                return YearMonth.parse(s);
            } catch (DateTimeParseException e) {
                System.out.println("Invalid month. Example: 2026-10");
            }
        }
    }
}