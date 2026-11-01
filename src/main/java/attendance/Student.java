package attendance;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Simple model for one student plus their attendance summary.
 */
public class Student {

    private final String id;
    private final String name;
    private final String branch;

    private int present = 0;
    private int total = 0;

    // column header (date/period or month) -> what to display (P / A / "18 / 24")
    private final Map<String, String> records = new LinkedHashMap<>();

    public Student(String id, String name, String branch) {
        this.id = id;
        this.name = name;
        this.branch = branch;
    }

    public void addRecord(String header, String display, int presentValue, int totalValue) {
        records.put(header, display);
        present += presentValue;
        total += totalValue;
    }

    public String getId()     { return id; }
    public String getName()   { return name; }
    public String getBranch() { return branch; }
    public int getPresent()   { return present; }
    public int getTotal()     { return total; }
    public Map<String, String> getRecords() { return records; }

    public double getPercentage() {
        if (total == 0) return 0.0;
        return Math.round(present * 10000.0 / total) / 100.0;
    }
}
