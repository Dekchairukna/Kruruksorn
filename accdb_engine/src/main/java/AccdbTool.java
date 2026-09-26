import com.healthmarketscience.jackcess.*;
import com.healthmarketscience.jackcess.crypt.CryptCodecProvider;
import org.json.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;

public class AccdbTool {
    static final String[] NUMF = {"UM01","UM02","MidtermMark","FinalMark","UnitMark","TotalMark","TotalPercent",
        "QM1","QM2","QM3","QM4","QM5","QM6","QM7","QM8","QualityMark","LM1","LM2","LM3","LM4","LM5","LiteratureMark"};
    static final String[] GRADEF = {"Grade","QGrade","LGrade"};

    public static void main(String[] args) throws Exception {
        String cmd = args[0], path = args[1], pw = args[2];
        Database db = new DatabaseBuilder(new File(path))
                .setCodecProvider(new CryptCodecProvider(pw))
                .setReadOnly(cmd.equals("read"))
                .open();
        try {
            if (cmd.equals("read")) doRead(db);
            else doWrite(db, args[3]);
        } finally { db.close(); }
    }

    static void doRead(Database db) throws Exception {
        JSONObject out = new JSONObject();
        JSONArray studs = new JSONArray();
        Table st = db.getTable("STUDENTS");
        for (Row r : st) {
            JSONObject o = new JSONObject();
            o.put("id", str(r.get("ID"))); o.put("prefix", str(r.get("PREFIX")));
            o.put("fn", str(r.get("FIRSTNAME"))); o.put("ln", str(r.get("LASTNAME")));
            o.put("room", str(r.get("ROOM"))); o.put("no", str(r.get("ORDINAL")));
            studs.put(o);
        }
        out.put("students", studs);
        out.put("ok", true);
        System.out.println(out.toString());
    }

    static void doWrite(Database db, String scoresPath) throws Exception {
        String js = new String(Files.readAllBytes(Paths.get(scoresPath)), "UTF-8");
        JSONObject payload = new JSONObject(js);
        JSONArray rows = payload.getJSONArray("rows");
        Map<String, JSONObject> byId = new HashMap<>();
        for (int i = 0; i < rows.length(); i++) {
            JSONObject r = rows.getJSONObject(i);
            byId.put(String.valueOf(r.get("sid")), r);
        }
        // TRANSCRIPTS
        Table tr = db.getTable("TRANSCRIPTS");
        Set<String> trCols = new HashSet<>();
        for (Column c : tr.getColumns()) trCols.add(c.getName());
        java.util.List<Row> trRows = new java.util.ArrayList<>();
        for (Row row : tr) trRows.add(row);
        for (Row row : trRows) {
            String id = str(row.get("ID"));
            JSONObject r = byId.get(id);
            if (r == null) continue;
            for (String f : NUMF) if (trCols.contains(f) && r.has(f)) row.put(f, r.optDouble(f, 0));
            for (String f : GRADEF) if (trCols.contains(f)) {
                String key = f.equals("Grade") ? "grade" : (f.equals("QGrade") ? "qgrade" : "lgrade");
                String v = r.optString(key, "").trim();
                row.put(f, v.isEmpty() ? null : v);
            }
            tr.updateRow(row);
        }
        // TRANSCRIPTS2 (sub-scores um<nn>_<k>)
        try {
            Table t2 = db.getTable("TRANSCRIPTS2");
            Set<String> t2cols = new HashSet<>();
            for (Column c : t2.getColumns()) t2cols.add(c.getName());
            java.util.List<Row> t2Rows = new java.util.ArrayList<>();
            for (Row row : t2) t2Rows.add(row);
            for (Row row : t2Rows) {
                String id = str(row.get("ID"));
                JSONObject r = byId.get(id);
                if (r == null || !r.has("sub")) continue;
                JSONObject sub = r.getJSONObject("sub");
                boolean changed = false;
                for (String u : sub.keySet()) {
                    JSONArray arr = sub.getJSONArray(u);
                    String nn = String.format("%02d", Integer.parseInt(u));
                    for (int k = 0; k < arr.length() && k < 5; k++) {
                        String col = "um" + nn + "_" + (k + 1);
                        if (t2cols.contains(col)) { row.put(col, arr.optDouble(k, 0)); changed = true; }
                    }
                }
                if (changed) t2.updateRow(row);
            }
        } catch (Exception ignore) {}
        db.flush();
        System.out.println("{\"ok\":true,\"updated\":" + rows.length() + "}");
    }

    static String str(Object o) { return o == null ? "" : String.valueOf(o); }
}
