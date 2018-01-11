import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelParser {
    private String filename;
    private Workbook workbook = null;
    private Iterator<Row> workbookItr = null;

    public ExcelParser(String file) {
        filename = file;
    }

    public boolean OpenWorkbook() {
        if (workbook != null) {
            System.out.println("Workbook is already loaded!");
            return false;
        }
        try {
            FileInputStream is = new FileInputStream(new File(filename));
            workbook = StreamingReader.builder()
                    .rowCacheSize(1000)
                    .bufferSize(4096)
                    .open(is);

            return true;
        } catch(FileNotFoundException ex) {
            ex.printStackTrace();
            return false;
        }
    }

    public List<List<String>> GetSheetData() {
        return GetSheetData(1000);
    }

    public List<List<String>> GetAllSheetData() {
        List<List<String>> sheetData = new ArrayList<List<String>>();
        Sheet sheet = workbook.getSheetAt((0));
        workbookItr = sheet.rowIterator();

        while(workbookItr.hasNext()) {
            List<String> rowData = new ArrayList<String>();
            for (Cell c : workbookItr.next())
                rowData.add(c.getStringCellValue());

            sheetData.add(rowData);

        }
        return sheetData;
    }

    public List<List<String>> GetSheetData(int limit) {
        List<List<String>> sheetData = new ArrayList<List<String>>();
        if (workbookItr == null) {
            Sheet sheet = workbook.getSheetAt((0));
            workbookItr = sheet.rowIterator();
        }

        for (int index = 0; workbookItr.hasNext(); index++) {
            if (index > limit)
                return sheetData;

            List<String> rowData = new ArrayList<String>();
            for (Cell c : workbookItr.next())
                rowData.add(c.getStringCellValue());

            sheetData.add(rowData);

        }

        return sheetData;
    }

    public static void main(String[] args) {
        String filename = args[0];
        final long startTime = System.currentTimeMillis();

        ExcelParser parser = new ExcelParser(filename);

        if (!parser.OpenWorkbook()) {
            System.out.println("Unable to open workbook!");
            return;
        }

        List<List<String>> data = parser.GetAllSheetData();
        final long endTime = System.currentTimeMillis();
        System.out.println("Parsed: " + data.size() + " rows in: " + (endTime - startTime) + "ms.");
        System.out.println(data.get(1));
    }
}
