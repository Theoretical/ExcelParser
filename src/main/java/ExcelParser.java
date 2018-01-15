import com.monitorjbl.xlsx.StreamingReader;
import com.sun.org.apache.xpath.internal.SourceTree;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.jopendocument.dom.spreadsheet.SpreadSheet;
import com.smartxls.WorkBook;


import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelParser {
    private String filename;
    private Workbook workbook = null;
    private Iterator<Row> workbookItr = null;
    private SpreadSheet odsSheet = null;
    private int odsIndex = 0;

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
            if (filename.endsWith(".xlsx") || filename.endsWith(".xlsm")) {
                workbook = StreamingReader.builder()
                        .rowCacheSize(1000)
                        .bufferSize(4096)
                        .open(is);
            } else if(filename.endsWith(".ods")) {
                odsSheet = SpreadSheet.createFromFile(new File(filename));
            } else {
                workbook = new HSSFWorkbook(is);
            }

            return true;
        } catch(FileNotFoundException ex) {
            ex.printStackTrace();
            return false;
        }
        catch (IOException ex) {
            ex.printStackTrace();
            return false;
        }
    }

    public List<List<String>> GetAllSheetData() {
        List<List<String>> sheetData = new ArrayList<List<String>>();

        if (odsSheet != null) {
            for(int sheetIndex = 0; sheetIndex < odsSheet.getSheetCount(); sheetIndex++) {
                org.jopendocument.dom.spreadsheet.Sheet sheet = odsSheet.getSheet(sheetIndex);

                for(int row = 0; row < sheet.getRowCount(); row++) {
                    List<String> rowData = new ArrayList<String>();
                    for (int col = 0; col < sheet.getColumnCount(); col++) {

                        if (sheet.getCellAt(col, row).getValueType() == null) {
                            break;
                        }
                        rowData.add(sheet.getCellAt(col, row).getTextValue());
                    }

                    sheetData.add(rowData);
                }

            }
            return sheetData;
        }

        for (int i = 0; i < workbook.getNumberOfSheets(); i++){
            Sheet sheet = workbook.getSheetAt(i);
            workbookItr = sheet.rowIterator();

            while (workbookItr.hasNext()) {
                List<String> rowData = new ArrayList<String>();
                for (Cell c : workbookItr.next())
                    rowData.add(c.getStringCellValue());

                sheetData.add(rowData);

            }
        }
        return sheetData;
    }

    public List<List<String>> GetSheetData() {
        return GetSheetData(0, 1000);
    }

    public List<List<String>> GetSheetData(int sheetIndex, int limit) {
        List<List<String>> sheetData = new ArrayList<List<String>>();

        if (odsSheet != null) {
            org.jopendocument.dom.spreadsheet.Sheet sheet = odsSheet.getSheet(sheetIndex);
            int rowCount = sheet.getRowCount();
            System.out.println(rowCount);
            System.out.println(sheet.getHeaderColumnCount());
            int row = odsIndex;

            if (row >= rowCount)
                return sheetData;

            for(; row < sheet.getRowCount(); row++) {
                List<String> rowData = new ArrayList<String>();
                for (int col = 0; col < sheet.getColumnCount(); col++) {

                    if (sheet.getCellAt(col, row).getValueType() == null) {
                        break;
                    }
                    rowData.add(sheet.getCellAt(col, row).getTextValue());
                }

                sheetData.add(rowData);
                if (row >= limit) break;
            }

            return sheetData;
        }
        if (workbookItr == null) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            workbookItr = sheet.rowIterator();
        }

        for (int index = 0; workbookItr.hasNext(); index++) {
            if (index > limit)
                return sheetData;

            List<String> rowData = new ArrayList<String>();
            for (Cell c : workbookItr.next())
                if (c.getCellTypeEnum() == CellType.STRING)
                    rowData.add(c.getStringCellValue());
                else if (c.getCellTypeEnum() == CellType.NUMERIC)
                    rowData.add(String.valueOf(c.getNumericCellValue()));

            sheetData.add(rowData);

        }

        return sheetData;
    }

    public static void main(String[] args) {
        String filename = "C:\\1test.ods";//args[0];

        if(filename.endsWith(".xlsb")) {

            try {
                WorkBook wb = new WorkBook();
                wb.readXLSB(filename);
                wb.writeXLSX("tmp.xlsx");
                filename = "tmp.xlsx";
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        final long startTime = System.currentTimeMillis();

        System.out.println("Parsing: " + filename);
        ExcelParser parser = new ExcelParser(filename);

        if (!parser.OpenWorkbook()) {
            System.out.println("Unable to open workbook!");
            return;
        }

        List<List<String>> data = parser.GetAllSheetData();
        final long endTime = System.currentTimeMillis();
        System.out.println("Parsed: " + data.size() + " rows in: " + (endTime - startTime) + "ms.");

        for (int i = 0; i < data.size(); i++)
            System.out.println(data.get(i));


    }
}
