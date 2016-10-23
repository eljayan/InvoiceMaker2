import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by R00715649 on 22-Oct-16.
 * This class reads all info from packing list.
 */
public class PackingList {
    XSSFWorkbook Workbook;
    List<Cell>UsedRange;
    Map<String, String>CaseRows;

    //Constructor
    PackingList(String filepath) throws FileNotFoundException, IOException {
        File file = new File(filepath);
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook = new XSSFWorkbook(fileInputStream);
        UsedRange = getUsedRange();
        CaseRows = getCaseRows();
    }

    List<String> getSheetNames() {
        List<String> sheetNames = new ArrayList<String>();
        Iterator<Sheet> sheetIterator = Workbook.iterator();
        while (sheetIterator.hasNext()) {
            sheetNames.add(sheetIterator.next().getSheetName());
        }
        return sheetNames;
    }

    List<Cell>getUsedRange(){
        List<Cell>cellList = new ArrayList<Cell>();
        Iterator<Row> rowIterator = Workbook.getSheet("Sheet1").rowIterator();
        while (rowIterator.hasNext()){
            Row row = rowIterator.next();
            Iterator<Cell>cellIterator = row.cellIterator();
            while (cellIterator.hasNext()){
                cellList.add(cellIterator.next());

            }
        }
        return cellList;
    }

    Integer findColumnIndex(Pattern pattern){
        //searches for a pattern in all cells and returns the
        //cell column index

        Iterator<Cell> cellIterator = UsedRange.iterator();
        while (cellIterator.hasNext()){
            Cell cell = cellIterator.next();
            Matcher matcher = pattern.matcher(cell.getStringCellValue());
            if (matcher.find()){
                return cell.getColumnIndex();
            }
        }
        return null;
    }
    Integer getCaseNumberColumn(){
        Pattern pattern = Pattern.compile("case\\.no\\.", Pattern.CASE_INSENSITIVE);
        return findColumnIndex(pattern);

    }
    Integer getPartNumberColumn(){
        Pattern pattern = Pattern.compile("part\\s+number", Pattern.CASE_INSENSITIVE);
        return findColumnIndex(pattern);
    }

    Integer getModelColumn(){
        Pattern pattern = Pattern.compile("model", Pattern.CASE_INSENSITIVE);
        return findColumnIndex(pattern);
    }

    Integer getDescriptionColumn(){
        Pattern pattern = Pattern.compile("description", Pattern.CASE_INSENSITIVE);
        return findColumnIndex(pattern);
    }

    Integer getQuantityColumn(){
        Pattern pattern = Pattern.compile("qty", Pattern.CASE_INSENSITIVE);
        return findColumnIndex(pattern);
    }

    Integer getUOMColumn(){
        Pattern pattern = Pattern.compile("uom", Pattern.CASE_INSENSITIVE);
        return findColumnIndex(pattern);
    }

    Integer getUnitPriceColumn(){
        Pattern pattern = Pattern.compile("unit\\sprice", Pattern.CASE_INSENSITIVE);
        return findColumnIndex(pattern);
    }

    Integer getTotalPriceColumn(){
        Pattern pattern = Pattern.compile("total\\sprice", Pattern.CASE_INSENSITIVE);
        return findColumnIndex(pattern);
    }

    Integer getNoteColumn(){
        Pattern pattern = Pattern.compile("note|remark", Pattern.CASE_INSENSITIVE);
        return findColumnIndex(pattern);
    }


    boolean isMergedCell(Cell cell){
        return true;
    }

    Map<String, String> getCaseRows(){
//    void getCaseRows(){
        //returns a dictionary with the row number and the case number it belongs to
        XSSFSheet sheet = Workbook.getSheet("Sheet1");
        List<CellRangeAddress> mergedcells = sheet.getMergedRegions();
        Iterator<CellRangeAddress> mergedCellsIterator = mergedcells.iterator();
        while (mergedCellsIterator.hasNext()){
            CellRangeAddress cellRangeAddress = mergedCellsIterator.next();
            int firstColumn = cellRangeAddress.getFirstColumn();
            int firstRow = cellRangeAddress.getFirstRow();
            int lastRow = cellRangeAddress.getLastRow();

            if (firstColumn == 0 && firstRow >3 ){

            }else{
                continue;
            }
        }
    }

    public static void main(String[] args) throws IOException {
        PackingList pl = new PackingList("D:\\myScripts\\InvoiceMaker2\\sample1\\0Y02181600000SHWA05K with price.xlsx");
        System.out.println(pl.getSheetNames());
        System.out.println("part number " + pl.getPartNumberColumn());
        System.out.println("model " + pl.getModelColumn());
        System.out.println("description " + pl.getDescriptionColumn());
        System.out.println("quantity " + pl.getQuantityColumn());
        System.out.print("uom " + pl.getUOMColumn());
        System.out.println("unit price " + pl.getUnitPriceColumn());
        pl.getCaseRows();
        System.out.println("this is a test");
    }

}

//class test{
//    public static void main(String[] args) throws IOException {
//        PackingList pl = new PackingList("D:\\myScripts\\InvoiceMaker2\\sample1\\0Y02181600000SHWA05K with price.xlsx");
//        System.out.println(pl.getSheetNames());
//        System.out.println("this is a test");
//    }
//}