import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by R00715649 on 22-Oct-16.
 * This class reads all info from packing list.
 */
public class PackingList {
    XSSFWorkbook Workbook;
    List<Cell>UsedRange;

    //Constructor
    PackingList(String filepath) throws FileNotFoundException, IOException {
        File file = new File(filepath);
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook = new XSSFWorkbook(fileInputStream);
        UsedRange = getUsedRange();
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

    public static void main(String[] args) throws IOException {
        PackingList pl = new PackingList("D:\\myScripts\\InvoiceMaker2\\sample1\\0Y02181600000SHWA05K with price.xlsx");
        System.out.println(pl.getSheetNames());
        System.out.println("part number " + pl.getPartNumberColumn());
        System.out.println("model " + pl.getModelColumn());
        System.out.println("description " + pl.getDescriptionColumn());
        System.out.println("quantity " + pl.getQuantityColumn());
        System.out.print("uom " + pl.getUOMColumn());
        System.out.println("unit price " + pl.getUnitPriceColumn());

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