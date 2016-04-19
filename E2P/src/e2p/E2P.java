
package e2p;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import e2p.fileInput.excelFile;
import static e2p.fileInput.excelInputFile;
import static e2p.fileInput.excelOutputFile;
import java.io.*;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
public class E2P
{
    public static FileInputStream file;
    public static Workbook workbook;
    public static Document doc;
    public static void writeToPdf(String s ) {
        

        try {
            doc.add(new Paragraph(s));
        } catch (DocumentException e) {
            e.printStackTrace();
        }

    }
    public static String cellToString(XSSFCell cell) {  
    int type;
    Object result=null;
    type = cell.getCellType();

    switch (type) {

        case Cell.CELL_TYPE_NUMERIC: // numeric value in Excel
        case Cell.CELL_TYPE_FORMULA: // precomputed value based on formula
            result = cell.getNumericCellValue();
            break;
        case Cell.CELL_TYPE_STRING: // String Value in Excel 
            result = cell.getStringCellValue();
            break;
        
        case Cell.CELL_TYPE_BLANK:
            result = "";
        case Cell.CELL_TYPE_BOOLEAN: //boolean value 
            result: cell.getBooleanCellValue();
            break;
        case Cell.CELL_TYPE_ERROR:
        default:  
            throw new RuntimeException("There is no support for this type of cell");                        
    }

    return result.toString();
}
    public void converter() throws DocumentException, FileNotFoundException, IOException
    {
        
            file = new FileInputStream(excelInputFile);
 
            //Create Workbook instance holding reference to .xlsx file
            workbook = new XSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);
            
            doc=new Document();
            doc.open();
            
            PdfWriter.getInstance(doc, new FileOutputStream(excelOutputFile));
            
            doc.open();
            
            Paragraph paragraph=new Paragraph();
            
            Iterator<Row> rowIterator;
            rowIterator = sheet.iterator();
            Row row = rowIterator.next();
            row = rowIterator.next();
            row = rowIterator.next();
            while(rowIterator.hasNext()){
                row = rowIterator.next();
                
                String data;
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell = cellIterator.next();
                String arr[]=new String[16];
                
                
                data=cellToString((XSSFCell) cell);
                int i=0;
                while(cellIterator.hasNext()) {
                    cell = cellIterator.next();
                    arr[i]=cellToString((XSSFCell) cell);
                    i++;
                }
                
                data="                                                                                                        " +arr[6]; //village
                writeToPdf(data);
                
                data="\n\n                  "+arr[0]+"                             "+arr[2]+"/"+arr[3]+"                             "+ arr[13] ; //acc no,billcy,billgrp,issuedate
                writeToPdf(data);
                
                data="\n                  "+arr[14]+"                                    "+"cash collection center"; //due date,cash collection center
                writeToPdf(data);
                data="\n\n\n                                                                                       "+arr[4];//name
                writeToPdf(data);
                data="                                                                                       s/o"+arr[5];//fathers name
                writeToPdf(data);
                data="                  bill no.                                                      "+arr[6];
                writeToPdf(data);
                doc.newPage();
            }
            doc.close();
            file.close();
        }
}
            