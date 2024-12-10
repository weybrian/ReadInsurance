package gainMiles;

import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadInsurance {

	public static void main(String[] args) {
		read("/Users/weybrian/Downloads/Q1.xlsx");
	}

	/**
	 * 讀取檔案
	 * @param string
	 */
	private static void read(String filepath) {
		try
        {
            FileInputStream file = new FileInputStream(filepath);
 
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook wb = new XSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            XSSFSheet ws = wb.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = ws.iterator();
            while (rowIterator.hasNext()) 
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                 
                while (cellIterator.hasNext()) 
                {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType()) 
                    {
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue());
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue());
                            break;
					default:
						break;
                    }
                }
                System.out.println("Reading File Completed.");
            }
            wb.close();
            file.close();
        } 
        catch (Exception ex) 
        {
            ex.printStackTrace();
        }
	}

}
