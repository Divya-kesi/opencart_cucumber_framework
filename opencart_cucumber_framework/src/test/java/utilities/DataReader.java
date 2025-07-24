package utilities;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataReader {
	// Static storage (not used in this method)
	public static HashMap<String, String> storeValues = new HashMap();

	public static List<HashMap<String, String>> data(String filepath, String sheetName) throws IOException 
	 {
		// 1. Initialize return structure
		List<HashMap<String, String>> mydata = new ArrayList<>();
		
		 // 2. Open Excel file
			FileInputStream file = new FileInputStream(filepath);
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet(sheetName);
			// 3. Get row info
			int totalRows=sheet.getLastRowNum(); // Last row index (0-based)
				
			XSSFRow headerRow=sheet.getRow(0); // Header row (index 0)
			
			// 4. Process data rows (start from index 1)
			for (int i = 1; i <= totalRows; i++) 
				{
				XSSFRow currentRow = sheet.getRow(i);
				
				HashMap<String, String> currentHash = new HashMap<String, String>();
				
				// 5. Process cells in current row
				for (int j = 0; j < currentRow.getLastCellNum(); j++) 
					{
					// Get header and data cells
					XSSFCell currentCell = currentRow.getCell(j); 
					// Store as key-value pair
					currentHash.put(headerRow.getCell(j).toString(), currentCell.toString());
				 }
				// 6. Add row data to result list
				mydata.add(currentHash);
				}
			// 7. Cleanup and return
			file.close();
			
		return mydata;
	}
}
