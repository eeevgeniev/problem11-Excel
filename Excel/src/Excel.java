import help.Office;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	
	// class in help package
	private static ArrayList<Office> offices = new ArrayList<Office>();
	
	public static void main(String[] args) {
		
		Locale.setDefault(Locale.ROOT);
		
		FileInputStream fileStream = null;
		
		// here will be saved columns names
		String[] colons = new String[6];
		
		// this checks if we read the first rows
		boolean isFirstRow = true;
		
		// counter for each cell in a specific row
		int counter = 0;
		
		double offcesTotal = 0D;
		
		try {
			
			fileStream = new FileInputStream(new File("3. Incomes-Report.xlsx"));
			
			XSSFWorkbook excelBook = new XSSFWorkbook(fileStream);
			
			XSSFSheet excelSheet = excelBook.getSheetAt(0);
			
			Iterator<Row> rowIterator = excelSheet.iterator();
			
			while (rowIterator.hasNext()) {
				
				String officeName = "";
				double income = 0D;
				double vat = 0D;
				
				Row currentRow = rowIterator.next();
				
				Iterator<Cell> cellIterator = currentRow.cellIterator();
				
				while (cellIterator.hasNext()) {
					
					Cell currentCell = cellIterator.next();
					
					// save the name of every column
					if (isFirstRow) {
						
						colons[counter] = currentCell.getStringCellValue();
						
					}
					else {
						
						// check 1 column
						if (counter == 0 && colons[0].equals("Office")) {
							
							officeName = currentCell.getStringCellValue();
							
						}
						// check 4 column
						else if (counter == 3 && colons[3].equals("Income")){
							
							income = currentCell.getNumericCellValue();
							
						}
						// check 5 column
						else if (counter == 4 && colons[4].equals("VAT")) {
							
							vat = retrunVAT(currentCell.getCellFormula());
						}
						
					}
					
					counter += 1;
					
				}
				
				// add new office if current row is different from first
				if (!isFirstRow) {
					
					addNewOffice(officeName, income, vat);
					
				}
				
				
				isFirstRow = false;
				
				counter = 0;
				
			}
			
			excelBook.close();
			
			fileStream.close();
			
		}
		catch (Exception ex) {
			
			try {
				
				if (fileStream != null) {
					
					fileStream.close();
					
				}
				
			}
			catch (Exception innerEx) {
				
				System.out.println(innerEx.getMessage());
			}
			
			System.out.println(ex.getMessage());
			
		}
		
		Collections.sort(offices);
		
		for (int i = 0; i < offices.size(); i++) {
			
			Office currentOffice = offices.get(i);
			
			double sumForOffice = currentOffice.returnTotalIncome();
			
			offcesTotal += sumForOffice;
			
			System.out.println(String.format(" %s total - > %.2f", 
					currentOffice.officeName, sumForOffice));
			
		}
		
		System.out.println(String.format(" %s total - > %.2f", 
				"Grand", offcesTotal));
		
	}
	
	// splits column vat which holds formulae
	private static double retrunVAT(String text) throws Exception {
		
		String[] values = text.split("\\*");
		
		double result = Double.parseDouble(values[1]);
		
		return result;
		
	}
	
	// checks if this office is added to offices
	private static void addNewOffice(String officeName, double income, double vat) {
		
		boolean isOfficeAdded = false;
		
		double totalIncome = income * vat + income;
		
		for (int i = 0; i < offices.size(); i++) {
			
			Office currentOffice = offices.get(i);
			
			if (currentOffice.officeName.equals(officeName)) {
				
				currentOffice.addIncome(totalIncome);
				isOfficeAdded = true;
				
				break;
				
			}
			
		}
		
		if (!isOfficeAdded) {
			
			Office newOffice = new Office();
			newOffice.officeName = officeName;
			newOffice.addIncome(totalIncome);
			
			offices.add(newOffice);
			
		}
		
	}

}

