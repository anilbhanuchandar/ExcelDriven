import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xlsx {

	public static void main(String[] args) throws IOException {
		
		FileInputStream file = new FileInputStream("C:\\Users\\Anil kumar V\\OneDrive\\Desktop\\Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		int sheets = workbook.getNumberOfSheets();
		for(int i=0;i<sheets;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> row = sheet.iterator();
				Row firstrow = row.next();
				Iterator<Cell> cell = firstrow.cellIterator();
				int k=0;
				int column=0;
				while(cell.hasNext())
				{
					Cell ce = cell.next();
					if(ce.getStringCellValue().equalsIgnoreCase("testcases"))
							{
						            column=k;
							}
					k++;
				}
				System.out.println(column);
				
				while(row.hasNext())
				{
					Row cl = row.next();
					if(cl.getCell(column).getStringCellValue().equalsIgnoreCase("purchase"))
					{
						Iterator<Cell> ceee = cl.cellIterator();
						while(ceee.hasNext())
						{
						System.out.println(ceee.next().getStringCellValue());
						}
					}
				}
			}
			
			
					}

	}

}
