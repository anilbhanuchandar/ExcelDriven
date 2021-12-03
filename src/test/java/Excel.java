import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public static void main(String[] args) throws IOException 
	{
		
		FileInputStream fis = new FileInputStream("C:\\Users\\Anil kumar V\\OneDrive\\Desktop\\Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheets = workbook.getNumberOfSheets();
		
		for(int i=0;i<sheets;i++)
		{
			
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata"))
			{
			XSSFSheet sheet = workbook.getSheetAt(i);
			//Identify Testcases coloum by scanning the entire 1st row
			Iterator<Row> row = sheet.iterator();
			Row firstrow = row.next();
			Iterator<Cell> cell = firstrow.cellIterator();
			int column=0;
			int k=0;

			////once coloumn is identified then scan entire testcase coloum to identify purcjhase testcase row
			while(cell.hasNext())
			{
				
			        Cell ce = cell.next();
			        if(ce.getStringCellValue().equalsIgnoreCase("testcase"))
			        {
			        	column=k;
			        }
			        k++;
			}
			System.out.println(column);
			
			
			while(row.hasNext())
			{
				Row ce = row.next();
				if(ce.getCell(column).getStringCellValue().equalsIgnoreCase("login"))
				{
				////after you grab purchase testcase row = pull all the data of that row and feed into test
					Iterator<Cell> cll = ce.cellIterator();
					while(cll.hasNext())
					{
						System.out.println(cll.next().getStringCellValue());
					}
				
				}
			}
			
			}
			
		}

	}

}
