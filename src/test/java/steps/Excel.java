package steps;


import java.io.FileInputStream;
import java.util.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	
	public static List<Map<String,String>> excelReadHashMap(String sExcelPath, String sSheetName) {

		List<Map<String,String>> dataList = new ArrayList<Map<String,String>>(); 
		Map<String,String> rowData = null;
		List<String> colList=new ArrayList<String>();
		List<String> rowList=null;
		
		
		try 
		{
			FileInputStream oFis = new FileInputStream(sExcelPath);
			Workbook workbook = null;
			
			if (sExcelPath.contains(".xlsx"))
			{
				workbook = new XSSFWorkbook(oFis);
			}
			else
			{
				workbook  = new HSSFWorkbook(oFis);
			}
		
			
			Sheet sheet = workbook.getSheet(sSheetName);
			Iterator<Row> rowIterator = sheet.iterator();
			DataFormatter formatter = new DataFormatter(Locale.US);
			while (rowIterator.hasNext())
			{
				Boolean bHeaderRow = false;
				rowList =new ArrayList<String>();
				Row row = rowIterator.next();
				if (row.getRowNum() == 0)
				{
					bHeaderRow = true;
				}

				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					if (bHeaderRow && (cell.getCellType() != Cell.CELL_TYPE_BLANK))
					{
						colList.add(formatter.formatCellValue(cell));
					}
					else if ((!bHeaderRow) && (colList.get(cell.getColumnIndex()) != null))
					{
						if(cell.getCellType() != Cell.CELL_TYPE_BLANK)
						{
							rowList.add(formatter.formatCellValue(cell));
						}
						else
						{
							rowList.add(null);
						}
					}
				}
				if ((colList.size() != 0) && (rowList.size() != 0)) 
				{
					rowData = new LinkedHashMap<String, String>();
					for (int i = 0; i < colList.size(); i++) 
					{
						if (i < rowList.size())
						{
							rowData.put(colList.get(i), rowList.get(i));
						}
						else
						{
							rowData.put(colList.get(i), null);
						}
					}
					dataList.add( rowData );
				}
			}
			workbook.close();
			oFis.close();
		} catch (Exception e) {
			System.out.println("Execption at excelReadHashMap(String sExcelPath, String sSheetName) in Excel.java:\n"+e.getMessage());
		}
		return dataList;
	}
	

}
