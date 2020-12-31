package test;

import java.util.List;
import java.util.Map;

import steps.Excel;

public class ExcelReader {

	public static void main(String[] args) {
		String filePath="C:\\Users\\A714978\\github_workspace\\ExcelReader\\src\\test\\resources\\TestData.xlsx";
		String sheetname="MyInformation";
		List<Map<String,String>> data=Excel.excelReadHashMap(filePath,sheetname);
		
		System.out.println("Excel data as below");
		for(int i=0;i<data.size();i++)
		{
			Map<String,String> map=data.get(i);
			System.out.println("=======================");
			System.out.println("Name :"+map.get("Name"));
			System.out.println("Age :"+map.get("Age"));
			System.out.println("Height :"+map.get("Height"));
			System.out.println("Address :"+map.get("Address"));
			System.out.println("Blood group :"+map.get("Blood group"));
			System.out.println("=======================");

		}

	}

}
