package ExcelTests;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.testng.annotations.Test;
import org.testng.Assert;

import ExcelProvider.ExcelFunctions;
public class ExcelTests {
	
	ExcelFunctions excel = new ExcelFunctions();
	
	@Test
	public void Excel_LoadFiles()
	{		
		excel.LoadFile("TestExcel", 0);
		String excelSheetName = ExcelFunctions.GetSheet().getSheetName();
		System.out.println("Excel Sheet Name: " + excelSheetName);
		Assert.assertEquals(excelSheetName,"Test1");
	}
	
	@Test
	public void Excel_GetUsedRowsCount()
	{		
		excel.LoadFile("TestExcel", "Test1");
		int usedRowsCount = ExcelFunctions.GetUsedRowsCount();
		System.out.println("Used Rows Count: " + usedRowsCount);
		Assert.assertEquals(usedRowsCount,11);
	}
	
	@Test
	public void Excel_GetUsedColumnsCount()
	{		
		excel.LoadFile("TestExcel", "Test1");
		int usedColumnsCount = ExcelFunctions.GetUsedColumnsCount();
		System.out.println("Used Columns Count: " + usedColumnsCount);
		Assert.assertEquals(usedColumnsCount,11);
	}
	
	@Test
	public void Excel_GetColumnNameFromIndex()
	{		
		excel.LoadFile("TestExcel", "Test1");
		String columnName = ExcelFunctions.ColumnIndexToColumnAlphabet(33);
		System.out.println("Column Name: " + columnName);
		Assert.assertEquals(columnName,"AG");
	}
	
	@Test
	public void Excel_GetColumnIndexFromName()
	{		
		excel.LoadFile("TestExcel", "Test1");
		int columnNumber = ExcelFunctions.ColumnAlphabetToColumnIndex("AH");
		System.out.println("Column Number: " + columnNumber);
		Assert.assertEquals(columnNumber,34);
	}
	
	@Test
	public void Excel_GetCellValue_Index()
	{		
		excel.LoadFile("TestExcel", "Test1");
		String cellValue = ExcelFunctions.GetCellValue(2,1);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals(cellValue,"2.001");
	}
	
	@Test
	public void Excel_GetCellValue_Range()
	{		
		excel.LoadFile("TestExcel", "Test1");
		String cellValue = ExcelFunctions.GetCellValue("A1");
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals(cellValue,"A");
	}
	
	@Test
	public void Excel_GetCellValue_Cell()
	{				
		excel.LoadFile("TestExcel", "Test1");
		
		Cell cell = ExcelFunctions.GetCell("A1");
		String cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("A", cellValue);
		
		cell = ExcelFunctions.GetCell("A2");
		cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("2.001", cellValue);
		
		cell = ExcelFunctions.GetCell("A3");
		cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("5", cellValue);
		
		cell = ExcelFunctions.GetCell("A4");
		cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("ABC", cellValue);
		
		cell = ExcelFunctions.GetCell("A5");
		cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("XYZ", cellValue);
		
		cell = ExcelFunctions.GetCell("A6");
		cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("2.4987506246876565", cellValue);
		
		cell = ExcelFunctions.GetCell("A7");
		cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("23.75", cellValue);
		
		cell = ExcelFunctions.GetCell("A8");
		cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("20191215 000000", cellValue);
		
		cell = ExcelFunctions.GetCell("A9");
		cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("20191215 000000", cellValue);
		
		cell = ExcelFunctions.GetCell("A10");
		cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("19000121 085248", cellValue);
		
		cell = ExcelFunctions.GetCell("A11");
		cellValue = ExcelFunctions.GetCellValue(cell);
		System.out.println("Cell value: " + cellValue);
		Assert.assertEquals("0.769", cellValue);
	}
	
	@Test
	public void Excel_GetRange_Range() {
		excel.LoadFile("TestExcel", "Test1");
		List<Object> rangeValues = ExcelFunctions.GetRangeValue("A1:A11");
	}
	@Test
	public void Excel_GetCell_Range_Index()
	{		
		excel.LoadFile("TestExcel", "Test1");
		Cell cell = ExcelFunctions.GetCell(5,3);
		System.out.println("Cell row value: " + cell.getRowIndex() + ". \nCell column value: " + cell.getColumnIndex());
	}
	
	@Test
	public void Excel_GetCell_Range_Range()
	{		
		excel.LoadFile("TestExcel", "Test1");
		Cell cell = ExcelFunctions.GetCell("C2");
		System.out.println("Cell row value: " + cell.getRowIndex() + ". \nCell column value: " + cell.getColumnIndex());
		
	}
	
	@Test
	public void Get2DArrayFromSheet() {
		excel.LoadFile("TestExcel", "Test1");
		String[][] Array2DFromSheet= ExcelFunctions.Get2DArrayFromSheet();
		
		Assert.assertEquals("A", Array2DFromSheet[0][0]);
	}
	
	@Test
	public void Get2DArrayFromSheetBySpecificRowsColumns() {
		excel.LoadFile("TestExcel", "Test1");
		String[][] Array2DFromSheet= ExcelFunctions.Get2DArrayFromSheetBySpecificRowsColumns(3,2);
		
		Assert.assertEquals("A", Array2DFromSheet[0][0]);
	}
	
	@Test
	public void GetColumnData() {
		excel.LoadFile("TestExcel", "Test1");
		List<String> ListFromSheet= ExcelFunctions.GetColumnData(1);
		
		Assert.assertEquals("A", ListFromSheet.get(0));
	}
	
	@Test
	public void SetForegroundColor() {
		excel.LoadFile("TestExcel", "Test1");
		ExcelFunctions.SetForegroundColor(3,2);
		ExcelFunctions.SaveWorkbook();
		
		//Assert.assertEquals("A", ListFromSheet.get(0));
	}
	
	@Test
	public void SetForegroundColor_Range() {
		excel.LoadFile("TestExcel", "Test1");
		ExcelFunctions.SetForegroundColor("B4");
		ExcelFunctions.SaveWorkbook();
		
		//Assert.assertEquals("A", ListFromSheet.get(0));
	}
	
	@Test
	public void GetCellForegroundColor() {
		excel.LoadFile("TestExcel", "Test1");
		int foregroundColor = ExcelFunctions.GetCellForegroundColor(4,2);
		
		Assert.assertEquals(foregroundColor,29);
	}
	
	@Test
	public void GetCellForegroundColor_Range() {
		excel.LoadFile("TestExcel", "Test1");
		int foregroundColor = ExcelFunctions.GetCellForegroundColor("B4");
		
		Assert.assertEquals(foregroundColor,29);
	}
	
	@Test
	public void SetCellFontColor() {
		excel.LoadFile("TestExcel", "Test1");
		ExcelFunctions.SetFontColor(4,2,36);
		ExcelFunctions.SaveWorkbook();
	}
	
	@Test
	public void SetCellFontColor_Range() {
		excel.LoadFile("TestExcel", "Test1");
		ExcelFunctions.SetFontColor("B4",36);
		ExcelFunctions.SaveWorkbook();
	}
	
	@Test
	public void GetCellFontColor() {
		excel.LoadFile("TestExcel", "Test1");
		int fontColor = ExcelFunctions.GetCellFontColor(4,2);
		
		Assert.assertEquals(fontColor,36);
	}
	
	@Test
	public void GetCellFontColor_Range() {
		excel.LoadFile("TestExcel", "Test1");
		int fontColor = ExcelFunctions.GetCellFontColor("B4");
		
		Assert.assertEquals(fontColor,36);
	}
	
	@Test
	public void SetCellValue() {
		excel.LoadFile("TestExcel", "Test1");
		ExcelFunctions.SetCellValue(4,2,"Dhamale");
		ExcelFunctions.SaveWorkbook();
	}
	
	@Test
	public void SetCellValue_Range() {
		excel.LoadFile("TestExcel", "Test1");
		ExcelFunctions.SetCellValue("B4","Dhamale");
		ExcelFunctions.SaveWorkbook();
	}
	
	@Test
	public void FindCell() {
		excel.LoadFile("TestExcel", "Test1");
		Cell cell = ExcelFunctions.FindCell("Dhamale");
		
		System.out.println("Cell row value: " + cell.getRowIndex() + ". \nCell column value: " + cell.getColumnIndex());
	}
	
	@Test
	public void BuildRowHeading() {
		excel.LoadFile("TestExcel", "Test1");
		HashMap<String, String> rowHeading = ExcelFunctions.BuildRowHeadingDictionary(1,1);
	}
	
	@Test
	public void GetRangeCoOrdinates() {
		excel.LoadFile("TestExcel", "Test1");
		List<Integer> rangeCoOrdinates = ExcelFunctions.GetRangeCoOrdinates("$A$1:$A$8");
	}
	
	@Test
	public void SetRangeValues() {
		excel.LoadFile("TestExcel", "Test1");
		List<Object> values= new ArrayList();
		values.add(5);
		values.add("Sandeep");
		values.add("Dhamale");
		ExcelFunctions.SetRangeValue("$C$1:$C$3", values);
		ExcelFunctions.SaveWorkbook();
	}
	
	@Test
	public void SetRangeValues1() {
		excel.LoadFile("TestExcel", "Test1");
		List<Object> values= new ArrayList();
		values.add(5);
		values.add("Sandeep");
		values.add("Dhamale");
		ExcelFunctions.SetRangeValue("$D1:$D3", values);
		ExcelFunctions.SaveWorkbook();
	}
	
	@Test
	public void GetCellProperties() {
		excel.LoadFile("TestExcel", "Test1");
		
		HashMap<String, Object> cellProperties = ExcelFunctions.GetCellProperties("A1");
	}
	
	@Test
	public void GetRangeProperties() {
		excel.LoadFile("TestExcel", "Test1");
		
		List<Object> cellProperties = ExcelFunctions.GetRangeProperties("A1:A11");
	}
}
