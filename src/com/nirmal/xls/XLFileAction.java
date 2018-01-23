
package com.nirmal.xls;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.nirmal.xls.model.Employee;

/**
 * @author muthu_m
 *
 */
public class XLFileAction 
{

	private static String FILE_NAME = "xl-file.xlsx";
	private static String[] columns = { "Name", "Email", "Date of Birth", "Salary" };

	private static List<Employee> listEmployee = new ArrayList<>();

	// Initializing employees data to insert into excel file
	static 
	{
		Calendar dateOfBirth = Calendar.getInstance();

		// Record 1
		dateOfBirth.set(1992, 07, 21);
		listEmployee.add(new Employee("Retna A", "retna@example.com", dateOfBirth.getTime(), 1200000.0));

		// Record 2
		dateOfBirth.set(1991, 04, 18);
		listEmployee.add(new Employee("Muthukumar M", "muthu@example.com", dateOfBirth.getTime(), 1500000.0));

		// Record 3
		dateOfBirth.set(1993, 06, 02);
		listEmployee.add(new Employee("James M", "james@example.com", dateOfBirth.getTime(), 1800000.0));
	}

	/**
	 * Let’s now look at the program to create and write to an excel file. 
	 * Note that I’ll be using an XSSFWorkbook to create a Workbook instance. This will generate the newer XML based excel file (.xlsx). 
	 * You may choose to use HSSFWorkbook if you want to generate the older binary excel format (.xls)
	 */
	public static void writeExcelSheet()
	{
		// Create Workbook
		Workbook workbook = new XSSFWorkbook();
		/**
		 * CreationHelper helps us create instances for various things like DataFormat,
		 * Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way
		 * 
		 */
		CreationHelper createHelper = workbook.getCreationHelper();

		// Create Sheet
		Sheet sheet = workbook.createSheet("Employee");

		// Create a font for styling header cells
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setColor(IndexedColors.RED.getIndex());

		// Create a CellStyle with the font
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Create a Row
		Row headerRow = sheet.createRow(0);

		// Creating cells
		for (int i = 0; i < columns.length; i++) 
		{
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}

		// Create Cell Style for formatting Date
		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

		// Create Other rows and cells with employees data
		int rowNum = 1;
		for (Employee employee : listEmployee)
		{
			Row row = sheet.createRow(rowNum++);

			row.createCell(0).setCellValue(employee.getName());

			row.createCell(1).setCellValue(employee.getEmail());

			Cell dateOfBirthCell = row.createCell(2);
			dateOfBirthCell.setCellValue(employee.getDateOfBirth());
			dateOfBirthCell.setCellStyle(dateCellStyle);

			row.createCell(3).setCellValue(employee.getSalary());
		}

		// Resize all columns to fit the content size
		for (int i = 0; i < columns.length; i++) 
		{
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut;
		try 
		{
			fileOut = new FileOutputStream(FILE_NAME);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch (FileNotFoundException e) 
		{
			e.printStackTrace();
		}
		catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * The following program shows you how to read an excel file using Apache POI. 
	 * Since we’re not using any file format specific POI classes, the program will work for both types of file formats - .xls and .xlsx.
	 * The program shows two different ways of iterating over sheets, rows, and columns in the excel file -
	 */
	
	public static void readXLSheet()
	{
		// Creating a Workbook from an Excel file (.xls or .xlsx)
		try 
		{
			Workbook workbook = WorkbookFactory.create(new File(FILE_NAME));
			 // Retrieving the number of sheets in the Workbook
	        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets ");
	        
	        // 1.you can use a for-each loop
	        for (Sheet sheet2 : workbook) 
	        {
				System.out.println("The available sheet name : " + sheet2.getSheetName());
			}
	        
	        System.out.println("===============================================================================");
	        
	        // 2. Or you can use a Java 8 forEach with lambda
	        workbook.forEach(sheet ->
	        {
	        	System.out.println("The avaliable sheet name : " + sheet.getSheetName());
	        });
	        
	        System.out.println("===============================================================================");
	        
	        // Getting the Sheet at index zero
	        Sheet sheet = workbook.getSheetAt(0);
	        
	        // Create a DataFormatter to format and get each cell's value as String
	        DataFormatter dataFormatter = new DataFormatter();
	        
	        // 1.you can use a for-each loop
	        
	        for (Row row : sheet)
	        {
				for (Cell cell2 : row)
				{
					String cellValue = dataFormatter.formatCellValue(cell2);
	                System.out.print(cellValue + "\t");
				}
				 System.out.println();
			}
	        
	        System.out.println("===============================================================================");
	        
	        // 2. Or you can use a Java 8 forEach with lambda
	        sheet.forEach(row ->
	        {
	        	row.forEach(cell ->
	        	{
	        		String cellValue = dataFormatter.formatCellValue(cell);
	                System.out.print(cellValue + "\t");
	        	});
	        	 System.out.println();
	        });
	        
		}
		catch (IOException e) 
		{
			e.printStackTrace();
		}
		catch (EncryptedDocumentException e) 
		{
			e.printStackTrace();
		}
		catch (InvalidFormatException e) 
		{
			e.printStackTrace();
		}
	}

	public static void main(String[] args) 
	{
		XLFileAction.writeExcelSheet();
		XLFileAction.readXLSheet();
	}

}
