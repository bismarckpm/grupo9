import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.JOptionPane;

import java.io.FileNotFoundException;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 * 
 * @author www.codejava.net
 *
 */
public class ExcelFileUpdateExample1 {

	public static void menu(){
		switch(JOptionPane.showInputDialog("1- VALIDAR ARCHIVO EXISTENTE\n2- CANTIDAD DE REGISTROS POR HOJA\n3- ACTUALIZAR REGISTRO")){
			case "1":
				break;
			case "2":
				break;
			case "3":
				Integer num = Integer.parseInt(JOptionPane.showInputDialog("INGRESE EL NUMERO DE IDENTIFICACION"));
				String author = JOptionPane.showInputDialog("INGRESE EL AUTOR");
				Integer price = Integer.parseInt(JOptionPane.showInputDialog("INGRESE EL PRECIO"));
				update(num, author, price);
				break;
			default:
				JOptionPane.showMessageDialog(null, "OPCION NO VALIDA", "ERROR", JOptionPane.ERROR_MESSAGE);
		};
	}

	public static void update(Integer num, String author, Integer price){

		try{

		String excelFilePath = "Inventario.xlsx";

		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		Workbook workbook = WorkbookFactory.create(inputStream);

		Sheet sheet = workbook.getSheetAt(0);
		Boolean b = false;
		String res = "";

		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			if((int)sheet.getRow(i).getCell(0).getNumericCellValue() == num){
				sheet.getRow(i).getCell(2).setCellValue(author);
				sheet.getRow(i).getCell(3).setCellValue(price);
				b = true;
			}
			res += (int)sheet.getRow(i).getCell(0).getNumericCellValue() + " | " + sheet.getRow(i).getCell(1) + " | " + sheet.getRow(i).getCell(2) + " | " + (int)sheet.getRow(i).getCell(3).getNumericCellValue() + "\n";
		}

		inputStream.close();

		FileOutputStream outputStream = new FileOutputStream(excelFilePath);
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();

		if(!b)
		JOptionPane.showMessageDialog(null, "NO SE ENCONTRO EL ID", "Error", JOptionPane.ERROR_MESSAGE);
		
		JOptionPane.showMessageDialog(null, res, "RESULTADO", JOptionPane.INFORMATION_MESSAGE);

		}catch(FileNotFoundException ex){
			ex.printStackTrace();
		}
		catch(IOException ex){
			ex.printStackTrace();
		}
		catch(InvalidFormatException ex){
			ex.printStackTrace();
		}

	}

	public static void main(String[] args) {
		String excelFilePath = "Inventario.xlsx";

		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = WorkbookFactory.create(inputStream);

			Sheet sheet = workbook.getSheetAt(0);

			Object[][] bookData = {
					{"El que se duerme pierde", "Tom Peter", 16},
					{"Sin lugar a duda", "Ana Gutierrez", 26},
					{"El arte de dormir", "Nico", 32},
					{"Buscando a Nemo", "Humble Po", 41},
			};

			int rowCount = sheet.getLastRowNum();

			for (Object[] aBook : bookData) {
				Row row = sheet.createRow(++rowCount);

				int columnCount = 0;
				
				Cell cell = row.createCell(columnCount);
				cell.setCellValue(rowCount);
				
				for (Object field : aBook) {
					cell = row.createCell(++columnCount);
					if (field instanceof String) {
						cell.setCellValue((String) field);
					} else if (field instanceof Integer) {
						cell.setCellValue((Integer) field);
					}
				}

			}

			inputStream.close();

			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();

			menu();
			
		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException ex) {
			ex.printStackTrace();
		}	
	}
}
