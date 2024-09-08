package com.exercise.read_write_excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File; 

public class ReadWriteEvaluateExcel {

	public static void main(String[] args) {
		try
		{
			double marathi=55,hindi=65,english=68,science=75,math=88,geography=77,history=66;
			String filePath="D:\\files\\Best of five.xlsx";
//			File file = new File(filePath);
			FileInputStream fis = new FileInputStream(new File(filePath));
			Workbook workbook = new XSSFWorkbook(fis);
			Sheet sheet = workbook.getSheetAt(0);
            
            double best_of_five=sheet.getRow(2).getCell(6).getNumericCellValue(); //Best of five marks
            System.out.println("Before update :"+best_of_five);
            
            sheet.getRow(1).getCell(1).setCellValue(marathi); 
            sheet.getRow(2).getCell(1).setCellValue(hindi); 
            sheet.getRow(3).getCell(1).setCellValue(english); 
            sheet.getRow(4).getCell(1).setCellValue(science); 
            sheet.getRow(5).getCell(1).setCellValue(math); 
            sheet.getRow(6).getCell(1).setCellValue(geography); 
            sheet.getRow(7).getCell(1).setCellValue(history); 
            
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
	         
	         
	         evaluator.evaluateFormulaCell(sheet.getRow(8).getCell(1));// suppose your formula is in B9
	         evaluator.evaluateFormulaCell(sheet.getRow(8).getCell(2));// suppose your formula is in C9
//	         evaluator.evaluateFormulaCell(sheet.getRow(2).getCell(4));// suppose your formula is in E3
	         evaluator.evaluateFormulaCell(sheet.getRow(1).getCell(6));// suppose your formula is in G2
	         
	         
	      // Read values from B2:B8 and store them in a list
	            List<Double> values = new ArrayList<>();
	            for (int i = 1; i <= 7; i++) { // B2 is index 1, B8 is index 7
	                Row row = sheet.getRow(i);
	                if (row != null) {
	                    Cell cell = row.getCell(1); // B column is index 1
	                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
	                        values.add(cell.getNumericCellValue());
	                    }
	                }
	            }

	            // Sort values in descending order
	            Collections.sort(values, Collections.reverseOrder());

	            // Sum the top 5 largest values
	            double sumOfTop5 = 0;
	            for (int i = 0; i < Math.min(5, values.size()); i++) {
	                sumOfTop5 += values.get(i);
	            }

	            // Write the result back to Excel (e.g., in cell C1)
	            sheet.getRow(2).getCell(4).setCellValue(sumOfTop5); //suppose your formula is in E3
	            evaluator.evaluateFormulaCell(sheet.getRow(2).getCell(6));// suppose your formula is in G3
	            
            fis.close();
				
			  FileOutputStream outFile =new FileOutputStream(filePath);
			  workbook.write(outFile); outFile.close();
				 
            
            workbook.close();
            
            Thread.sleep(1000);
            
//            FileInputStream inputStream1 = new FileInputStream(filePath);

            XSSFWorkbook workbook1 = new XSSFWorkbook(filePath);
            XSSFSheet sheet1 = workbook1.getSheetAt(0);
            
            best_of_five=sheet1.getRow(2).getCell(6).getNumericCellValue(); //Best of five marks
            System.out.println("After update :"+best_of_five);
            
            workbook1.close();
		}
		catch(Exception e){
			e.printStackTrace();
		}
	}

}
