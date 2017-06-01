package com.kachundena.ExcelCompare;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Hello world!
 *
 */
public class App 
{
    public CellStyle styleMod;
    public void main(String[] args) throws IOException {
			String szFile1 = "prestamo.xlsx";
			String szFile2 = "prestamo2.xlsx";
			int iReturn = checkEqualWorkSheets (szFile1, szFile2);
			System.out.println("Fin. Número de diferencias: " + iReturn);
	}
    
    public int checkEqualWorkSheets(String szFile1, String szFile2) throws IOException {
    	int iReturn = 0;
    	try {
    		// get File 1
			FileInputStream fileWS1 = new FileInputStream(new File(szFile1));
			XSSFWorkbook wb1 = new XSSFWorkbook(fileWS1);
			
			// get File 2
			FileInputStream fileWS2 = new FileInputStream(szFile2);
			XSSFWorkbook wb2 = new XSSFWorkbook(fileWS2);

			CellStyle styleMod = wb2.createCellStyle();
			styleMod.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
			styleMod.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
			iReturn = checkEqualWorkBook(wb1, wb2, iReturn);
			FileOutputStream filemodsave = new FileOutputStream(szFile2);
			wb2.write(filemodsave);
			wb1.close();
			wb2.close();
    	}
		catch (FileNotFoundException e) {
            e.printStackTrace();
        } 
    	catch (IOException e) {
    		e.printStackTrace();
    	}  	
    	return iReturn;

    }
    
    private int checkEqualWorkBook(XSSFWorkbook pwb1, XSSFWorkbook pwb2, int piReturn) {
    	try {
    		
    		for (int iNumSheets = 0; iNumSheets < pwb1.getNumberOfSheets(); iNumSheets++) {
    			Sheet sh1 = pwb1.getSheetAt(0);   
    			Sheet sh2 = pwb2.getSheetAt(0);
    			piReturn = checkEqualSheet(sh1, sh2, piReturn);    		}
    	}
    	catch (Exception e) {
    		piReturn += 1;
    		e.printStackTrace();
    	}	
    	return piReturn;
    }

    private int checkEqualSheet(Sheet psh1, Sheet psh2, int piReturn) {
    	try {
			Iterator<Row> rowIterator1 = psh1.iterator();
			while (rowIterator1.hasNext()){
				Row row1 = rowIterator1.next();
				Row row2 = psh2.getRow(row1.getRowNum());
				piReturn = checkEqualRow(row1, row2, piReturn);
			}  		
    	}
    	catch (Exception e) {
    		piReturn += 1;
    		e.printStackTrace();
    	}
    	return  piReturn;
    }
    
    private int checkEqualRow (Row prw1, Row prw2, int piReturn) {
    	try {
	        Iterator<Cell> cellIterator1 = prw1.cellIterator();
		    while (cellIterator1.hasNext()) {
		    	Cell cell1 = (Cell) cellIterator1.next();
		    	Cell cell2 = prw2.getCell(cell1.getColumnIndex());
		    	piReturn = checkEqualCell (cell1, cell2, piReturn);
		    }
    	}
    	catch (Exception e) {
    		piReturn += 1;
    		e.printStackTrace();
    	}
    	return piReturn;
    }
    	
    private int checkEqualCell (Cell pcl1, Cell pcl2, int piReturn) {
    	try {
    		if (pcl1 != null && pcl2 != null)
            {
				switch(pcl1.getCellTypeEnum()) {
					case NUMERIC:
					    if( DateUtil.isCellDateFormatted(pcl1) ){
					    	if (!pcl1.getDateCellValue().equals(pcl2.getDateCellValue())) {
					    		pcl2.setCellStyle(styleMod);
					    		piReturn += 1;
					    	}	    	
					    }
					    else{
					    	if (pcl1.getNumericCellValue() != pcl2.getNumericCellValue()) {
					    		pcl2.setCellStyle(styleMod);
					    		piReturn += 1;
					    	}
					    }
					    break;
					case STRING:
				    	if (!pcl1.getStringCellValue().equals(pcl2.getStringCellValue())) {
				    		pcl2.setCellStyle(styleMod);
				    		piReturn += 1;
				    	}
					    break;
					case BOOLEAN:
				    	if (pcl1.getBooleanCellValue() != pcl2.getBooleanCellValue()) {
				    		pcl2.setCellStyle(styleMod);
				    		piReturn += 1;
				    	}
						break;
					default:
						break;							
				}
            }    		
    	}
    	catch (Exception e) {
    		piReturn += 1;
    		e.printStackTrace();
    	}
    	return piReturn;
    }
}
