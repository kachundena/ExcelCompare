package com.tutorialacademy.rest.ExcelCompare;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Hello world!
 *
 */
public class App 
{
    
    public static void main(String[] args) throws IOException {
		try {
			String szFileOri = "prestamo.xlsx";
			String szFileMod = "prestamo2.xlsx";
			// fichero original, el que utilizan los usuarios
			// FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            // Workbook workbook = new XSSFWorkbook(excelFile);
			FileInputStream fileori = new FileInputStream(new File(szFileOri));
			Workbook workbookori = new XSSFWorkbook(fileori);
			Sheet sheetori = workbookori.getSheetAt(0);
			Iterator<Row> rowIteratorori = sheetori.iterator();
			
			// fichero que hemos modificado
			FileInputStream filemod = new FileInputStream(szFileMod);
			XSSFWorkbook workbookmod = new XSSFWorkbook(filemod);
			XSSFSheet sheetmod = workbookmod.getSheetAt(0);
			
			
			int isDifferent = 0;
			int numRow = 0;
			
			CellStyle styleMod = workbookmod.createCellStyle();
			styleMod.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
			styleMod.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
			Row rowori;
			Row rowmod;
			
			
			while (rowIteratorori.hasNext()){
			    rowori = rowIteratorori.next();
			    numRow = rowori.getRowNum();
			    rowmod = sheetmod.getRow(numRow);
		        Iterator<Cell> cellIteori = rowori.cellIterator();
			    isDifferent = 0;
			    while (cellIteori.hasNext()) {
			    	Cell cellori = (Cell) cellIteori.next();
		            int numCol = cellori.getColumnIndex();
		            Cell cellmod = rowmod.getCell(numCol);
		            if (cellori != null && cellmod != null)
		            {
						switch(cellori.getCellTypeEnum()) {
							case NUMERIC:
							    if( DateUtil.isCellDateFormatted(cellori) ){
							    	if (!cellori.getDateCellValue().equals(cellmod.getDateCellValue())) {
							    		cellmod.setCellStyle(styleMod);
							    		isDifferent = isDifferent + 1;
							    	}	    	
							    }
							    else{
							    	if (cellori.getNumericCellValue() != cellmod.getNumericCellValue()) {
							    		cellmod.setCellStyle(styleMod);
							    		isDifferent = isDifferent + 1;
							    	}
							    }
							    break;
							case STRING:
						    	if (!cellori.getStringCellValue().equals(cellmod.getStringCellValue())) {
						    		cellmod.setCellStyle(styleMod);
						    		isDifferent = isDifferent + 1;
						    	}
							    break;
							case BOOLEAN:
						    	if (cellori.getBooleanCellValue() != cellmod.getBooleanCellValue()) {
						    		cellmod.setCellStyle(styleMod);
						    		isDifferent = isDifferent + 1;
						    	}
								break;
							default:
								break;							
						}
		            }
		        }

			}
			FileOutputStream filemodsave = new FileOutputStream(szFileMod);
			workbookmod.write(filemodsave);
			workbookori.close();
			workbookmod.close();
		}
		 catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }		
		}



}
