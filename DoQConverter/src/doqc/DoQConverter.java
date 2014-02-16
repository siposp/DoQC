package doqc;

import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;


import org.apache.poi.hssf.usermodel.examples.CellTypes;
//import org.apache.poi.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class DoQConverter {
	String path;
	Workbook wb=null;
	InputStream in=null;
	
	public DoQConverter(String input){
		this.path=input;
	}
	
	public void open() throws InvalidFormatException, IOException{
		in=new FileInputStream(path);
		wb = WorkbookFactory.create(in);
	}
	
	
	public void printRow(int rowNum,int columnCount){
		Sheet sheet = wb.getSheetAt(0);
	    Row row = sheet.getRow(rowNum);
	    System.out.println("Row no. "+rowNum);
	    if(row!=null){
	    for(int i=0;i<columnCount;i++){
	    	Cell cell=row.getCell(i);
	    	if(cell==null){
	    		System.out.print("nil");
	    	}else{
	    		switch(cell.getCellType()){
	    		case Cell.CELL_TYPE_NUMERIC: System.out.print(cell.getNumericCellValue()); break;
	    		case Cell.CELL_TYPE_STRING: System.out.print(cell.getStringCellValue()); break;
	    		}
	    	}
	    	if(i<columnCount-1){
	    		System.out.print(":");
	    	}
	    }
	    System.out.println();
	    }else{
	    	System.out.println("No data in the row");
	    }
	}
	
	public void close() throws IOException{
		in.close();
	}
	
	public void readAndWrite(String input,String output){
//		try {
//			
//		   
//		    Cell cell = row.getCell(3);
//		    if (cell == null)
//		        cell = row.createCell(3);
//		    cell.setCellType(Cell.CELL_TYPE_STRING);
//		    cell.setCellValue("a test");
//
//		    // Write the output to a file
//		    FileOutputStream fileOut = new FileOutputStream(output);
//		    wb.write(fileOut);
//		    fileOut.close();
//			
//			
//			
//			
//			in.close();
//		} catch (FileNotFoundException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (InvalidFormatException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
	}
	
	
	
}
