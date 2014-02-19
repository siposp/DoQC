package doqc;

import java.io.IOException;
import java.io.InputStream;


import java.io.OutputStream;
import java.io.OutputStreamWriter;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * This class provides functionality to convert a given XLS, XLSX document to the ARFF format.
 * The class functions heavily depend on the Apache POI library,that provides access to spreadsheet cells.
 * The POI library differentiate 6 cell types: BLANK, ERROR, NUMERIC, STRING, BOOLEAN, FORMULA 
 * For these types this library acts as follow:
 * BLANK,ERROR - stops current process and yields an exception if blank cell is not appropriate in the particular state e.g. in the middle of data block
 * NUMERIC - written out in the format of default Java double to string conversion
 * STRING - written out as is
 * BOOLEAN - written out in the format of default Java bool to string conversation
 * FORMULA - no formula evaluation is done during the ARFF conversation, the cached formula results are used (Cached formula results are stored in the document to make quickly available the document after opening )
 * 
 * PLEASE NOTE!
 * In this version I assumed, that all FORMULA cells have cached values, which can be used.    
 *    
 * Numeric cells contain data in double data type. For presumably integer data (for which the condition data-(int)data == 0 is true) truncation after decimal point is applied. This behaviour can be change by setNumericTruncation function
 * During the conversion some standard IOExceptions, one DoQ-specific NullDataCellException and one POI-specific InvalidFormatException can be thrown
 *    
 * The conversion starts with the first row containing something. From left to the right the ARFF header will contain all non-null cells. The width of header defines the width of data block. Valid data block should end with a full blank row under the header.   
 */
public class DoQConverter {
	
	public static final char DEFAULT_ARFF_HEADER_CHAR='#';
	public static final char DEFAULT_ARFF_NEW_LINE_CHAR='\n';
	public static final char DEFAULT_ARFF_DELIMITER_CHAR=';';
	
	String path;
	Workbook wb=null;
	InputStream in=null;
	
	public DoQConverter(InputStream in){
		this.in=in;
	}
	
	public void writeFirstBlockToARFF(OutputStream destination) throws InvalidFormatException, IOException, NullDataCellException{
		writeFirstBlockToARFF(destination,';',DoQConverter.DEFAULT_ARFF_NEW_LINE_CHAR,DoQConverter.DEFAULT_ARFF_HEADER_CHAR);
	}
	
	public void writeFirstBlockToARFF(OutputStream destination,char arffDelim) throws InvalidFormatException, IOException, NullDataCellException{
		writeFirstBlockToARFF(destination, arffDelim, DoQConverter.DEFAULT_ARFF_NEW_LINE_CHAR,DoQConverter.DEFAULT_ARFF_HEADER_CHAR);
	}
	
	public void writeFirstBlockToARFF(OutputStream destination,char arffDelim,char arffNewLine,char arffHeader) throws IOException, InvalidFormatException, NullDataCellException{
		//Open the document input stream
		this.open();
		OutputStreamWriter osw=new OutputStreamWriter(destination);
		
		Sheet sheet = wb.getSheetAt(0);
		if(sheet==null){
			throw new NullDataCellException();
		}
		
		int firstRow=sheet.getFirstRowNum();
		int firstColumn=sheet.getLeftCol();
		
	    
	    int rowNum=firstRow;
	    Row currentRow=sheet.getRow(firstRow);
	    if(currentRow==null){
	    	throw new NullDataCellException();
	    }
	    
	    int headerWidth=write(osw, arffDelim, arffNewLine, arffHeader, currentRow, firstColumn,true,-1);
	    
	    
	    for(int i=firstColumn;i<firstColumn+headerWidth;i++){
	    	
	    }
	    
	    
	    int rowWidth=headerWidth;
	    //for every full width row go to the next
	    //maximize the width - to the header width
	    while(rowWidth==headerWidth){
	    	rowNum++;
	    	currentRow=sheet.getRow(rowNum);
	    	rowWidth=write(osw, arffDelim, arffNewLine, arffHeader, currentRow, firstColumn,false,headerWidth);
	    }
	    
	    //the row width can 
	    //if 0 cells were written check whether it is a blank row or some invalid cells
	    
	    int nonBlankCellsNum=countNonBlankCells(currentRow, firstColumn, headerWidth);
	    if((nonBlankCellsNum>0)&&(nonBlankCellsNum<headerWidth)){
	    	throw new NullDataCellException();
	    }
	    
	    
		osw.flush();
		
		//Close the opened document
		this.close();
	}
	
	
	private int countNonBlankCells(Row row,int startCol,int colCount){
		int count=0;
		//if we have a non-empty row, we count the non blank cells
		if(row!=null){
		Cell cell=null;
		for(int i=startCol;i<startCol+colCount;i++){
			cell=row.getCell(i);
			if(cell==null||cell.getCellType()==Cell.CELL_TYPE_BLANK||cell.getCellType()==Cell.CELL_TYPE_ERROR){
				
			}else{
				count++;
			}
		}
		}
		
		return count;
	}
	
	
	
	/**
	 * This function writes the given row to the destination stream as an ARFF header. It writes all non-null and valid cells and returns the number of cells written.
	 * @param osw Destination stream
	 * @param arffDelim ARFF delimiter character
	 * @param arffNewLine
	 * @param arffHeader
	 * @param startRow The row that contains data for the header
	 * @param startColumn The number of column from which should we start
	 * @param limit Maximal number of columns to be written, -1 is unlimited
	 * @return Number of columns written to the header
	 * @throws IOException 
	 */
	private int write(
			OutputStreamWriter osw,
			char arffDelim,
			char arffNewLine,
			char arffHeader,
			Row startRow,
			int startColumn,
			boolean isHeader,
			int limit) throws IOException
		{
			if(startRow==null){
				return 0;
			}
	    	Cell cell=startRow.getCell(startColumn);
	    	int column=startColumn;
	    	while(
	    			cell!=null 
	    			&& cell.getCellType()!=Cell.CELL_TYPE_BLANK
	    			&& cell.getCellType()!=Cell.CELL_TYPE_ERROR
	    			)
	    	{
	    		//Add the Header character before the first item
	    		if(column==startColumn && isHeader){
	    			osw.write(arffHeader);
	    		}
	    		//Add the ARFF delimiter character before every next item
	    		if(column!=startColumn){
	    			osw.write(arffDelim);
	    		}
	    		
	    		writeCell(osw, cell);
	    		
	    		
	    		column++;
	    		if ((limit!=-1)&&(column==limit)) break;
	    		
	    		cell=startRow.getCell(column);
	    	}
	    	
	    	//At least one cell was added to the header, so we add a NewLine delimiter after the line
	    	if(column!=startColumn){
	    		osw.write(arffNewLine);
	    	}
		
		
		return (column-startColumn);
	}
	
	
	private void writeCell(OutputStreamWriter osw,Cell cell) throws IOException{
		switch(cell.getCellType()){
		case Cell.CELL_TYPE_NUMERIC: 
			if(numericTruncationEnabled){
				double cellVal=cell.getNumericCellValue();
				int intCellVal=(int)cellVal;
				if((cellVal-intCellVal) == 0){
					osw.write(""+intCellVal);
				}else{
					osw.write(""+cell.getNumericCellValue());
				}
			}else{
				osw.write(""+cell.getNumericCellValue());
			}
			break;
		case Cell.CELL_TYPE_STRING: 
			osw.write(cell.getStringCellValue()); 
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			osw.write(""+cell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA:
			osw.write(""+cell.getCachedFormulaResultType());
			break;
	}
	}
	
	
	private void open() throws InvalidFormatException, IOException{
		wb = WorkbookFactory.create(in);
	}
	
	
	/**
	 * Closes the opened InputStream in, that was opened for document reading
	 * @throws IOException
	 */
	private void close() throws IOException{
		in.close();
	}
	
	boolean numericTruncationEnabled=true;
	
	/**
	 * Sets the behaviour of numeric cell conversion
	 * Cell containing 123 is converted as follow:
	 * 123 - numeric truncation is enabled (default behaviour)
	 * 123.0 - numeric truncation is disabled
	 * 
	 * Numbers with non-zero values after decimal point are not affected by this settings
	 * 
	 * Please note, that double-integer comparison is used for decision.
	 * 
	 * @param enabled true for enable numeric truncation, false to disable numeric truncation
	 */
	public void setNumericTruncation(boolean enabled){
		this.numericTruncationEnabled=enabled;
	}
	
	/**
	 * Returns the boolean value representing the numeric truncation state 
	 * @return true enabled numeric truncation, false for disabled state 
	 */
	public boolean getNumericTruncation(){
		return this.numericTruncationEnabled;
	}
	
	
	
}
