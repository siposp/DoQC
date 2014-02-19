package test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import doqc.DoQConverter;
import doqc.NullDataCellException;

public class DoQTest {
	public static void main(String[] args) throws FileNotFoundException{
		//DoQConverter doqc=new DoQConverter();
		if(args.length==1){
			System.out.println("Input file:"+args[0]);
			
			String inputFile=args[0];
			//String outputFile=inputFile.substring(0,inputFile.indexOf(".xls"))+"_test.xls";
			
			
			DoQConverter doqc=new DoQConverter(new FileInputStream(inputFile));
			System.out.println("Parsing file: "+inputFile);
			
			System.out.println("Opening...");
			try {
				doqc.writeFirstBlockToARFF(System.out);
			} catch (InvalidFormatException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (NullDataCellException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			System.out.println("End of testing...");
		}else{
			System.err.println("Please specify an input file...\nUsage: DoQTest <input_file>");
		}
	}
}
