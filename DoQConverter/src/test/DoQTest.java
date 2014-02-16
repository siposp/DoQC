package test;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import doqc.DoQConverter;

public class DoQTest {
	public static void main(String[] args){
		//DoQConverter doqc=new DoQConverter();
		if(args.length==1){
			System.out.println("Input file:"+args[0]);
			
			String inputFile=args[0];
			//String outputFile=inputFile.substring(0,inputFile.indexOf(".xls"))+"_test.xls";
			
			
			DoQConverter doqc=new DoQConverter(inputFile);
			System.out.println("Parsing file: "+inputFile);
			
			System.out.println("Opening...");
			try {
				doqc.open();
			} catch (InvalidFormatException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			for(int i=0;i<8;i++){
				doqc.printRow(i, 5);
			}
			
			System.out.println("Closing...");
			try {
				doqc.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			
			System.out.println("End of testing...");
		}else{
			System.err.println("Please specify an input file");
		}
	}
}
