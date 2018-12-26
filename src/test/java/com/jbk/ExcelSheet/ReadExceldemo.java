package com.jbk.ExcelSheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
public class ReadExceldemo {
	public static void main(String[] args) throws IOException {
		File fl=new File("F:\\Selenium prgm\\Mayur.xls");
		FileInputStream fi=new FileInputStream(fl);

		
		HSSFWorkbook  wk=new HSSFWorkbook(fi);
		
		HSSFSheet st=wk.getSheetAt(0);
		
		
		
		int Rowcount=st.getLastRowNum()-st.getFirstRowNum();
		for (int i = 0; i < Rowcount+1; i++) {
			
			
			 Row row = st.getRow(i);
			 
			

		        
		        for (int j = 0; j < row.getLastCellNum(); j++) {

		            

		            System.out.print(row.getCell(j)+"|| ");

		        }

		        System.out.println();

		    }
			
		
	}
	
}

