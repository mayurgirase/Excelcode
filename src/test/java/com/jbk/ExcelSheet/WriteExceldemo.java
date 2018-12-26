package com.jbk.ExcelSheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("unused")
public class WriteExceldemo {
	public static void main(String[] args) throws IOException {
		File f= new File("F:\\Selenium prgm\\Mayur3.xlsx");
		 System.out.println(1);
		 FileInputStream fi=new FileInputStream(f);
		System.out.println(2);
		XSSFWorkbook sw=new XSSFWorkbook(fi);
		
		 System.out.println(3);
		 XSSFSheet st=sw.getSheetAt(0);
		
		  System.out.println(4);
		   st.getRow(0).createCell(0).setCellValue("data");
		  
		  System.out.println(5);
		  FileOutputStream fo=new FileOutputStream(f);
		  sw.write(fo);
		  System.out.println(6);
		  sw.close();
	}
	
		
	}



			

