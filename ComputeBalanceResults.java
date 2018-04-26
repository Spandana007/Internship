package demo;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ComputeBalanceResults 
{
	 
	
    

    
    
	
 
    
   
	
  // private ArrayList<CustomerVO> firstList;

public void computeBalanceToExcelFile(ArrayList<CustomerVO> list, ArrayList<CustomerVO> list1, ArrayList<CustomerVO> list2) throws IOException
{
        
	
		
       
	 
         writeExcelID(list,1,0);
         writeExcelBalDB(list1,1,1);
         writeExcelBalEx(list2,1,2);
     
       // firstList.addAll(list);
       // firstList.addAll(list1);
}      
            XSSFWorkbook workbook = new XSSFWorkbook();

//Create a blank sheet
           XSSFSheet sheet = workbook.createSheet("computeBalanceResults");
           XSSFSheet ExcelWSheet = sheet;
           XSSFCell Cell;
           XSSFRow Row;
          public void writeExcelID(ArrayList<CustomerVO> list,int RowNum ,int ColNum) throws IOException
          {
            //Blank workbook
        	
            int size1=list.size();
            try
            {
            	for(int i=RowNum-1; i<size1; i++)
            	{
                  
            	  Row = sheet.getRow(i+1); 
            	  if(Row == null)
            	  {
            	    Row = sheet.createRow(i+1);
            	  }
            	  if(Row.getCell(ColNum)==null)
            	  {
            	      Cell = Row.createCell(ColNum);
            	      Cell.setCellValue(list.get(i).getCustId());
            	  }
            	  
            	  else
            	  {
            	       Row.getCell(ColNum).setCellValue(list.get(i).getCustId()); // this method returns a date..
            	  }
            	 }
            

                    FileOutputStream out = new FileOutputStream(new File("C:/Users/688714/Desktop/Results.xlsx"));
                    workbook.write(out);
                    out.close();
                    System.out.println("Results.xlsx written successfully on disk.");
                
            }
                catch (NullPointerException e)
                {
                    e.printStackTrace();
                }
            
               
        }
        public void writeExcelBalDB(ArrayList<CustomerVO> list1,int RowNum ,int ColNum) throws IOException
        {
            //Blank workbook
        	
            int size1=list1.size();
            try
            {
            	for(int i=RowNum-1; i<size1; i++)
            	{

            	  Row = sheet.getRow(i+1); 
            	  if(Row == null)
            	  {
            	    Row = sheet.createRow(i+1);
            	  }
            	  if(Row.getCell(ColNum)==null)
            	  {
            	      Cell = Row.createCell(ColNum);
            	      Cell.setCellValue(list1.get(i).getComputedBalance());
            	  }
            	  else
            	  {
            		  System.out.println(list1.get(i).getComputedBalance()); 
            		  Row.getCell(ColNum).setCellValue(list1.get(i).getComputedBalance()); 
            	  }
            	 }
            

                    FileOutputStream out = new FileOutputStream(new File("C:/Users/688714/Desktop/Results.xlsx"));
                    workbook.write(out);
                    out.close();
                    System.out.println("Results.xlsx written successfully on disk.");
                
            }
                catch (NullPointerException e)
                {
                    e.printStackTrace();
                }
            

        }
        public void writeExcelBalEx(ArrayList<CustomerVO> list2,int RowNum ,int ColNum) throws IOException
        {
            //Blank workbook
        	
            int size1=list2.size();
            try
            {
            	for(int i=RowNum-1; i<size1; i++)
            	{

            	  Row = sheet.getRow(i+1); 
            	  if(Row == null)
            	  {
            	    Row = sheet.createRow(i+1);
            	  }
            	  if(Row.getCell(ColNum)==null)
            	  {
            		  
            	      Cell = Row.createCell(ColNum);
            	     
            	      Cell.setCellValue(list2.get(i).getComputedBalance());
            	  }
            	  else
            	  {
            		   
            		   
            		   
            	       Row.getCell(ColNum).setCellValue(list2.get(i).getComputedBalance()); 
            		  // this method returns a date..
            	  }
            	 }
            

                    FileOutputStream out = new FileOutputStream(new File("C:/Users/688714/Desktop/Results.xlsx"));
                    workbook.write(out);
                    out.close();
                    System.out.println("Results.xlsx written successfully on disk.");
                
            }
                catch (NullPointerException e)
                {
                    e.printStackTrace();
                }
            

        }
        
        

}
    
   

	
   

