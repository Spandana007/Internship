package demo;

import javax.swing.JFileChooser;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ValueCompare 
{
	
    

	static ArrayList<CustomerVO> IDList = new ArrayList<CustomerVO>();
    static ArrayList<CustomerVO> BalFromDBList=new ArrayList<CustomerVO>();
    static ArrayList<CustomerVO> BalFromExcelList=new ArrayList<CustomerVO>();
    static CustomerVO foundInDB=new CustomerVO();
    static CustomerVO foundInEx=new CustomerVO();

	



	
	public static void main(String args[]) throws Exception
    {
      
      try
      {
          FileInputStream file = new FileInputStream(new File("C:/Users/688714/Desktop/CustomerRecords.xls"));
         ArrayList<CustomerVO> custlist= new ArrayList<CustomerVO>();
         CustomerVO cvo= new CustomerVO();
         
          //Create Workbook instance holding reference to .xlsx file
          HSSFWorkbook workbook = new HSSFWorkbook(file);

          //Get first/desired sheet from the workbook
          HSSFSheet sheet = workbook.getSheetAt(0);

          

          //Iterate through each rows one by one
          Iterator<Row> rows = sheet.rowIterator();
          String id = null;
          
           
          while (rows.hasNext()) 
          {
              HSSFRow row = (HSSFRow) rows.next();
              if(row.getRowNum()==0)
              {
            	  continue;
              }
              //For each row, iterate through all the columns
              Iterator<Cell> cellIterator = row.cellIterator();

              while (cellIterator.hasNext()) 
              {
                  Cell cell = cellIterator.next();
                  //Check the cell type and format accordingly              
                  
                      
                         if(cell.getCellType()==HSSFCell.CELL_TYPE_NUMERIC)
                         {          
                        	      if(cell.getColumnIndex()==0)
                        	      {
                                   
                                	   DataFormatter fmt1 = new DataFormatter();
                                       
                                       id = fmt1.formatCellValue(cell);
                                       
                                       cvo= new CustomerVO();
                                	   cvo.setCustId(id);
                                	   
                                	  // custlist.add(cvo);
                                	   
                                	     BalanceComputation1 bc=new BalanceComputation1();
                                	   
                                	     foundInDB=bc.computeBalance(id);
                                	     if(foundInDB!=null)
                                	     {
                                	       
                                	       BalFromDBList.add(foundInDB);
                                	       IDList.add(foundInDB);
                                	                                       	       
                                	     }
                                	    
                                	     
                                	     
                                	     
                                	   
                        	      }
                        	      else if(cell.getColumnIndex()==3)
                        	      {
                        	    	  DataFormatter fmt1 = new DataFormatter();
                                      
                                      String currentyr = fmt1.formatCellValue(cell);
                                      cvo.setCurrentYr(currentyr);
                                      
                                      
                        	      }
                        	      else if(cell.getColumnIndex()==4)
                        	      {
                        	    	  
                        	    	  cvo.setLastyrDeposit(cell.getNumericCellValue());
                        	    	  
                        	      }
                        	      else if(cell.getColumnIndex()==5)
                        	      {
                        	    	  cvo.setCurrentYrDeposit(cell.getNumericCellValue());
                        	    	  
                        	      }
                        	      else if(cell.getColumnIndex()==6)
                        	      {
                        	    	  cvo.setCurrentYrWithdrawal(cell.getNumericCellValue());
                        	    	  
                        	      }
                        	    
                                      
                         }
                         
                                       
           
                  
                        
                              
                          
                       }
                       
                       if(foundInDB!=null)
                       {
                           foundInEx=computeExcelBalanceValue(cvo); 
                       
                       
                          
                            BalFromExcelList.add(foundInEx);
                            
                          
                         
                       }
                       
                      else
                      {
               	        System.out.println("In computeExcelBalance method. No record found in database with custid!! "+id);
               	  
                      }
                    
                    
              
                   
              
             
          

              System.out.println("");
             
          }
         // ArrayList<CustomerVO> custList=new ArrayList<CustomerVO>();
        //  CustomerVO cvo1=new CustomerVO();
         // custList.add(cvo1);
          
          
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            
          
          
           ComputeBalanceResults cb=new ComputeBalanceResults();
          
          cb.computeBalanceToExcelFile(IDList,BalFromDBList,BalFromExcelList);
          
          file.close();
         
         
          
         
      } 
      catch (NullPointerException e) 
      {
          e.printStackTrace();
          
      }
      
      
      
    
    }
 
  

@SuppressWarnings("unchecked")
static
CustomerVO computeExcelBalanceValue(CustomerVO cvo1) throws IOException
{
   
	CustomerVO cvo2=new CustomerVO();

   try{

	   FileOutputStream out =  new FileOutputStream(new File("C:/Users/688714/Desktop/formulaDemo2.xlsx"));

       

      

       
    @SuppressWarnings("resource")
	XSSFWorkbook workbook = new XSSFWorkbook();
    

    XSSFSheet sheet = workbook.createSheet("computeExcelBalanceValue");
    ArrayList<CustomerVO> custlist1= new ArrayList<CustomerVO>();
    custlist1.add(cvo1);

     //From your method
    
    for (CustomerVO cvo : custlist1) 
    {  
      Row row1 = sheet.createRow(0);

      row1.createCell(0).setCellValue("CurrentYr");

      row1.createCell(1).setCellValue(cvo.getCurrentYr());

         

    Row row2 = sheet.createRow(1);

    row2.createCell(0).setCellValue("LastYrBalance");

    row2.createCell(1).setCellValue(cvo.getLastyrDeposit());

    Row row3=sheet.createRow(2);
    row3.createCell(0).setCellValue("CurrentYrDeposit");
    row3.createCell(1).setCellValue(cvo.getCurrentYrDeposit());
    
    Row row4=sheet.createRow(3);
    row4.createCell(0).setCellValue("CurrentYrWithdrawal");
    row4.createCell(1).setCellValue(cvo.getCurrentYrWithdrawal());
    
    Row row5=sheet.createRow(4);
    row5.createCell(0).setCellValue("Total Income Current Yr");
    row5.createCell(1).setCellFormula("B2+B3");
    
    Row row6=sheet.createRow(5);
    row6.createCell(0).setCellValue("Total Withdrawal");
    row6.createCell(1).setCellFormula("B4");
    
    Row row7=sheet.createRow(6);
    row7.createCell(0).setCellValue("Remaining Balance");
    row7.createCell(1).setCellFormula("B5-B6");
    
    Row row8=sheet.createRow(7);
    row8.createCell(0).setCellValue("End of Yr Balance");
    row8.createCell(1).setCellFormula("B7");
    
    workbook.write(out);
    
    
    FileInputStream in = new FileInputStream(new File("C:/Users/688714/Desktop/formulaDemo2.xlsx"));
    
    System.out.println();
    System.out.println("***BALANCE COMPUTATION FROM SPREADSHEET***");


    //Create Workbook instance holding reference to .xlsx file

    @SuppressWarnings("resource")
	XSSFWorkbook workbook1 = new XSSFWorkbook(in);



    FormulaEvaluator evaluator = workbook1.getCreationHelper().createFormulaEvaluator();

     

    //Get first/desired sheet from the workbook

    XSSFSheet sheet1 = workbook1.getSheetAt(0);



    //Iterate through each rows one by one

    Iterator<Row> rowIterator = sheet1.iterator();

    while (rowIterator.hasNext()) 

    {

        Row row = rowIterator.next();

        //For each row, iterate through all the columns

        Iterator<Cell> cellIterator = row.cellIterator();

         

        while (cellIterator.hasNext()) 

        {

            Cell cell = cellIterator.next();
            

            //Check the cell type after eveluating formulae

            //If it is formula cell, it will be evaluated otherwise no change will happen

            switch (evaluator.evaluateInCell(cell).getCellType()) 

            {
               
                case Cell.CELL_TYPE_NUMERIC:
                	
                	if(cell.getColumnIndex()==1 && cell.getRowIndex()==7)
                	{
                	 
                	 
                	// DataFormatter fmt1 = new DataFormatter();
                    
                    // String valueint = fmt1.formatCellValue(cell);
                     

                     System.out.println(" : Rs. "+cell.getNumericCellValue());
                     double balance=cell.getNumericCellValue();
                     cvo2.setComputedBalance(balance);
                     
                     
                     
                	}
                	else
                	{
                		continue;
                	}

                    break;

                case Cell.CELL_TYPE_STRING:
                	if(cell.getColumnIndex()==0 && cell.getRowIndex()==7)
                	{

                    System.out.print(cell.getStringCellValue());
                	}
                	else
                	{
                		continue;
                		
                	}

                    break;

                case Cell.CELL_TYPE_FORMULA:

                    //Not again

                    break;

            }
            

        }

        

    }
    
    
    
  }
   
    out.close();
    workbook.close();
    

 }

    

      

    catch (FileNotFoundException e) 
    {

        e.printStackTrace();

    } catch (IOException e) 
    {

        e.printStackTrace();

    }
return cvo2;
    

   
   }
}


  

