package demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;

import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import prac.ExcelFile;


public class BalanceComputation1
{
   
   	
	
  XSSFWorkbook workbook = new XSSFWorkbook();
 

   XSSFSheet sheet = workbook.createSheet("computeExcelBalanceValue");
    
   
    

       
  
  CustomerVO computeBalance(String id) throws Exception 
  {
	  CustomerVO cvo1=new CustomerVO();
	  
	  
	  
	  @SuppressWarnings("rawtypes")
	
		//BalanceComputation bc=new BalanceComputation();
	String url = "jdbc:ucanaccess://C:/Users/688714/workspace/Balance/WebContent/HTMLAccess.mdb";    
    Connection conn = null;
    conn = DriverManager.getConnection(url,"","");
    PreparedStatement st ;
    
    
  //  System.out.println("****BALANCECOMPUTATION1******\n\n");   
    
    
    String query = "SELECT * FROM [CustomerData] WHERE ID="+id+" ";
  //  System.out.println("query " + query);
    
    
    try 
    {
    	st = conn.prepareStatement(query);
        
        ResultSet  rs = st.executeQuery();
        
    	    
    	if(rs.next())
    	{
    		
    	     rs=st.executeQuery();
    			
    		  while(rs.next())
    		  {
    		  String id1=rs.getString("ID");
    		  String FirstName=rs.getString("FIRSTNAME");
    		  String LastName=rs.getString("LASTNAME");
    		  String CurrentYr=rs.getString("CURRENTYR");
    		  double LastYrBalance1=rs.getDouble("LASTYRBALANCE");
    		  double CurrentYrDeposit=rs.getDouble("CurrentYrDeposit");
    		  double CurrentYrWithdrawal=rs.getDouble("CurrentYrWithdrawal");
              cvo1.setCustId(id1);
              cvo1.setFirstName(FirstName);
              cvo1.setLastName(LastName);
    	      cvo1.setCurrentYr(CurrentYr);
              
    	      cvo1.setLastyrDeposit(LastYrBalance1);
    	       
    	      cvo1.setLastyrDeposit(CurrentYrDeposit);
    	      
              cvo1.setLastyrDeposit(CurrentYrWithdrawal);
              double computedBalance=LastYrBalance1+CurrentYrDeposit-CurrentYrWithdrawal;
    	      cvo1.setComputedBalance(computedBalance);
              
              
              
        
      //  System.out.println("In BalanceComputation1.java");
    	      if(computedBalance>0.0)
    	      {
    	    	 
    	           System.out.println("***BALANCE COMPUTATION FROM DATABASE***");
                   System.out.println("Details of CUSTOMER WITH id "+id1);
                   System.out.println("Cust Id: "+id1);
                   System.out.println("Customer First Name: "+FirstName);
                   System.out.println("Customer Last Name: "+LastName);
                   System.out.println("End of Yr balance of id "+ id +" is : Rs. "+computedBalance);
                   
                   
           	      
           	        
           	          
           	                	        
           	          
           	          // Row row2=sheet.createRow(1);
           	                  	         
                      
           	        //  System.out.println(computedBalance);
           	       //   Cell cell=row2.createCell(0);
           	          
           	        	  
           	          
           	        //  cell.setCellValue(computedBalance);
           	         
           	         
           	         
           	         
           	         
           	   
              
    	      }
    	      else
    	      {
    	    	  System.out.println("Invalid computed balance");
    	      }
    		  }
    		  
    		  
    	}
    	
    	else	
    	{
    			System.out.println("In computeBalance method. No record found in database with custid " +id);
    			return null;
    	}
    
    	    
    	 
    	
   	st.close();
   	conn.close();
  }
    catch(NumberFormatException e)
    {
    	System.out.println(e);
    }
    
    
    
    
    
    
    return cvo1;
  
    
  }

  






  
 
  

	 
 }
  
