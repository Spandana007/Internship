package demo;

import java.util.Scanner;
import demo.BalanceComputation1;
public class CommandLineExecutor {
	
	
	//BalanceComputation bc=new BalanceComputation();
	public static void main(String args[]) throws Exception
	{
		
		//BalanceComputation bc=new BalanceComputation();
	    
		@SuppressWarnings("resource")
		Scanner sc=new Scanner(System.in);
		System.out.println("ACCESS to DATABASE");
		System.out.println("Enter the cust id:");
		String id=sc.next();
		
		/*URL balance = new URL("http://localhost:8085/Balance2/Main");
	    URLConnection b = balance.openConnection();
	    BufferedReader in = new BufferedReader(new InputStreamReader(b.getInputStream()));
	    String inputLine;

	    while ((inputLine = in.readLine()) != null) 
	        System.out.println(inputLine);
	    in.close();
	    */
		BalanceComputation1 m=new BalanceComputation1();
		m.computeBalance(id);

	}
	
}

