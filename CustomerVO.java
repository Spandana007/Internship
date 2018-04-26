package demo;

import java.util.Iterator;
import java.util.List;

public class CustomerVO 
{
  String CustId;
  String FirstName;
  String LastName;
  String CurrentYr;
  double LastyrDeposit;
  double CurrentYrDeposit;
  double CurrentYrWithdrawal;
  

double computedBalance;

public CustomerVO() {
	// TODO Auto-generated constructor stub
}
public String getCustId() {
	return CustId;
}
public void setCustId(String id) {
	CustId = id;
}
public String getFirstName() {
	return FirstName;
}
public void setFirstName(String firstName) {
	FirstName = firstName;
}
public String getLastName() {
	return LastName;
}

public void setLastName(String lastName) {
	LastName = lastName;
}
public String getCurrentYr() {
	return CurrentYr;
}
public void setCurrentYr(String currentYr) {
	CurrentYr = currentYr;
}
public double getLastyrDeposit() {
	return LastyrDeposit;
}
public void setLastyrDeposit(double lastyrDeposit) {
	LastyrDeposit = lastyrDeposit;
}
public double getCurrentYrDeposit() {
	return CurrentYrDeposit;
}
public void setCurrentYrDeposit(double currentyrDeposit) {
	CurrentYrDeposit = currentyrDeposit;
}
public double getCurrentYrWithdrawal() {
	return CurrentYrWithdrawal;
}
public void setCurrentYrWithdrawal(double withdrawal) {
	CurrentYrWithdrawal = withdrawal;
}
public double getComputedBalance() {
	
	
	return computedBalance;
	
}

public void setComputedBalance(double computedBalance) {
	this.computedBalance = computedBalance;
}

public static Object getCustomerVOList() {
	// TODO Auto-generated method stub
	return null;
}
public CustomerVO get(int rowIndex) {
	// TODO Auto-generated method stub
	return null;
}
                                                      

  
}

