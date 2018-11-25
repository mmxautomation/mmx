package com.org.mmx;



import org.testng.annotations.Test;

public class Check_test {

	@Test(dataProviderClass=dataProviders.ExcelReader.class, dataProvider="dataForLogin")
	public  void test(String Name, String Number)
	{
		//int Numbers = Integer.valueOf(Number);
		System.out.println("Your Name is "+ Name + " & number is " + Number);

	}
}

