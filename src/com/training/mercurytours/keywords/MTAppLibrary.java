package com.training.mercurytours.keywords;

import org.openqa.selenium.By;

import com.training.mercurytours.utils.CommonLib;

public class MTAppLibrary extends CommonLib{
	
	By edtUserName=By.name("userName");
	By edtPassword=By.name("password");
	By btnLogin=By.name("login");
	By lnkIternary=By.linkText("ITINERARY");
	By lstPassngr=By.name("passCount");
	By lstDept=By.name("fromPort");
	By lstTrvlDate=By.name("fromDay");
	By lstArrival=By.name("toPort");
	By lstRtnDate=By.name("toDay");
	By radioClass=By.xpath("//input[@value='Business']");
	By lstAirlines=By.name("airline");
	By elemSelectFlight=By.xpath("//img[contains(@src,'selectflight')]");
	By btnFindFlights=By.name("findFlights");
	By NoOfTrvlFlights=By.xpath("//input[@name='outFlight']");
	By NoOfRtnFlights=By.xpath("//input[@name='inFlight']");
	By btnReserveFlght=By.name("reserveFlights");
	By edtUserFirstName=By.name("passFirst0");
	By edtUserLastName=By.name("passLast0");
	By edtCreditNo=By.name("creditnumber");
	By btnBuyFlights=By.name("buyFlights");
	By elemFlightConfirm=By.xpath("//img[contains(@src,'mast_confirmation')]");
	By btnLogout=By.xpath("//img[contains(@src,'Logout')]");
	
	public String loginToMT(String sUserName,String sPassword)
	{
		try
		{
			enterText(edtUserName, sUserName);
			enterText(edtPassword, sPassword);
			clickElem(btnLogin);
			if(isElemDisplayed(getElem(lnkIternary)))
			{
				System.out.println("Logged into Mercury Tours Successfully");
				return "Pass";
			}
			else
			{
				System.out.println("Failed to Login to Mercury Tours");
				takeScreenshot("LoginFailed.png");
				return "Login to Mercurytours got failed";
			}
			
		}
		catch(Throwable t)
		{
			System.out.println(t.getMessage());
			return "Login to Mercurytours got failed";
		}
	}
	
	
	public String findFlights(String sPasscount,String sDeparture, String sTravelDt,String sArrival,String sRetnDate,String sAirLinesName)
	{
		try
		{
			selectValFromListbox(getElem(lstPassngr), sPasscount);
			selectValFromListbox(getElem(lstDept), sDeparture);
			selectValFromListbox(getElem(lstTrvlDate), sTravelDt);
			selectValFromListbox(getElem(lstArrival), sArrival);
			selectValFromListbox(getElem(lstRtnDate), sRetnDate);
			clickElem(radioClass);
			selectValFromListbox(getElem(lstAirlines), sAirLinesName);
			clickElem(btnFindFlights);
			if(isElemDisplayed(getElem(elemSelectFlight)))
			{
				System.out.println("Searching flights is successful");
				return "Pass";
			}
			else
			{
				takeScreenshot("FlightSearchFailed");
				System.out.println("Flight search is failed");
				return "Failed to search for flights";
			}
			
		}
		catch(Throwable t)
		{
			System.out.println("Failed to search for flights");
			return "Failed to search for flights";
		
		}
	}
	
	
	public String selectFlightsForRoundTrip(String sDeptAirlines,String sRtnAirlines)
	{
		
		try
		{
			clickOnElemByusingName(getElems(NoOfTrvlFlights), sDeptAirlines);
			clickOnElemByusingName(getElems(NoOfRtnFlights), sRtnAirlines);
			clickElem(btnReserveFlght);
			if(isElemDisplayed(getElem(edtUserFirstName)))
			{
				System.out.println("Flight selection is successful");
				return "Pass";
			}
			else
			{
				takeScreenshot("Flight selection failed");
				System.out.println("Flight selection failed");
				return "Flight Selection is failed";
			}
		}
		catch(Throwable t)
		{
			takeScreenshot("Flight selection failed");
			System.out.println("Flight selection failed because of "+t.getMessage());
			return "Flight Selection is failed";
		}
	}
	
	public String enterPassengerDetails(String sPassFN, String sPassLN,String sCreditNo) {
		try {
			enterText(edtUserFirstName, sPassFN);
			enterText(edtUserLastName, sPassLN);
			enterText(edtCreditNo, sCreditNo);
			clickElem(btnBuyFlights);
			if (isElemDisplayed(getElem(elemFlightConfirm)))
			{
				System.out.println("Entered Passenger details successfully");
				return "Pass";

			} else {
				System.out.println("Error occurred while entering passenger details");
				takeScreenshot("FailedPasgr.png");
				return "Failed to enter Passenger details";
			}
		} catch (Throwable t) {
			System.out.println("Failed while entering passenger details");
			takeScreenshot("FailedPasgr.png");
			return "Failed to enter Passenger details";
		}
	}
	
	
	public String logOutFromMercuryTours()
	{
		try
		{
			clickElem(btnLogout);
			if(isElemDisplayed(getElem(edtUserName)))
			{
				System.out.println("Logging Out MT Successfully");
				return "Pass";
			}
			else
			{
				takeScreenshot("Logoutfailed.png");
				System.out.println("Logout failed");
				return "Failed to Logout from MT";
			}
		}
		catch(Throwable t)
		{
			takeScreenshot("Logoutfailed.png");
			System.out.println("Logout failed becuase of "+t.getMessage());
			return "Failed to logout from MT";
		}
	}


}
