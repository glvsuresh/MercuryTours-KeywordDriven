package com.training.mercurytours.main;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.training.mercurytours.keywords.MTAppLibrary;
import com.training.mercurytours.utils.CommonLib;
import com.training.mercurytours.utils.Constants;

public class MainDriverScript {
	
	public static Workbook oWorkBook;
	public static Sheet oWorkTCSheet,oWorkTSSheet;
	public static String sTCID,sRunFlag,sTSID,sTSShtTCID,sKeyWord,sData,sRunStatus,sOutPutFile;
	public static Object[] oData;
	public static CellStyle oStyle;
	
	public static void main(String[] args) throws IOException {
		excelDriver();
	}
	public static void excelDriver() throws IOException
	{
		try
		{
			/*
			 *  Read the Total No.of Test Cases
			 *  Iterate Through each of Test Case
			 *  For Each Test Case , Read the Test Steps
			 *  For Each Test Step , Read the Keyword and Data
			 *  Execute the TestCases
			 *  Update the Test cases to Pass or Failed bases on the Execution
			 *  Save the Test Report. 
			 */
				FileInputStream oFile=new FileInputStream(Constants.XLPath);
				if(Constants.XLPath.endsWith(".xlsx"))
				{
					oWorkBook=new XSSFWorkbook(oFile);
				}
				else
				{
					oWorkBook=new HSSFWorkbook(oFile);
				}
				
				oWorkTCSheet=oWorkBook.getSheet(Constants.XLTCShtName);
				int iRowCnt=oWorkTCSheet.getPhysicalNumberOfRows();
				int iColCnt=oWorkTCSheet.getRow(0).getPhysicalNumberOfCells();
				for(int iRow=1;iRow<iRowCnt;iRow++)
				{
					sTCID=oWorkTCSheet.getRow(iRow).getCell(Constants.iTC_IDCol).getStringCellValue();
					sRunFlag=oWorkTCSheet.getRow(iRow).getCell(Constants.iRunFlagCol).getStringCellValue();
					if(sRunFlag.equalsIgnoreCase("yes"))
					{
						oWorkTSSheet=oWorkBook.getSheet(Constants.XLTSShtName);
						{
							int iTSRowCnt=oWorkTSSheet.getPhysicalNumberOfRows();
							int iTSColCnt=oWorkTSSheet.getRow(0).getPhysicalNumberOfCells();
							for(int iTSRow=1;iTSRow<iTSRowCnt;iTSRow++)
							{
								sTSShtTCID=oWorkTSSheet.getRow(iTSRow).getCell(Constants.iTC_IDCol).getStringCellValue();
								if(sTCID.equalsIgnoreCase(sTSShtTCID))
								{
									sTSID=oWorkTSSheet.getRow(iTSRow).getCell(Constants.iTS_IDCol).getStringCellValue();
									sKeyWord=oWorkTSSheet.getRow(iTSRow).getCell(Constants.iKeywordCol).getStringCellValue();
									sData=oWorkTSSheet.getRow(iTSRow).getCell(Constants.iDataCol).getStringCellValue();
									 if (!sData.isEmpty())
						             {
						                 oData = sData.split(",");
						             }
						             else
						             {
						            	 oData = new Object[] { };
						             }
									System.out.println("For the TestCase ID "+sTCID+"....Test Step ID is"+sTSID+"....Keyword Is..."+sKeyWord+"....Data is....."+sData);
									sRunStatus=executeTestCases(sKeyWord, oData);
									oStyle=oWorkBook.createCellStyle();
									if(sRunStatus.toLowerCase().contains("pass"))
									{
										oWorkTCSheet.getRow(iRow).getCell(4).setCellValue("Pass");
										oWorkTCSheet.getRow(iRow).getCell(4).setCellStyle(oStyle);
										oStyle.setFillForegroundColor(IndexedColors.GREEN.index);
					                	oStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
					        			
					        		}
									else
									{
										oWorkTCSheet.getRow(iRow).getCell(4).setCellValue("Fail");
										oWorkTCSheet.getRow(iRow).getCell(4).setCellStyle(oStyle);
										oStyle.setFillForegroundColor(IndexedColors.RED.index);
					                	oStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
					        		}
									//System.out.println("The Test case status is "+ExecuteTestcases(sKeyWord,oData));
								}
							}
						}
					}
					else
					{
						oStyle=oWorkBook.createCellStyle();
						System.out.println("User do not want to run the test case and runflag set as:"+sRunFlag);
						oWorkTCSheet.getRow(iRow).getCell(4).setCellValue("Skipped");
						oWorkTCSheet.getRow(iRow).getCell(4).setCellStyle(oStyle);
						oStyle.setFillForegroundColor(IndexedColors.ORANGE.index);
						oStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
												
					}
				}
				sOutPutFile=System.getProperty("user.dir")+"\\Reports\\KeyWordDrivenFrameworkReport"+CommonLib.getDateTimeStamp()+".xlsx";
				System.out.println(sOutPutFile);
	        	FileOutputStream oExcelReport = new FileOutputStream(sOutPutFile);
	        	oWorkBook.write(oExcelReport);
	        	oExcelReport.close();
	        	oWorkBook=null;
			}
			catch (FileNotFoundException e) 
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	
	public static String executeTestCases(String sKeyWord,Object[] oData)
	{
		String sRtnVal=null;
		try
		{
			MTAppLibrary oAppLib=new MTAppLibrary();
			Method[] oMethods=oAppLib.getClass().getMethods();
			for(int i=0;i<oMethods.length;i++)
			{
				if(oMethods[i].getName().equalsIgnoreCase(sKeyWord))
				{
					try
					{
					sRtnVal=oMethods[i].invoke(oAppLib, oData).toString();
					break;
					}
					catch(Exception e)
					{
						System.out.println("There is an issue while calling the Keyword"+sKeyWord+"=="+oData.toString());
					}
				}
				
			}
			return sRtnVal;
			
		}
		catch(Throwable t)
		{
			System.out.println("Error occurred while calling the keyword due to "+t.getMessage());
			return "Failed";
		}
	}
	

}
