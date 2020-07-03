package com.fdds.keywordframe.testScript;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import com.fdds.keywordframe.configuration.Constant;
import com.fdds.keywordframe.configuration.KeyWordsAction;
import com.fdds.keywordframe.util.ExcelUtil;
import com.fdds.keywordframe.util.Log;
import org.apache.log4j.xml.DOMConfigurator;
import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class TestSuiteByExcel {
	public static KeyWordsAction keyWordsAction;
	public static boolean testResult = true;
	public static int testStep;
	public static int lastStep;
	
	
	@BeforeClass
	public void BeforeClass(){
		DOMConfigurator.configure("log4j.xml");
	}
	
	@Test
	public void testTestSuite() throws Exception{
		
		keyWordsAction = new KeyWordsAction();
		
		String excelFilePath = Constant.Path_ExcelFile;
		ExcelUtil.setExcelFile(excelFilePath);
		int rowCount = ExcelUtil.getRowCount(Constant.Sheet_TestSuite);
		
		for(int testcaseNo = 1;testcaseNo <= rowCount; testcaseNo++){
			String testcaseID = ExcelUtil.getCellData(Constant.Sheet_TestSuite, testcaseNo, Constant.Col_TestCaseID);
			testStep = ExcelUtil.getFirstRowContainsTestCaseID(Constant.Sheet_TestSteps, testcaseID, Constant.Col_TestCaseID);
			lastStep = ExcelUtil.getLastRowContainsTestCaseID(Constant.Sheet_TestSteps, testcaseID, Constant.Col_TestCaseID);
			for(;testStep<=lastStep;testStep++){
				String keyword = ExcelUtil.getCellData(Constant.Sheet_TestSteps, testStep, Constant.Col_KeyWordAction);
				String keyvalue = ExcelUtil.getCellData(Constant.Sheet_TestSteps, testStep, Constant.Col_ActionValue);
				String locatorExpression = ExcelUtil.getCellData(Constant.Sheet_TestSteps, testStep, Constant.Col_LocatorExpression);
				Method method = KeyWordsAction.class.getMethod(keyword,String.class,String.class);
				execute_Action(method,locatorExpression,keyvalue);
				if(!testResult){
					ExcelUtil.setCellData(Constant.Sheet_TestSuite, testcaseNo, Constant.Col_TestSuiteTestResult, "测试用例执行失败");
					Log.endTestCase(testcaseID);
					break;
				}				
				if(testResult){
					ExcelUtil.setCellData(Constant.Sheet_TestSuite, testcaseNo, Constant.Col_TestSuiteTestResult, "测试用例执行成功");
				}
			}			
		}
	}
	
	public static void execute_Action(Method method,Object ... args) throws Exception, IllegalArgumentException, InvocationTargetException{
		try{
			method.invoke(keyWordsAction,args);
			if(testResult ){
				ExcelUtil.setCellData(Constant.Sheet_TestSteps, testStep, Constant.Col_TestStepTestResult, "测试步骤执行成功");
			}else{
				ExcelUtil.setCellData(Constant.Sheet_TestSteps, testStep, Constant.Col_TestStepTestResult, "测试步骤执行失败");
				keyWordsAction.close_browser("", "");
			}		
		}catch(Exception e){
			Assert.fail("执行出现异常，测试用例执行失败！");
		}
		
		
	}
	
}
