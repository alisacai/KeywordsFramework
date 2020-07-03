package com.fdds.keywordframe.configuration;

import java.util.List;

import com.fdds.keywordframe.testScript.TestSuiteByExcel;
import com.fdds.keywordframe.util.KeyBoardUtil;
import com.fdds.keywordframe.util.Log;
import com.fdds.keywordframe.util.ObjectMap;
import com.fdds.keywordframe.util.WaitUtil;


import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.Assert;


public class KeyWordsAction {
	public static WebDriver driver;
	public static ObjectMap objectMap = new ObjectMap(Constant.Path_ConfigurationFile);
	static {
		DOMConfigurator.configure("log4j.xml");
	}
	//此方法的名称对应Excel文件“关键字”列中的open_browser关键字
	//Excel文件中“操作值”列中的内容用于指定测试用例用何种浏览器运行测试用例
	//实际传入的参数，仅为了通过反射机制统一地使用两个函数参数来调用此函数
	public static void open_browser(String string,String browserName){
		if(browserName.equals("ie")){
			System.setProperty("webdriver.ie.driver","E:\\Selenium\\BrowserDriver\\IEDriverServer64.exe");
			driver = new InternetExplorerDriver();
			Log.info("IE浏览器实例已经声明");
		}else if(browserName.equals("Chrome")){
			System.setProperty("webdriver.chrome.driver","E:\\Selenium\\BrowserDriver\\chromedriver.exe");
			driver = new ChromeDriver();
			Log.info("Chrome浏览器实例已经声明");
		}else if(browserName.equals("Firefox")){
			driver = new FirefoxDriver();
			Log.info("Firefox浏览器实例已经声明");
		}
	}
	
	public static void navigate(String string, String baseUrl){
		driver.navigate().to(baseUrl);
		Log.info("浏览器访问网址"+baseUrl);
	}	
	
	public static void input(String locatorExpression,String inputString) throws Exception{
		try{
			driver.findElement(objectMap.getLocator(locatorExpression)).clear();
			Log.info("清除"+locatorExpression+"输入框中的所有内容");
			driver.findElement(objectMap.getLocator(locatorExpression)).sendKeys(inputString);
			Log.info("在"+locatorExpression+"输入框中输入"+inputString);
		}catch(Exception e){
			TestSuiteByExcel.testResult = false;
			Log.info("在"+locatorExpression+"输入框中输入"+inputString+"时出现异常，具体异常信息："+e.getMessage());
			e.printStackTrace();
		}	
	}

	public static void click(String locatorExpression,String string){
		try{
			driver.findElement(objectMap.getLocator(locatorExpression)).click();
			Log.info("单击"+locatorExpression+"页面元素成功");
		}catch(Exception e){
			TestSuiteByExcel.testResult = false;
			Log.info("单击"+locatorExpression+"页面元素失败，具体异常信息："+e.getMessage());
			e.printStackTrace();
		}		
	}
	
	public static void WaitFor_Element(String locatorExpression,String inputString){
		try{
			WaitUtil.waitWebElement(driver,objectMap.getLocator(locatorExpression));
			Log.info("显示等待元素出现成功，元素是"+locatorExpression);
		}catch(Exception e){
			TestSuiteByExcel.testResult = false;
			Log.info("显示等待元素出现异常，具体异常信息:"+e.getMessage());
			e.printStackTrace();
		}
	}
	
	public static void press_Tab(String string1,String string2){
		try{
			Thread.sleep(2000);
			KeyBoardUtil.pressTabKey();
			Log.info("按Tab键成功");
		}catch(Exception e){
			TestSuiteByExcel.testResult = false;
			Log.info("按Tab键时出现异常,具体异常信息："+e.getMessage());
			e.printStackTrace();
		}
	}
	
	public static void pasteString(String string,String pasteContent){
		try{			
			KeyBoardUtil.SetAndCtrlVClipboardData(pasteContent);
			Log.info("成功粘贴邮件正文："+ pasteContent);
		}catch(Exception e){
			TestSuiteByExcel.testResult = false;
			Log.info("粘贴邮件正文信息出现异常,具体异常信息："+e.getMessage());
			e.printStackTrace();
		}
	}
	
	public static void press_Enter(String string1,String string2){
		try{
			KeyBoardUtil.pressEnterKey();
			Log.info("按Enter键成功");
		}catch(Exception e){
			TestSuiteByExcel.testResult = false;
			Log.info("按Enter键时出现异常,具体异常信息："+e.getMessage());
			e.printStackTrace();
		}
	}
	
	public static void sleep(String string1,String sleepTime){
		try{
			WaitUtil.sleep(Integer.parseInt(sleepTime));
			Log.info("休眠 "+Integer.parseInt(sleepTime)/1000+"秒成功");
		}catch(Exception e){
			TestSuiteByExcel.testResult = false;
			Log.info("线程休眠时出现异常,具体异常信息："+e.getMessage());
			e.printStackTrace();
		}
	}
	
	public static void click_sendMailButton(String locatorExpression,String string){
		try{
			List<WebElement> buttons = driver.findElements(objectMap.getLocator(locatorExpression));
			buttons.get(0).click();
			Log.info("单击发送邮件按钮成功");
		}catch(Exception e){
			TestSuiteByExcel.testResult = false;
			Log.info("单击发送邮件按钮出现异常,具体异常信息："+e.getMessage());
			e.printStackTrace();
		}
	}

	public static void Assert_String(String string,String assertString){
		try{
			Assert.assertTrue(driver.getPageSource().contains(assertString));
			Log.info("成功断言关键字“"+assertString+"”");
		}catch(Exception e){
			TestSuiteByExcel.testResult = false;
			Log.info("断言关键字出现异常,具体异常信息："+e.getMessage());
			e.printStackTrace();
		}
	}
	
	public static void close_browser(String string,String string2){
		try{
			System.out.println("浏览器关闭函数被执行");
			driver.quit();
			Log.info("关闭浏览器窗口");
		}catch(Exception e){
			TestSuiteByExcel.testResult = false;
			Log.info("关闭浏览器窗口出现异常,具体异常信息："+e.getMessage());
			e.printStackTrace();
		}
	}
}