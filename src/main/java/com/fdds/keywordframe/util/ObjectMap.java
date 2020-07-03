package com.fdds.keywordframe.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.By;

public class ObjectMap {

	Properties properties;

	public ObjectMap(String propFile){
		properties = new Properties();
		try {
			FileInputStream input = new FileInputStream(propFile);
			properties.load(input);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			System.out.println("读取文件出错");
			e.printStackTrace();
		}
	}

	public static By getLocator(String propName) throws Exception{

		String locatorType = propName.split(".")[0];
		String locatorValue = propName.split(".")[1];
		if(locatorType.toLowerCase().equals("id")){
			return By.id(locatorValue);
		}else if(locatorType.toLowerCase().equals("class")){
			return By.name(locatorValue);
		}else if(locatorType.toLowerCase().equals("class")||locatorType.toLowerCase().equals("classname")){
			return By.className(locatorValue);
		}else if(locatorType.toLowerCase().equals("tagname")){
			return By.tagName(locatorValue);
		}else if(locatorType.toLowerCase().equals("link")||locatorType.toLowerCase().equals("linktext")){
			return By.linkText(locatorValue);
		}else if(locatorType.toLowerCase().equals("partiallinktext")){
			return By.partialLinkText(locatorValue);
		}else if(locatorType.toLowerCase().equals("cssSelector")||locatorType.toLowerCase().equals("css")){
			return By.cssSelector(locatorValue);
		}else if(locatorType.toLowerCase().equals("xpath")){
			return By.xpath(locatorValue);
		}else{
			throw new Exception("输入的locator type未在程序中被定义：" +locatorType);
		}
	}
}
