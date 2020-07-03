package com.fdds.keywordframe.configuration;

public class Constant {

	//测试数据相关常量设定
	public static final String Path_ExcelFile = "src\\main\\java\\com\\fdds\\keywordframe\\data\\KeyWordsActionTestCases.xlsx";
	public static final String Path_ConfigurationFile = "ObjectMap.properties";
	
	//测试数据Sheet中的列号常量设定
	//第一列用0表示，测试用例序号列
	public static final int Col_TestCaseID = 0;
	//第四列用3表示，关键字列
	public static final int Col_KeyWordAction = 3;
	//第五列用4表示，操作元素的定位表达方式
	public static final int Col_LocatorExpression = 4;
	//第六列用5表示，操作值列
	public static final int Col_ActionValue = 5;
	//第7列用6表示，测试结果列
	public static final int Col_TestStepTestResult = 6;
	
	//测试用例集合Sheet中的列号常量设定
	public static final int Col_RunFlag = 2;
	public static final int Col_TestSuiteTestResult = 3;
	
	//测试用例Sheet名称的常量设定
	public static final String Sheet_TestSteps = "SendEmail";
	//测试用例集Sheet名称的常量设定
	public static final String Sheet_TestSuite = "TestSuite";
}
