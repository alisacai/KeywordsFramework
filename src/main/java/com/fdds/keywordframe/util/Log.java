package com.fdds.keywordframe.util;


import org.apache.log4j.Logger;

public class Log {
	private static Logger log = Logger.getLogger(Log.class.getName());

	//定义测试用例开始执行的打印方法，在日志中打印测试用例开始执行的信息
	public static void startTestCase(String testCaseName){
		log.info("-----------------------\"" + testCaseName + "\"开始执行---------------");
	}

	//定义测试用例执行完毕后的打印方法，在日志中打印测试用例执行完毕的信息
	public static void endTestCase(String testCaseName){
		log.info("-----------------------\"" + testCaseName + "\"测试执行结束---------------");
	}

	//定义打印info级别日志的方法
	public static void info(String message){
		log.info(message);
	}

	//定义打印error级别日志的方法
	public static void error(String message){
		log.error(message);
	}

	//定义打印debug级别日志的方法
	public static void debug(String message){
		log.debug(message);
	}
}
