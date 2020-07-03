package com.fdds.keywordframe.util;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;

public class KeyBoardUtil {
	//按Enter键的封装方法
	public static void pressEnterKey(){
		Robot robot = null;
		try {
			robot = new Robot();
		} catch (AWTException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//调用KeyPress方法来实现按下Enter键
		robot.keyPress(KeyEvent.VK_ENTER);
		//调用keyRelease方法来实现按下Enter键
		robot.keyRelease(KeyEvent.VK_ENTER);		
	}
	//按Tab键的封装方法
	public static void pressTabKey(){
		Robot robot = null;
		try {
			robot = new Robot();
		} catch (AWTException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//调用KeyPress方法来实现按下Tab键
		robot.keyPress(KeyEvent.VK_TAB);
		//调用keyRelease方法来实现按下Tab键
		robot.keyRelease(KeyEvent.VK_TAB);
	}
	
	/*
	 * 封装好的在富文本框中指定字符串内容的粘贴函数
	 * 将指定字符串设为剪切板的内容，然后执行粘贴操作
	 * 将页面焦点切换到输入框后，调用此函数可以将指定字符串粘贴到输入框中
	 */
	public static void SetAndCtrlVClipboardData(String string){
		Robot robot = null;
		try {
			robot = new Robot();
		} catch (AWTException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		StringSelection ss = new StringSelection(string);
		//使用ToolKit对象的setContent方法将字符串放到剪切板中
		Toolkit.getDefaultToolkit().getSystemClipboard().setContents(ss,null);
		//按下和释放Ctrl+V组合键
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_V);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_V);
	}
}
