package com.wbl.test;

import org.testng.annotations.BeforeClass;
import org.junit.Assert;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

import com.wbl.login.ExclUtil;
import com.wbl.login.Inistilization;
import com.wbl.login.page;

public class testLogin extends ExclUtil {

	page p;
	ExclUtil e;
	Inistilization i;
	@BeforeClass
	public void loginPage() {

		i = new Inistilization();
		BrowserInistilization();

	}
	@Test
	public void login() {
		p.send(e.getCellData(1,0));
		p.loginButton();
		String expected=driver.getCurrentUrl();
		String actual="https://ping.telebu.com/chat";
		Assert.assertEquals(expected,actual);

		
	}

}
