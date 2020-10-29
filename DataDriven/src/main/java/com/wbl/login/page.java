package com.wbl.login;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class page extends ExclUtil {
	public static WebDriver driver;
	

    @FindBy(xpath="//input[@type='number']") private WebElement  number;
    @FindBy(xpath="//button[@class='btn btn-primary mb-5 float-right mr-5 mt-2']")
	private WebElement login;

public  page(WebDriver driver) {
	this.driver = driver;
	PageFactory.initElements(driver, this);
}

public void send(String string) {
	number.sendKeys();
}

public void loginButton() {
	login.click();
}

public void url() {
	String actual="https://ping.telebu.com/chat";
}

}
