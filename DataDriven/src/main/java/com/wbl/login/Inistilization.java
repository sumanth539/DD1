package com.wbl.login;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Inistilization {

	public static WebDriver driver;
    public static Properties prop;

    public  Inistilization() {
    	
    	try {
			prop = new Properties();
			FileInputStream ip = new FileInputStream(System.getProperty("user.dir")+"/DataDriven/config.properties");

			prop.load(ip);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch(IOException e) {
			e.printStackTrace();
		}
	}
    public static void BrowserInistilization() {
		
	System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"/DataDriven/driver/chromedriver.exe");	
	driver = new ChromeDriver();
    prop.getProperty("pageurl");

}
    
    
}
