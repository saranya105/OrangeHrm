package Apachepoi;
import java.io.FileInputStream;
import java.time.Duration;
import java.time.format.DateTimeFormatter;

import javax.swing.text.DateFormatter;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;



public class Orangehrm_myinfo {

	public static void main(String[] args) {
			
						 try {
						    	ChromeDriver driver =new ChromeDriver();
							    driver.manage().window().maximize();
							    driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
							    driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login");
							    Thread.sleep(2000);  
							   FileInputStream file = new FileInputStream("C:\\Users\\91939\\Desktop\\Orangehrm.xls");
							    XSSFWorkbook workbook = new XSSFWorkbook();
							    XSSFSheet sheet =  workbook.getSheet("Sheet1");
							    int rowcount = sheet.getLastRowNum();
							    int colcount = sheet.getRow(1))).getLastCellNum();
							    
							    for(int i=1; i<=rowcount; i++)
							    {
							    	String value1= new DataFormatter().formatCellValue(sheet.getRow(i))).getCell(1));
							    	String value2= new DataFormatter().formatCellValue ( sheet.getRow(i))).getCell(2));
							    	String value3= new DataFormatter().formatCellValue(sheet.getRow(i))).getCell(3));
							    	String value4= new DataFormatter().formatCellValue(sheet.getRow(i))).getCell(4));
							    
							    
							    	
							   //Validate wrong username & valid pwd
							    	driver.findElement(By.xpath("//input[@name='username']")).sendKeys("saranya");
							    	Thread.sleep(1000);
							    	driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
							    	Thread.sleep(1000);
							    	driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
							    	 Thread.sleep(2000);
							    	 
				               //Validate username & wrong pwd			    	 
							    	 driver.findElement(By.xpath("//input[@name='username']")).sendKeys("Admin");
								    	Thread.sleep(1000);
								    	driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("saranya@11");
								    	Thread.sleep(1000);
								    	driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
								    	 Thread.sleep(2000);
								     
				               //validate correc username & valid pwd   
								    	 driver.findElement(By.xpath("//input[@name='username']")).sendKeys("Admin");
									    	Thread.sleep(1000);
									    	driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
									    	Thread.sleep(1000);
									    	driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
									    	 Thread.sleep(2000);
								     
				                //personal details
									    	 driver.findElement(By.xpath("//input[@name='Employeename']")).sendKeys("DinhNguyen");
										    	Thread.sleep(1000);
										    	driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
										    	Thread.sleep(1000);
										    	driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
										    	 Thread.sleep(2000);
								     
				               //middle name
										    	 driver.findElement(By.xpath("//input[@name='middle name']")).sendKeys("Thi MyVan Canhan");
											    	Thread.sleep(1000);
											    	driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
											    	Thread.sleep(1000);	
											    	driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
											    	 Thread.sleep(2000);
								     
								//Nick name
											    	 driver.findElement(By.xpath("//input[@name='Nick name']")).sendKeys("CollingsSR");
												    	Thread.sleep(1000);
												    	driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");	
												    	Thread.sleep(1000);
												    	driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
												    	 Thread.sleep(2000);
									     
								     
							    //Employee  Id
												    	 driver.findElement(By.xpath("//input[@name='Employee Id']")).sendKeys("0024657");
													    	Thread.sleep(1000);
													    	driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													    	Thread.sleep(1000);
													    	driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													    	 Thread.sleep(2000);
										     

							    // Other Id
													    	 driver.findElement(By.xpath("//input[@name=' Other Id']")).sendKeys("100");
														    	Thread.sleep(1000);
														    	driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
														    	Thread.sleep(1000);
														    	driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
														    	 Thread.sleep(2000);
											     
							    	 
							    //Driver's License Number
														    	 driver.findElement(By.xpath("//input[@name='Driver's License Number']")).sendKeys("B0938874873");
															     Thread.sleep(1000);
															    driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");	
															    Thread.sleep(1000);
															    driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
															    Thread.sleep(2000);
												     
							    // License Expiry Date			    	 
															 driver.findElement(By.xpath("//input[@name='License  Expiry Date ']")).sendKeys("11-07-2023");
														     Thread.sleep(1000);
														     driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
															Thread.sleep(1000);
															driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
															Thread.sleep(2000);
								
							    //SSN Number
														 driver.findElement(By.xpath("//input[@name='SSN Number']")).sendKeys("SSN678948");
														Thread.sleep(1000);
														driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
														Thread.sleep(1000);
														driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
														 Thread.sleep(2000);
									
							    //SIN Number

														 driver.findElement(By.xpath("//input[@name='SIN Number']")).sendKeys("SIN85679");
														Thread.sleep(1000);
														driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
														Thread.sleep(1000);
														driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													    Thread.sleep(2000);
										
							    	
							    //Nationality
													 driver.findElement(By.xpath("//input[@name='Nationality']")).sendKeys("Romanian");
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);						    	
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);
											
								    	
							   //Martial Status
													 driver.findElement(By.xpath("//input[@name= 'Status']")).sendKeys("Single");
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);
												

							    //Date of Birth
													 driver.findElement(By.xpath("//input[@name='Date of Birth']")).sendKeys("07-11-2023");
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);
												
								//Gender
													 driver.findElement(By.xpath("//input[@name='Gender']")).sendKeys("male");
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);
												
						    	
							    //Military Service
													 driver.findElement(By.xpath("//input[@name=Experience']")).sendKeys("2+Years");
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);
												
							    	
							    	 //Smoker
													 driver.findElement(By.xpath("//input[@name=YES/NO']")).sendKeys("Not Smoker Required");
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);
												
							    	
							    	 //Blood Type
													 driver.findElement(By.xpath("//input[@name='Blood Type']" )).sendKeys("A+");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);
												
							    	
							    	 //Attachments
													 driver.findElement(By.xpath("//input[@name='Attachments']" )).sendKeys("Add Attachments->Select file added then Save");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);
							    	 //Contact details
							    	
													 driver.findElement(By.xpath("//input[@name='Contact details']" )).sendKeys("8523556842");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);
									//Street1
													 driver.findElement(By.xpath("//input[@name='Street1']" )).sendKeys("3576 Airport Way");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);
									//Street2

													 driver.findElement(By.xpath("//input[@name='Street2']" )).sendKeys("Unit9");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000); 
								 //City
													 driver.findElement(By.xpath("//input[@name='City']" )).sendKeys("FAIRBANKS");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000); 
								//State
													 driver.findElement(By.xpath("//input[@name='State']" )).sendKeys("AK");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000); 
							//ZIP/Postal Code
													 driver.findElement(By.xpath("//input[@name='ZIP/Postal Code']" )).sendKeys("99709");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000); 
							//Country
													 driver.findElement(By.xpath("//input[@name='Country']" )).sendKeys("FAIRBANKS North Star");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													 Thread.sleep(2000);

							// Home
													 driver.findElement(By.xpath("//input[@name=' Home']" )).sendKeys("USA");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													Thread.sleep(2000);
							//Mobile Number 
													driver.findElement(By.xpath("//input[@name='Mobile Number']" )).sendKeys("+1 8846755853");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													Thread.sleep(2000);
							//Work
													driver.findElement(By.xpath("//input[@name='Work']" )).sendKeys("DinhNguyen@gmail.com");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													Thread.sleep(2000);
							//Work Email
													driver.findElement(By.xpath("//input[@name='Work']" )).sendKeys("Nguyen2573@gmail.com");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													Thread.sleep(2000);
							//Other Email
													driver.findElement(By.xpath("//input[@name='Work']" )).sendKeys("Nguyenexcelr@gmail.com");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													Thread.sleep(2000);
						    //Name
													driver.findElement(By.xpath("//input[@name='Name']" )).sendKeys("John");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													Thread.sleep(2000);
						  //Relationship
													driver.findElement(By.xpath("//input[@name='Relation']" )).sendKeys("Single");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													Thread.sleep(2000);
						 //Home Telephone
													driver.findElement(By.xpath("//input[@name='Home Telephone']" )).sendKeys("+1 6758584958");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													Thread.sleep(2000);
						 //Mobile Number
													driver.findElement(By.xpath("//input[@name='Mobile Number']" )).sendKeys("+1 7586985872");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													Thread.sleep(2000);
						 //Work Telephone
													driver.findElement(By.xpath("//input[@name='Mobile Number']" )).sendKeys("+1 689879578");			
													 Thread.sleep(1000);
													driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys("admin123");
													Thread.sleep(1000);
													driver.findElement(By.xpath("//button[@type=\"submit\"]")).click(); //Click on login button
													Thread.sleep(2000);
							    
							    
							    
							    
							    
							    }
							    

							    driver.quit();
							    workbook.close();
					     	} 
						    catch (Exception e) {
							   System.out.println(e.getMessage());
					     	}
					}

			

				


			}



	}

}
