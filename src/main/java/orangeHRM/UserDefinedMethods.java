package orangeHRM;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.NoSuchWindowException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.MediaEntityBuilder;

import io.github.bonigarcia.wdm.WebDriverManager;
//import io.netty.handler.timeout.TimeoutException;

public class UserDefinedMethods extends ReadWriteExcel {

	public static WebDriver driver;
	static WebDriverWait wait;

	static String filePath = "D:\\Eclipse\\eclipse-workspace\\Automation\\Command.xlsx";
//	public static String textFilePath = "C:\\Users\\churi\\Desktop\\OrangeHRM\\orangeHrm.txt";

	static String resourceSheetName = "Resources";
	static int pathColumnNumber = -2;
	static int locatorTypeColumnNumber = -2;
	static String locatorType = "";
	static String path = "";

	// Method to Launch Different Browser based on users choice.
	public static String launchBrowser(String browser) {
		try {
			writeToTextFile("~launchBrowser(" + browser + ")#steps {\\n");
			if (browser.equals("Firefox")) {
				WebDriverManager.edgedriver().setup();
				driver = new FirefoxDriver();
			} else if (browser.equals("Chrome")) {
				WebDriverManager.chromedriver().setup();				
//				System.setProperty("webdriver.chrome.driver",
//						"C:\\Program Files (x86)\\Simplify3x\\SimplifyQA\\libs\\drivers\\chromedriver.exe");
				driver = new ChromeDriver();
			} else if (browser.equals("Edge")) {
				WebDriverManager.edgedriver().setup();
//				System.setProperty("webdriver.edge.driver",
//						"C:\\Program Files (x86)\\Simplify3x\\SimplifyQA\\libs\\drivers\\msedgedriver.exe");
				driver = new EdgeDriver();
			}

			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

			pathColumnNumber = retrieveColumnNumber(resourceSheetName, "Path");
			locatorTypeColumnNumber = retrieveColumnNumber(resourceSheetName, "Locator Type");

			writeToTextFile("}");
			return "";
		} catch (NoSuchWindowException e) {
			writeToTextFile("}");
			return "Window Closed/Not Found";
		} catch (WebDriverException e) {
			writeToTextFile("}");
			return "Web Driver Not Found";
		} catch (Exception e) {
			writeToTextFile("}");
			return "Failed to Create Session";
		}
	}

	// Method to Login to the Website
	public static boolean adminLogin(String username, String password, String methodName, ExtentTest loginTest)
			throws IOException {

		writeToTextFile("~Login(" + username + "," + password + "," + methodName + ")#steps {");

		driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login");

		int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Username");
		int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Password");
		int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Button");

		if (rowNo >= 0 && rowNo2 >= 0 && rowNo3 >= 0) {

			// Retrieving the path of username
			locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
			path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

			enterText(locatorType, path, username);

			// Retrieving the path of Password
			locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
			path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

			enterText(locatorType, path, password);

			// Retrieving the path of Button
			locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
			path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

			click(locatorType, path);

			wait = new WebDriverWait(driver, Duration.ofSeconds(10));

			try {

				rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Invalid credentials");
				rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Account disabled");

				if (rowNo >= 0 && rowNo2 >= 0) {

					// Retrieving the path of Invalid credentials
					locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

					if (checkEmpty(locatorType, path)) {

						// Retrieving the path of Account disabled
						locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

						if (checkEmpty(locatorType, path)) {
							System.out.println("Login Successful");
							reportPassCase(loginTest, username + " Login Successfull");
							writeToTextFile(" }");
							return true;
						} else {
							System.out.println("Account is disabled");
							reportFailCase(loginTest, username + " Account is Disabled", "Login Page");
							writeToTextFile(" }");
							return false;
						}
					} else {
						System.out.println("Invalid Credentials");
						reportFailCase(loginTest, "Invalid Credentials", "Login Page");
						writeToTextFile(" }");
						return false;
					}

				} else {
					System.out.println("Xpath Not Found");
					reportFailCaseWitoutScreenshot(loginTest, "Xpath Not Found");
					writeToTextFile(" }");
					return false;
				}

			} catch (NoSuchWindowException e) {
				System.out.println("Window Closed/Not Found");
				reportFailCaseWitoutScreenshotException(loginTest, e);
				writeToTextFile(" }");
				return false;
			} catch (NoSuchElementException e) {
				System.out.println("Not Found Element Inside Login Function");
				reportFailCaseWitoutScreenshotException(loginTest, e);
				e.printStackTrace();
				writeToTextFile(" }");
				return false;
			} catch (Exception e) {
				reportFailCaseWitoutScreenshotException(loginTest, e);
				e.printStackTrace();
				writeToTextFile(" }");
				return false;
			}
		} else {
			System.out.println("Xpath Not Found");
			reportFailCaseWitoutScreenshot(loginTest, "Xpath Not Found");
			writeToTextFile(" }");
			return false;
		}

	}

	public static void checkResult(ExtentTest test, String expectedResult, String actualResult) {

		if (expectedResult.equalsIgnoreCase(actualResult)) {
			reportPassCase(test, actualResult);
		} else {
			reportFailCase(test, expectedResult, actualResult);
		}
	}

	// Method to Login to the Website
	public static boolean login(int currentRowNo, String sheetName, String methodName, ExtentTest loginTest)
			throws IOException {

		writeToTextFile("~Login(" + currentRowNo + "," + sheetName + "," + methodName + ")#steps {\n");

		// Logging In as a User
		int usernameColumn = retrieveColumnNumber("Sheet1", "Username");
		int passwordColumn = retrieveColumnNumber("Sheet1", "Password");
		int expectedResultColumn = retrieveColumnNumber("Sheet1", "Expected Result");

		String username = getExcelData(currentRowNo, usernameColumn, "Sheet1");
		String password = getExcelData(currentRowNo, passwordColumn, "Sheet1");
		String expectedResult = getExcelData(currentRowNo, expectedResultColumn, "Sheet1");

		driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login");

		int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Username");
		int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Password");
		int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Button");

		if (rowNo >= 0 && rowNo2 >= 0 && rowNo3 >= 0) {

			// Retrieving the path of username
			locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
			path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

			enterText(locatorType, path, username);

			// Retrieving the path of Password
			locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
			path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

			enterText(locatorType, path, password);

			// Retrieving the path of Button
			locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
			path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

			click(locatorType, path);

			wait = new WebDriverWait(driver, Duration.ofSeconds(10));

			try {

				rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Invalid credentials");
				rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Account disabled");

				if (rowNo >= 0 && rowNo2 >= 0) {

					// Retrieving the path of Invalid credentials
					locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

					if (checkEmpty(locatorType, path)) {

						// Retrieving the path of Account disabled
						locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

						if (checkEmpty(locatorType, path)) {
							// Login Successful
							// Checking the Expected Result and Actual Result
							checkResult(loginTest, expectedResult, "Login Successful");
							writeToTextFile(" }");
							return true;
						} else {
							checkResult(loginTest, expectedResult, "Account is disabled");
							writeToTextFile(" }");
							return false;
						}
					} else {
						checkResult(loginTest, expectedResult, "Login failed. Invalid Credentials.");
						writeToTextFile(" }");
						return false;
					}
				} else {
					System.out.println("Xpath Not Found");
					checkResult(loginTest, expectedResult, "Xpath Not Found");
					writeToTextFile(" }");
					return false;
				}
			} catch (NoSuchWindowException e) {
				System.out.println("Window Closed/Not Found");
				reportFailCaseWitoutScreenshotException(loginTest, e);
				writeToTextFile(" }");
				return false;
			} catch (NoSuchElementException e) {
				System.out.println("Not Found Element Inside Login Function");
				reportFailCaseWitoutScreenshotException(loginTest, e);
				e.printStackTrace();
				writeToTextFile(" }");
				return false;
			} catch (Exception e) {
				reportFailCaseWitoutScreenshotException(loginTest, e);
				e.printStackTrace();
				writeToTextFile(" }");
				return false;
			}
		} else {
			System.out.println("Xpath Not Found");
			checkResult(loginTest, expectedResult, "Xpath Not Found");
			writeToTextFile(" }");
			return false;
		}

	}

	// Method to read data from a specific cell.
	public static String getExcelData(int row, int column, String sheetName) {

		writeToTextFile("getExcelData(" + row + "," + column + "," + sheetName + "), ");
		return readExcelSpecificCellData(filePath, sheetName, row, column);

	}

	// Method to Write data into a specific cell.
	public static void updateExcelSheet(int row, int column, String sheetName, String data) {
		writeToTextFile("updateExcelSheet(" + row + "," + column + "," + sheetName + "," + data + ")");
		writeExcelSpecificCellData(filePath, sheetName, row, column, data);

	}

	// Method to Enter the Text.
	public static void enterText(String type, String path, String value) {
		writeToTextFile("enterText(" + type + "," + path + "," + value + ")");
		By locator = getBy(type, path);
		driver.findElement(locator).sendKeys(value);

	}

	public static boolean checkEmpty(String type, String path) {
		By locator = getBy(type, path);
		writeToTextFile("checkEmpty(" + type + "," + path + "), ");
		return driver.findElements(locator).isEmpty();

	}

	// Method to Clear the existing text and Enter the new text in single click.
	public static void clearAndEnterText(String type, String path, String value) {
		writeToTextFile("clearAndEnterText(" + type + "," + path + "," + value + "), ");
		By locator = getBy(type, path);
		driver.findElement(locator).sendKeys(Keys.CONTROL + "a");
		driver.findElement(locator).sendKeys(Keys.DELETE);
		driver.findElement(locator).sendKeys(value);

	}

	// Method to click a web element. Returns True if Successful else False
	public static boolean click(String type, String path) {
		By locator = getBy(type, path);
		try {
			writeToTextFile("click(" + type + "," + path + "), ");
			driver.findElement(locator).click();
			return true;
		} catch (NoSuchWindowException e) {
			System.out.println("Window Closed/Not Found");
			return false;
		} catch (NoSuchElementException e) {
			System.out.println(path + " " + "Not Found in the Web Page");
			return false;
		} catch (Exception e) {
			System.out.println("Error: Unable to Perform Click in " + path);
			return false;
		}
	}

	// Method to get the read the static data from the Web Page
	public static String getData(String type, String path) {
		By locator = getBy(type, path);
		writeToTextFile("getData(" + type + "," + path + "), ");
		return driver.findElement(locator).getText();
	}

	// Method to Use different types of selectors
	public static By getBy(String type, String path) {
		switch (type.toLowerCase()) {
		case "id":
			return By.id(path);
		case "name":
			return By.name(path);
		case "xpath":
			return By.xpath(path);
		case "classname":
			return By.className(path);
		case "cssselector":
			return By.cssSelector(path);
		case "linktext":
			return By.linkText(path);
		case "partiallinktext":
			return By.partialLinkText(path);
		case "tagname":
			return By.tagName(path);
		case "href":
			return By.cssSelector(path);
		default:
			throw new IllegalArgumentException("Unsupported locator type: " + type);
		}
	}

	// Checking whether the Web Element is present or not. Return False is not
	// visible and
	// True is Visible
	public static boolean checkVisibility(String type, String path) {

//		wait = new WebDriverWait(driver, Duration.ofSeconds(10));
//
//		By locator = getBy(type, path);
//		try {
//			if (driver.findElements(locator).isEmpty()) {
//				return true;
//			} else {
//				return false;
//			}
//		} catch (Exception e) {
//			e.printStackTrace();
//			return true;
//		}

		try {
			writeToTextFile("checkVisibility(" + type + "," + path + "), ");
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
			wait.until(ExpectedConditions.visibilityOf(driver.findElement(getBy(type, path))));
			return true;
		} catch (TimeoutException e) {
			// Element is not visible or not found
			return false;
		} catch (NoSuchElementException e) {
			return false;
		} catch (Exception e) {
			System.out.println("Error occur in exception block in checkVisibility");
			e.printStackTrace();
			return false;
		}

	}

	// Method to get the Dynamic Value from the Web Page.
	public static String getValue(String type, String path) {
		By locator = getBy(type, path);
		String eid = driver.findElement(locator).getDomProperty("value");
		writeToTextFile("getValue(" + type + "," + path + "), ");
		return eid;
	}

	// Method to Validate the Employee Creation
	public static boolean validateNewEmployee(String firstName, String lastName, String id, String methodName) {

		try {
			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Records Check");

			if (rowNo >= 0) {
				// Retrieving the path of No Records Found
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				if (checkVisibility(locatorType, path)) {

					System.out.println("No Records Found");
					return false;
				}

				rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "First Name");
				int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Last Name");
				int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Id");

				if (rowNo >= 0 && rowNo2 >= 0 && rowNo3 >= 0) {

					// Retrieving the path of First Name
					String firstNamelocatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
					String firstNameXpath = getExcelData(rowNo, pathColumnNumber, resourceSheetName);
					firstNameXpath = replaceText(firstNameXpath, firstName);

					// Retrieving the path of Last Name
					String lastNamelocatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					String lastNameXpath = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);
					lastNameXpath = replaceText(lastNameXpath, lastName);

					// Retrieving the path of Id
					String idXpathlocatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					String idXpath = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);
					idXpath = replaceText(idXpath, id);

					boolean isFirstNamePresent = checkEmpty(firstNamelocatorType, firstNameXpath);
					boolean isLastNamePresent = checkEmpty(lastNamelocatorType, lastNameXpath);
					boolean isIdPresent = checkEmpty(idXpathlocatorType, idXpath);

					if (isFirstNamePresent && isLastNamePresent && isIdPresent) {
						return true;
					} else {
						return false;
					}

				} else {
					return false;
				}

			} else {
				return false;
			}
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}

	}

	public static String replaceText(String originalString, String replacementText) {
		// Replace the placeholder with the actual text
		return originalString.replace("dynamicData", replacementText);
	}

	// Method to Punch In and Punch Out
	public static void lodgePunchInOut(String methodName, String expectedResult, ExtentTest testReport)
			throws InterruptedException {

		try {
			writeToTextFile("~lodgePunchInOut(" + methodName + "," + expectedResult + ")# steps { ");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Time Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Attendance");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Punch In/Out");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0) {
				// Retrieving the path of Time Module
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Opening Time Web Page
				click(locatorType, path);

				// Retrieving the path of Attendance
				locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

				// Opening Attendance Section
				click(locatorType, path);

				// Retrieving the path of Punch In/Out
				locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

				// Opening Punch In/Out Web Page
				click(locatorType, path);

				rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Punch In Text");

				if (rowNo >= 0) {
					// Retrieving the path of Punch In Label
					locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

					if ("Punch In".equals(expectedResult)) {
						// Checking whether the Punching Test is visible
						if (checkVisibility(locatorType, path)) {
							// Performing Punch In action
							rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Punch In");

							if (rowNo >= 0) {
								locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
								path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

								// To Click Punch In Button
								click(locatorType, path);

								rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName,
										"Successfully Saved");

								path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

								if (checkVisibility(locatorType, path)) {
									checkResult(testReport, expectedResult, "Punch In");
								} else {
									checkResult(testReport, expectedResult, "Failed to Punch In");
								}

							} else {
								checkResult(testReport, expectedResult, "Punch In Button Xpath not found");
							}

						} else {
							System.out.println("User already Punch In");
							checkResult(testReport, expectedResult, "User already Punched In");
						}

					} else {
						// Checking whether the Punching Test is Not visible
						if (checkVisibility(locatorType, path) == false) {
							// Performing Punch Out action
							rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Punch Out");

							if (rowNo >= 0) {
								locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
								path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

								// To Click Punch Out Button
								click(locatorType, path);

								rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName,
										"Successfully Saved");

								path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

								if (checkVisibility(locatorType, path)) {
									checkResult(testReport, expectedResult, "Punch Out");
								} else {
									checkResult(testReport, expectedResult, "Failed to Punch Out");
								}

							} else {
								checkResult(testReport, expectedResult, "Punch Out Button Xpath not found");
							}
						} else {
							System.out.println("User already Punch Out");
							checkResult(testReport, expectedResult, "User already Punch Out");
						}

					}
				} else {
					System.out.println("Xpath Not Found");
					checkResult(testReport, expectedResult, "Xpath Not Found");
				}
			} else {
				System.out.println("Xpath Not Found");
				checkResult(testReport, expectedResult, "Xpath Not Found");

			}
			writeToTextFile(" } ");
		} catch (IOException e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		}
	}

	// Method to Retrieve the last Row Number
	public static int retrieveLastRow(String sheetName) throws IOException {
		writeToTextFile("retrieveLastRow(" + sheetName + "), ");
		return getLastRowNumber(filePath, sheetName);
	}

	// Method to Retrieve the Column Number of the Cell
	public static int retrieveColumnNumber(String sheetName, String searchText) throws IOException {

		writeToTextFile("retrieveColumnNumber(" + sheetName + "," + searchText + "), ");
		return findColumnNumber(filePath, sheetName, searchText);
	}

	// Method to Retrieve the Row Number of the Cell
	public static int retrieveRowNumber(String sheetName, String searchText) throws IOException {
		writeToTextFile("retrieveRowNumber(" + sheetName + "," + searchText + "), ");
		return findRowNumber(filePath, sheetName, searchText);
	}

	public static void message(String msg) {
		if (!msg.equals(""))
			System.out.println(msg);

	}

	// Method to Search employee based on Employee Name
	public static boolean searchEmployee(String firstName, String lastName, String methodName, String expectedResult,
			ExtentTest searchEmployeeTest) {

		try {
			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Pim Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee List");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Name");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Search");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0) {
				// Retrieving the path of PIM Module
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				if (click(locatorType, path)) {
					// Retrieving the path of Employee List
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					// Retrieving the path of Employee Name
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					// Entering Employee Name in the field
					enterText(locatorType, path, firstName + " " + lastName);

					Thread.sleep(1000);

					// Retrieving the path of Search Button
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					// Clicking Search Button
					click(locatorType, path);

					rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "No Records Found");

					// Checking if Records found or not
					locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

					// Checking if Records found or not
					if (checkEmpty(locatorType, path) == false) {
						System.out.println("No Records Found");
						checkResult(searchEmployeeTest, expectedResult, "No Records Found");
						return true;
					} else {
						checkResult(searchEmployeeTest, expectedResult, "Employee Found Successfully");
						return false;
					}
				} else {
					System.out.println("You Don't have access to search Employee.");
					checkResult(searchEmployeeTest, expectedResult, "You Don't have access to search Employee.");
					return false;
				}

			} else {
				System.out.println("Xpath Not Found");
				checkResult(searchEmployeeTest, expectedResult, "Xpath Not Found");
				return false;
			}

		} catch (IOException e) {
			reportFailCaseWitoutScreenshotException(searchEmployeeTest, e);
			return false;
		} catch (Exception e) {
			reportFailCaseWitoutScreenshotException(searchEmployeeTest, e);
			return false;
		}
	}

	// Method to Search employee based on Employee Id
	public static boolean searchEmployee(String empId, String methodName, String expectedResult,
			ExtentTest searchEmployeeTest) throws InterruptedException {

		try {

			writeToTextFile("~searchEmployee(" + empId + "," + methodName + "," + expectedResult + ")# steps { ");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Pim Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee List");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Id");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Search");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0) {
				// Retrieving the path of PIM Module
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				if (click(locatorType, path)) {
					// Retrieving the path of Employee List
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					// Retrieving the path of Employee Id
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					// Entering Employee Id in the field
					enterText(locatorType, path, empId);

					Thread.sleep(1000);

					// Retrieving the path of Search Button
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					// Clicking Search Button
					click(locatorType, path);

					rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "No Records Found");

					// Checking if Records found or not
					locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

					// Checking if Records found or not
					if (checkEmpty(locatorType, path) == false) {
						System.out.println("No Records Found");
						checkResult(searchEmployeeTest, expectedResult, "No Records Found");
						writeToTextFile(" } ");
						return false;
					} else {
						checkResult(searchEmployeeTest, expectedResult, "Employee Found Successfully");
						writeToTextFile(" } ");
						return true;
					}
				} else {
					System.out.println("You Don't have access to search Employee.");
					checkResult(searchEmployeeTest, expectedResult, "You Don't have access to search Employee.");
					writeToTextFile(" } ");
					return false;
				}
			} else {
				System.out.println("Xpath Not Found");
				checkResult(searchEmployeeTest, expectedResult, "Xpath Not Found");
				writeToTextFile(" } ");
				return false;
			}

		} catch (IOException e) {
			reportFailCaseWitoutScreenshotException(searchEmployeeTest, e);
			e.printStackTrace();
			writeToTextFile(" } ");
			return false;
		}

	}

	// Driver Close Method
	public static void finishExecution() {
		driver.quit();
	}

	// Method to Post New News Feed
	public static void postNewsfeed(String data, String methodName, String expectedResult,
			ExtentTest postNewsFeedTest) {

		try {
			Thread.sleep(300);
			writeToTextFile("#postNewsfeed(" + data + "," + methodName + "," + expectedResult + ")# steps { ");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Buzz Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Post Field");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Post Button");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0) {
				// Retrieving the path of Buzz Module
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				if (click(locatorType, path)) {
					Thread.sleep(300);
					// Retrieving the path of Post Field
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					enterText(locatorType, path, data);

					// Retrieving the path of Post Button
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					// Clicking the Post button
					click(locatorType, path);

					int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName,
							"Successfully Saved");

					// Retrieving the path of Successfully Saved
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					if (checkVisibility(locatorType, path)) {
						System.out.println("Newsfeed Posted Successfully");
						checkResult(postNewsFeedTest, expectedResult, "Newsfeed Posted Successfully");
					} else {
						System.out.println("Failed to Post News Feed");
						checkResult(postNewsFeedTest, expectedResult, "Failed to Post News Feed");
					}

				} else {
					System.out.println("Failed to Open Buzz Module");
					checkResult(postNewsFeedTest, expectedResult, "Failed to Open Buzz Module");

				}
			} else {
				System.out.println("Xpath Not Found");
				checkResult(postNewsFeedTest, expectedResult, "Xpath Not Found");
			}
			writeToTextFile(" } ");
		} catch (IOException e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(postNewsFeedTest, e);
			writeToTextFile(" } ");
		} catch (NoSuchElementException e) {
			System.out.println("Failed to Post the Newsfeed");
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(postNewsFeedTest, e);
			writeToTextFile(" } ");
		} catch (Exception e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(postNewsFeedTest, e);
			writeToTextFile(" } ");
		}

	}

	// Method to Delete the News Feed
	public static void deleteNewsfeed(String methodName, ExtentTest deleteTestReport) {

		try {
			Thread.sleep(300);
			writeToTextFile("~deleteNewsfeed(" + methodName + ")# steps { ");
			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Buzz Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Dot icon");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Delete Post");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Yes, Delete");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0) {
				// Retrieving the path of Buzz Module
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				if (click(locatorType, path)) {
					Thread.sleep(300);
					// Retrieving the path of Dot Icon
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					// Retrieving the path of Delete Post
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					// Clicking the Delete Post
					click(locatorType, path);

					// Retrieving the path of Yes, Delete
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

//					// Locate the element you want to click
//			        WebElement elementToClick = driver.findElement(By.xpath(path));
//
//			        // Use WebDriverWait to wait for the element to be clickable
//			        WebElement clickableElement = wait.until(ExpectedConditions.elementToBeClickable(elementToClick));

					// Clicking the Yes, Delete
//			        clickableElement.click();
					if (click(locatorType, path)) {
						System.out.println("Newsfeed Deleted Successfully");
						checkResult(deleteTestReport, "Newsfeed Deleted Successfully", "Newsfeed Deleted Successfully");
					} else {
						System.out.println("Failed to Delete News Feed");
						checkResult(deleteTestReport, "Failed to Delete News Feed", "Failed to Delete News Feed");
					}

					int rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName,
							"Successfully Deleted");

					// Retrieving the path of Successfully Saved
					locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					wait = new WebDriverWait(driver, Duration.ofSeconds(10));

					WebElement ele = driver.findElement(By.xpath(path));
					wait.until(ExpectedConditions.visibilityOf(ele));

				}

			} else {
				System.out.println("Xpath Not Found");
			}
			writeToTextFile(" } ");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			writeToTextFile(" } ");
		} catch (NoSuchElementException e) {
			System.out.println("Failed to Delete the Newsfeed");
			writeToTextFile(" } ");
		} catch (Exception e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		}

	}

	// Method to Update the Contact Details of the User
	public static void updateContactDetails(int i, int additionalInfoColumnNumber, String methodName,
			String expectedResult, ExtentTest contactDetailsUpdateTest) {

		try {

			writeToTextFile("~updateContactDetails(" + i + "," + additionalInfoColumnNumber + "," + methodName + ","
					+ expectedResult + ")# steps { ");

			String numberType = getExcelData(i, additionalInfoColumnNumber, "Task");
			String number = getExcelData(i, additionalInfoColumnNumber + 1, "Task");
			String emailType = getExcelData(i, additionalInfoColumnNumber + 2, "Task");
			String email = getExcelData(i, additionalInfoColumnNumber + 3, "Task");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "My Info");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Contact Details");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Contact Number");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Contact Email");
			int rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Submit Button");
			int rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Successfully Updated");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0 && rowNo6 > 0) {
				// Retrieving the path of My Info
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Opening the My Info Page
				if (click(locatorType, path)) {
					// Retrieving the path of Contact Details
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					// Opening the Contact Details Page
					if (click(locatorType, path)) {

						try {
							Thread.sleep(10000);
						} catch (InterruptedException e) {
							e.printStackTrace();
						}

						locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

						path = replaceText(path, numberType);

						// Entering the Mobile Number
						clearAndEnterText(locatorType, path, number);

						locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

						path = replaceText(path, emailType);

						// Entering the Email Id
						clearAndEnterText(locatorType, path, email);

						locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

						// Scrolling and clicking the "Save" button
						scrollIntoView(locatorType, path);
						click(locatorType, path);

						path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);

						// Checking whether the Successful message Visible or Not
						if (checkVisibility(locatorType, path)) {
							System.out.println("Contact Details Updated Successfully");
							checkResult(contactDetailsUpdateTest, expectedResult,
									"Contact Details Updated Successfully");
						} else {
							System.out.println("Failed to Update Contact Details");
							checkResult(contactDetailsUpdateTest, expectedResult, "Failed to Update Contact Details");
						}

					} else {
						System.out.println("Failed to Open Contact Details Section");
						checkResult(contactDetailsUpdateTest, expectedResult, "Failed to Open Contact Details Section");
					}
				} else {
					System.out.println("Failed to Open My Info Module");
					checkResult(contactDetailsUpdateTest, expectedResult, "Failed to Open My Info Module");
				}
			} else {
				System.out.println("Xpath Not Found");
				checkResult(contactDetailsUpdateTest, expectedResult, "Xpath Not Found");
			}
			writeToTextFile(" } ");
		} catch (IOException e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(contactDetailsUpdateTest, e);
			writeToTextFile(" } ");
		} catch (NoSuchElementException e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(contactDetailsUpdateTest, e);
			writeToTextFile(" } ");
		} catch (Exception e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(contactDetailsUpdateTest, e);
			writeToTextFile(" } ");
		}

	}

	// Javascript Executor Method to Scroll the page.
	public static void scrollIntoView(String locatorType, String locatorValue) {
		WebElement element = driver.findElement(getBy(locatorType, locatorValue));
		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", element);
	}

	// Method to assign Leave to any employee
	public static void assignLeave(int i, int additionalInfoColumnNumber, String methodName, String expectedResult,
			ExtentTest assignLeaveTest) {

		try {

			writeToTextFile("~assignLeave(" + i + "," + additionalInfoColumnNumber + "," + methodName + ","
					+ expectedResult + ")# steps { ");

			String employeeName = getExcelData(i, additionalInfoColumnNumber, "Task");
			String leaveType = getExcelData(i, additionalInfoColumnNumber + 1, "Task");
			String fromDate = getExcelData(i, additionalInfoColumnNumber + 2, "Task");
			String toDate = getExcelData(i, additionalInfoColumnNumber + 3, "Task");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Leave Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Assign Leave");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Name");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Select Employee Name");
			int rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Leave Type Option");
			int rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Leave Type Value");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0 && rowNo6 > 0) {
				// Opening Leave Module and Selecting the Assign Leave Option
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Opening the Leave Module Page
				if (click(locatorType, path)) {
					// Retrieving the path of Assign Leave
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					// Opening the Assign Leave Page
					click(locatorType, path);

					// Retrieving the path of Employee Name
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					// Inserting the Employees Name
					enterText(locatorType, path, employeeName);

					// Retrieving the path of Select Employee Name
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					if (click(locatorType, path)) {

						// Retrieving the path of Leave Type Option
						locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

						// Selecting the Leave Type
						click(locatorType, path);

						// Retrieving the path of Leave Type Value
						locatorType = getExcelData(rowNo6, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);

						path = replaceText(path, leaveType);

						click(locatorType, path);

						rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "From Date");
						rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "To Date");
						rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Invalid Date");

						rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Partial Days");
						rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "All Days");
						rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Leave Duration");
						int commentRow = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Comments");

						if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0 && rowNo6 > 0
								&& commentRow > 0) {
							// Selecting From Date
							locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

							// Entering the From Date
							enterText(locatorType, path, fromDate);

							// Selecting To Date
							locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

							// Entering the To Date
							clearAndEnterText(locatorType, path, toDate);

							// Retrieving the Invalid Date Text Path
							locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

							// Checking whether the date is valid or not.
							if (checkVisibility(locatorType, path) == false) {

								// Retrieving the path of the Comment Label
								locatorType = getExcelData(commentRow, locatorTypeColumnNumber, resourceSheetName);
								path = getExcelData(commentRow, pathColumnNumber, resourceSheetName);

								// Clicking on the Comment label to refresh page.
								click(locatorType, path);

								// Retrieving the path of Partial Days
								locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
								path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

								// Checking whether the Partial Days option visible or not. If yes then
								// selecting All Days
								if (checkVisibility(locatorType, path)) {

									click(locatorType, path);

									// Retrieving the path of the All Days Option
									locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
									path = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

									click(locatorType, path);

									// Retrieving the path of the Leave Duration
									locatorType = getExcelData(rowNo6, locatorTypeColumnNumber, resourceSheetName);
									path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);

									// Selecting the Leave Duration as Half Day - Morning
									click(locatorType, path);

									rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName,
											"Half Day - Morning");

									if (rowNo6 > 0) {

										// Retrieving the path of Half Day - Morning
										locatorType = getExcelData(rowNo6, locatorTypeColumnNumber, resourceSheetName);
										path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);

										click(locatorType, path);
									} else {
										System.out.println("Xpath Not Found");
										checkResult(assignLeaveTest, expectedResult, "Xpath Not Found");
									}
								}

								rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Assign");
								rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Ok");
								rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName,
										"Successfully Saved");

								if (rowNo > 0 && rowNo2 > 0 && rowNo3 > 0) {

									locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
									path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

									// Assigning the Leave
									click(locatorType, path);

									locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
									path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

									// Checking if getting the message of insufficient leave or not if yes then
									// selecting ok
									if (checkVisibility(locatorType, path))
										click(locatorType, path);

									locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
									path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

									checkResult(assignLeaveTest, expectedResult, "Leave Assigned Successfully");

								} else {
									System.out.println("Xpath Not Found");
									checkResult(assignLeaveTest, expectedResult, "Xpath Not Found");
								}

							} else {
								System.out.println("End Date should be after the Start Date");
								checkResult(assignLeaveTest, expectedResult, "End Date should be after the Start Date");
							}
						} else {
							System.out.println("Xpath Not Found");
							checkResult(assignLeaveTest, expectedResult, "Xpath Not Found");
						}

					} else {
						System.out.println("Employee Not Found");
						checkResult(assignLeaveTest, expectedResult, "Employee Not Found");
					}
				} else {
					System.out.println("Leave Module Not Found");
					checkResult(assignLeaveTest, expectedResult, "Leave Module Not Found");
				}
			} else {
				System.out.println("Xpath Not Found");
				checkResult(assignLeaveTest, expectedResult, "Xpath Not Found");
			}
			writeToTextFile(" } ");
		} catch (IOException e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(assignLeaveTest, e);
			writeToTextFile(" } ");
		} catch (NoSuchElementException e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(assignLeaveTest, e);
			writeToTextFile(" } ");
		} catch (Exception e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(assignLeaveTest, e);
			writeToTextFile(" } ");
		}
	}

	// Method to add new Vacancy
	public static void addVacancy(int i, int additionalInfoColumnNumber, String methodName) {

		try {

			writeToTextFile("~addVacancy(" + i + "," + additionalInfoColumnNumber + "," + methodName + ")# steps { ");

			String vacancyName = getExcelData(i, additionalInfoColumnNumber, "Task");
			String jobTitle = getExcelData(i, additionalInfoColumnNumber + 1, "Task");
			String description = getExcelData(i, additionalInfoColumnNumber + 2, "Task");
			String hirringManager = getExcelData(i, additionalInfoColumnNumber + 3, "Task");
			String numberOfPosition = getExcelData(i, additionalInfoColumnNumber + 4, "Task");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Recruitment Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Vacancies");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Add");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Vacancy Name");
			int rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Job Title Option");
			int rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Job Title Value");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0 && rowNo6 > 0) {
				// Opening Recruitment Module and Selecting the Vacancies Option
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Opening the Recruitment Module Page
				if (click(locatorType, path)) {
					// Retrieving the path of Vacancies
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					// Opening the Vacancies Page
					click(locatorType, path);

					// Retrieving the path of Add Button
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					// Clicking the Add Button
					click(locatorType, path);

					// Retrieving the path of Vacancy Name Field
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					// Inserting the Vacancy Name
					enterText(locatorType, path, vacancyName);

					// Retrieving the path of Job Title Option
					locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

					// Selecting the Job Title
					click(locatorType, path);

					// Retrieving the path of Job Title Value
					locatorType = getExcelData(rowNo6, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);
					path = replaceText(path, jobTitle);
					click(locatorType, path);

					rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Description");
					rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Hiring Manager");
					rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Hiring Manager Select");
					rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Number of Positions");
					rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Save");

					if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0) {

						// Retrieving the path of Description
						locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

						// Inserting the Description
						enterText(locatorType, path, description);

						// Retrieving the path of Hiring Manager
						locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

						// Selecting the Hirring Manager
						enterText(locatorType, path, hirringManager);

						// Retrieving the path of Hiring Manager Select
						locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

						if (click(locatorType, path)) {

							// Retrieving the path of Number of Positions
							locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

							enterText(locatorType, path, numberOfPosition);

							// Retrieving the path of Save Button
							locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

							click(locatorType, path);

						} else {
							System.out.println("Hirring Manager Not Found");
						}

					} else {
						System.out.println("Xpath not found");
					}

				} else {
					System.out.println("You don't have access to Recruitment Module.");
				}

			} else {
				System.out.println("Xpath Not Found");
			}

		} catch (IOException e) {
			e.printStackTrace();

		} catch (NoSuchElementException e) {
			System.out.println("Failed to Post the Newsfeed");
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// Method to Validate and Convert the Date Format
	public static String validateAndConvert(String inputDate) {
		// Check if the inputDate is in yyyy-mm-dd format
		if (inputDate.matches("\\d{4}-\\d{2}-\\d{2}")) {
			return inputDate; // Already in the correct format
		}

		// Attempt to parse the inputDate
		try {
			Date date = new SimpleDateFormat("dd-MM-yyyy").parse(inputDate);

			// Format the parsed date into yyyy-mm-dd
			return new SimpleDateFormat("yyyy-MM-dd").format(date);
		} catch (ParseException e) {
			System.out.println("Invalid date format: " + inputDate);
			return null; // Return null if unable to parse or format
		}
	}

	// Return True If User Employee Found else False
	public static boolean validateEmployeeCreation(String firstName, String middleName, String lastName, String empId,
			String methodName) {

		try {
			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Pim Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee List");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Id");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Search");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0) {
				// Opening Pim Module
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Opening the Pim Module
				if (click(locatorType, path)) {
					// Retrieving the path of Employee List Section
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					// Opening the Employee List Section
					click(locatorType, path);

					// Retrieving the path of Employee Id Field
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					wait = new WebDriverWait(driver, Duration.ofSeconds(10));
					WebElement ele = driver.findElement(By.xpath(path));
					wait.until(ExpectedConditions.visibilityOf(ele));

					enterText(locatorType, path, empId);

					// Retrieving the path of Search Button
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					if (validateNewEmployee(firstName + " " + middleName, lastName, empId, "validateNewEmployee")) {
						return true;
					} else {
						return false;
					}

				} else {
					System.out.println("Unable to open Pim Module.");
					return false;
				}

			} else {
				System.out.println("Xpath Not Found");
				return false;
			}

		} catch (IOException e) {
			e.printStackTrace();
			return false;

		} catch (NoSuchElementException e) {
			System.out.println("Failed to Post the Newsfeed");
			e.printStackTrace();
			return false;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}

	}

	// Method to Delete Existing Employee
	public static void deleteEmployee(String empId, String methodName) {
		try {
			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Record Found");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Id");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Yes, Delete");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Successfully Saved");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0) {
				// Retrieving the Path of Number of Records
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Checking whether one records found or multiple
				if (checkVisibility(locatorType, path)) {
					System.out.println("Multiple User Found");
				} else {
					// Retrieving the path of Employee Id
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);
					path = replaceText(path, empId);
					click(locatorType, path);

					// Retrieving the Path of Yes, Delete Button

					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					// Locate the element you want to click
					WebElement elementToClick = driver.findElement(By.xpath(path));

					// Use WebDriverWait to wait for the element to be clickable
					WebElement clickableElement = wait.until(ExpectedConditions.elementToBeClickable(elementToClick));
					clickableElement.click();

					// Retrieving the Path of Successfully Saved Message
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					WebElement ele = driver.findElement(By.xpath(path));
					wait.until(ExpectedConditions.visibilityOf(ele));
					System.out.println("Successfully Deleted");

				}

			} else {
				System.out.println("Xpath Not Found");
			}

		} catch (IOException e) {
			e.printStackTrace();

		} catch (NoSuchElementException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// Method to View Employees Time Sheet
	public static void viewEmployeeTimeSheet(String name, String methodName, String expectedResult,
			ExtentTest myTimeSheetTest) {

		try {
			writeToTextFile("~viewEmployeeTimeSheet(" + name + "," + methodName + "," + expectedResult + ")# steps { ");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Time Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Timesheets");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Timesheets");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "My Timesheets");
			int rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Name");
			int rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Name Select");
			int rowNo7 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "View");
			int rowNo8 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Timesheet Visibility");
			int rowNo9 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Invalid Username");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0 && rowNo6 > 0 && rowNo7 > 0
					&& rowNo8 > 0 && rowNo9 > 0) {
				// Retrieving the Path of Time Module
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Opening Time Module
				if (click(locatorType, path)) {

					// Retrieving the Path of Time Module
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					// Retrieving the Path of Employee Timesheets
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					// Viewing the Employees Time Table
					if (click(locatorType, path) == false) {
						System.out.println("You don't have access to view other employee timesheet.");
						checkResult(myTimeSheetTest, expectedResult,
								"You don't have access to view other employee timesheet.");

//						// Retrieving the Path of My Timesheets
//						locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
//						path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);
//
//						click(locatorType, path);
					} else {
						// Retrieving the Path of Employee Name
						locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

						enterText(locatorType, path, name);

						// Retrieving the Path of Employee Name Select
						locatorType = getExcelData(rowNo6, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);

						click(locatorType, path);

						// Retrieving the Path of Invalid Username
						locatorType = getExcelData(rowNo9, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo9, pathColumnNumber, resourceSheetName);

						if (checkVisibility(locatorType, path) == false) {

							// Retrieving the Path of View Button
							locatorType = getExcelData(rowNo7, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo7, pathColumnNumber, resourceSheetName);

							click(locatorType, path);

							// Retrieving the Path of Timesheet Visibility
							locatorType = getExcelData(rowNo8, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo8, pathColumnNumber, resourceSheetName);
							path = replaceText(path, "Time");

							// Checking if the Page opened or not
							if (checkVisibility(locatorType, path)) {
								System.out.println("Timesheets visible successfully");
								checkResult(myTimeSheetTest, expectedResult, "Employee Timesheet is Visible");
							} else {
								System.out.println("Unable to View Employees Time Sheet");
								checkResult(myTimeSheetTest, expectedResult, "Unable to View Employees Time Sheet");
							}

						} else {
							System.out.println("Username not found");
							checkResult(myTimeSheetTest, expectedResult, "Username not found");
						}
					}

				} else {
					System.out.println("Time Table Module Not Found");
					checkResult(myTimeSheetTest, expectedResult, "Time Table Module Not Found");
				}

			} else {
				System.out.println("Xpath Not Found");
				checkResult(myTimeSheetTest, expectedResult, "Xpath Not Found");
			}
			writeToTextFile(" } ");
		} catch (IOException e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(myTimeSheetTest, e);
			writeToTextFile(" } ");
		} catch (NoSuchElementException e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(myTimeSheetTest, e);
			writeToTextFile(" } ");
		} catch (Exception e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(myTimeSheetTest, e);
			writeToTextFile(" } ");
		}

	}

	// Method to View My Time Sheet
	public static void viewMyTimeSheet(String methodName, String expectedResult, ExtentTest myTimeSheetTest) {

		try {
			writeToTextFile("~viewMyTimeSheet(" + methodName + "," + expectedResult + ")# steps { ");
			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Time Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Timesheets");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "My Timesheets");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Timesheet Visibility");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0) {
				// Retrieving the Path of Time Module
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Clicking on Time Module
				if (click(locatorType, path)) {

					// Retrieving the Path of Timesheets
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					// Retrieving the Path of My Timesheets
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					// Viewing the My Time Table
					click(locatorType, path);

					// Retrieving the Path of Timesheet Visibility
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					// Validating whether MY timesheet is visible or not
					if (checkVisibility(locatorType, path)) {
						checkResult(myTimeSheetTest, expectedResult, "My Timesheet is Visible");
					} else {
						checkResult(myTimeSheetTest, expectedResult, "Unable to view My Timesheet");
					}

				} else {
					System.out.println("Time Module Not Found");
					checkResult(myTimeSheetTest, expectedResult, "Time Module Not Found");
				}
			} else {
				System.out.println("Xpath Not Found");
				checkResult(myTimeSheetTest, expectedResult, "Xpath Not Found");
			}
			writeToTextFile(" } ");
		} catch (IOException e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		} catch (NoSuchElementException e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		} catch (Exception e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		}

	}

	// Method to Create New PIM User
	public static void createPIM(int newEmprow, String sheetName, String methodName, ExtentTest creationTest) {

		writeToTextFile("~createPIM(" + newEmprow + "," + sheetName + "," + methodName + ")# steps { ");
		try {

			int firstNameColumnNumber = 0;
			int middleNameColumnNumber = 0;
			int lastNameColumnNumber = 0;
			int newUsernameColumnNumber = 0;
			int newEmployeeIdColumnNumber = 0;
			int newPasswordColumnNumber = 0;

			int expectedResultColumn = retrieveColumnNumber(sheetName, "Expected Result");
			String expectedResult = getExcelData(newEmprow, expectedResultColumn, sheetName);
			System.out.println(expectedResult);

			firstNameColumnNumber = retrieveColumnNumber(sheetName, "First Name");
			middleNameColumnNumber = retrieveColumnNumber(sheetName, "Middle Name");
			lastNameColumnNumber = retrieveColumnNumber(sheetName, "Last Name");
			newEmployeeIdColumnNumber = retrieveColumnNumber(sheetName, "Employee Id");
			newUsernameColumnNumber = retrieveColumnNumber(sheetName, "Username");
			newPasswordColumnNumber = retrieveColumnNumber(sheetName, "Password");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Pim Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Add Employee Button");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Login Credentials");
			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "First Name");
			int rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Middle Name");
			int rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Last Name");
			int rowNo7 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "New Username");
			int rowNo8 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "New Password");
			int rowNo9 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Confirm Password");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0 && rowNo6 > 0 && rowNo7 > 0
					&& rowNo8 > 0 && rowNo9 > 0) {
				// Retrieving the Path of PIM Module
				String locatorTypePim = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				String pathPim = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Opening PIM Web Page
				if (click(locatorTypePim, pathPim) == true) {

					// Retrieving the Path of Add Button
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					// Retrieving the Path of Login Credentials
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					// Locate the element you want to click
					WebElement elementToClick = driver.findElement(By.xpath(path));

					// Use WebDriverWait to wait for the element to be clickable
					WebElement clickableElement = wait.until(ExpectedConditions.elementToBeClickable(elementToClick));

					// Enabling creation with Login Credentials.
					clickableElement.click();

					// Retrieving Employees details from the Excel Sheet
					String firstName = getExcelData(newEmprow, firstNameColumnNumber, sheetName);
					String middleName = getExcelData(newEmprow, middleNameColumnNumber, sheetName);
					String lastName = getExcelData(newEmprow, lastNameColumnNumber, sheetName);
					String newUsername = getExcelData(newEmprow, newUsernameColumnNumber, sheetName);
					String newPassword = getExcelData(newEmprow, newPasswordColumnNumber, sheetName);

					// Entering New Employees Details

					// Retrieving the Path of First Name
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					enterText(locatorType, path, firstName);

					// Retrieving the Path of Middle Name
					locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

					enterText(locatorType, path, middleName);

					// Retrieving the Path of Last Name
					locatorType = getExcelData(rowNo6, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);

					enterText(locatorType, path, lastName);

					// Retrieving the Path of New Username
					locatorType = getExcelData(rowNo7, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo7, pathColumnNumber, resourceSheetName);

					enterText(locatorType, path, newUsername);

					// Retrieving the Path of New Password
					locatorType = getExcelData(rowNo8, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo8, pathColumnNumber, resourceSheetName);

					enterText(locatorType, path, newPassword);

					// Retrieving the Path of Confirm Password
					locatorType = getExcelData(rowNo9, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo9, pathColumnNumber, resourceSheetName);

					enterText(locatorType, path, newPassword);

					System.out.println(newUsername);
					System.out.println(newPassword);

					rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Id Value");
					rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName,
							"Username already exists");
					rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Submit");

					rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Successfully Saved");
					rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Dashboard");
					rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Current User Name");

					if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0 && rowNo6 > 0) {

						// Retrieving the Path of Employee Id Value
						locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

						String empId = getValue(locatorType, path);
						// Getting the Auto-generated Employee Id
						message("Employee Id - " + empId);

						// Retrieving the Path of Username already exists
						locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

						// Checking whether the username already exists or not.
						if (checkVisibility(locatorType, path) == false) {
							// If username is available creating the new user.
							// Retrieving the Path of Submit
							locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

							click(locatorType, path);

							// Waiting for the action to complete.
							wait = new WebDriverWait(driver, Duration.ofSeconds(10));

							// Retrieving the Path of Successfully Saved
							locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

							// Checking if Successfully Saved message visible or not
							if (checkVisibility(locatorType, path)) {

								// Updating the Employee ID in the Excel
//								updateExcelSheet(newEmprow, newEmployeeIdColumnNumber, sheetName, empId);

								if (validateEmployeeCreation(firstName, middleName, lastName, empId,
										"validateEmployeeCreation")) {

									// Logging out current user
									currentUserLogOut("currentUserLogOut", "User Log Out Successfully", creationTest);

									// Validating whether the new User can Log in or not.
									boolean check = adminLogin(newUsername, newPassword, "login", creationTest);
									if (check == true) {
										// Retrieving the Path of Dashboard
										locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
										path = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

										message("Page Name - " + getData(locatorType, path));

										// Retrieving the Path of Current User Name
										locatorType = getExcelData(rowNo6, locatorTypeColumnNumber, resourceSheetName);
										path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);
										message("Current User's Name - " + getData(locatorType, path));

										// Since, Employee Created Successfully and Updating the Status as Pass.
										updateExcelSheet(newEmprow, newEmployeeIdColumnNumber, sheetName, "Pass");
										checkResult(creationTest, expectedResult, "Creation Successful");

										currentUserLogOut("currentUserLogOut", "User Log Out Successfully",
												creationTest);

										adminLogin("Admin", "admin123", "login", creationTest);

										writeToTextFile(" } ");

									} else {
										System.out.println("Failed to Log in New User");
										checkResult(creationTest, expectedResult, "Failed to Log in New User");
										writeToTextFile(" } ");
									}
								} else {
									System.out.println("Employee Creation Failed");
									checkResult(creationTest, expectedResult, "Employee Created but not found");
								}
							} else {
								System.out.println("Employee Creation Failed");
								checkResult(creationTest, expectedResult, "Employee Creation Failed");
								writeToTextFile(" } ");
							}
						} else {
							System.out.println("Username Already Exists");
							checkResult(creationTest, expectedResult, "Creation Failed. Username Already Exists.");
							writeToTextFile(" } ");
						}

					} else {
						System.out.println("Xpath Not Found");
						checkResult(creationTest, expectedResult, "Xpath Not Found");
						writeToTextFile(" } ");
					}
				} else {
					System.out.println("You don't have access to create new employee");
					checkResult(creationTest, expectedResult,
							"You don't have access to create new employee or page is missing.");
					writeToTextFile(" } ");
				}

			} else {
				System.out.println("Xpath Not Found");
				checkResult(creationTest, expectedResult, "Xpath Not Found");
				writeToTextFile(" } ");
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			writeToTextFile(" } ");
		} catch (NoSuchElementException e) {
			writeToTextFile(" } ");
			e.printStackTrace();
		} catch (Exception e) {
			writeToTextFile(" } ");
			e.printStackTrace();
		}

	}

	// Method to Log Out Current User
	public static void currentUserLogOut(String methodName, String expectedResult, ExtentTest currentUserLogOutTest) {

		writeToTextFile("currentUserLogOut(), ");
		try {
			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Drop Down Menu");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Log Out Option");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Logout Check");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0) {
				// Retrieving the Path of Drop Down Menu
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Clicking on Drop Down Menu
				if (click(locatorType, path)) {

					// Retrieving the Path of Log Out Option
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					// Retrieving the Path of Logout Check
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					if (checkVisibility(locatorType, path)) {
						checkResult(currentUserLogOutTest, expectedResult, "User Log Out Successfully");
					} else {
						checkResult(currentUserLogOutTest, expectedResult, "Failed to Log out user");
					}
				} else {
					System.out.println("Unable to Locate Log Out Option.");
					checkResult(currentUserLogOutTest, expectedResult,
							"Unable to Locate Log Out Option. Failed to Log out user");
				}
			} else {
				System.out.println("Xpath Not Found");
				checkResult(currentUserLogOutTest, expectedResult, "Xpath Not Found");
			}
		} catch (IOException e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(currentUserLogOutTest, e);
		} catch (NoSuchElementException e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(currentUserLogOutTest, e);
		} catch (Exception e) {
			e.printStackTrace();
			reportFailCaseWitoutScreenshotException(currentUserLogOutTest, e);
		}
	}

	// Method to Go to Apply Leave Page
	public static void applyLeave(String methodName) {

		try {

			writeToTextFile("~applyLeave(" + methodName + ")# steps { ");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Leave Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Apply");

			if (rowNo >= 0 && rowNo2 > 0) {
				// Retrieving the Path of Leave Module
				locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Clicking on Leave Module
				if (click(locatorType, path)) {
					// Retrieving the Path of Apply
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

				}

			} else {
				System.out.println("Xpath Not Found");
			}
			writeToTextFile(" } ");
		} catch (IOException e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		} catch (NoSuchElementException e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		} catch (Exception e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		}

	}

	// Method to add New System User
	public static void addSystemUser(int i, int additionalInfoColumnNumber, String methodName) {

		try {
			writeToTextFile("~addSystemUser(" + i + "," + additionalInfoColumnNumber + "," + methodName + "), ");

			String userRole = getExcelData(i, additionalInfoColumnNumber, "Task");
			String employeeName = getExcelData(i, additionalInfoColumnNumber + 1, "Task");
			String status = getExcelData(i, additionalInfoColumnNumber + 2, "Task");
			String newUsername = getExcelData(i, additionalInfoColumnNumber + 3, "Task");
			String newPassword = getExcelData(i, additionalInfoColumnNumber + 4, "Task");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Admin Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Add");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "User Role");

			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "User Role Select");
			int rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Name");
			int rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Employee Name Select");
			int rowNo7 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Status");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0 && rowNo6 > 0 && rowNo7 > 0) {
				// Retrieving the Path of Admin Module
				String locatorTypeAdmin = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				String pathAdmin = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Switching to Admin Tab
				if (click(locatorTypeAdmin, pathAdmin)) {

					// Retrieving the Path of Add Button
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					// Retrieving the Path of User Role
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					// Retrieving the Path of User Role Select
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);
					path = replaceText(path, userRole);

					click(locatorType, path);

					// Retrieving the Path of Employee Name
					locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

					// Entering Employees Name
					enterText(locatorType, path, employeeName);

					// Retrieving the Path of Employee Name Select
					locatorType = getExcelData(rowNo6, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);
					path = replaceText(path, employeeName);

					if (checkVisibility(locatorType, path)) {
						System.out.println("Employee not found : " + employeeName);

					} else {

						click(locatorType, path);

						// Retrieving the Path of Status
						locatorType = getExcelData(rowNo7, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo7, pathColumnNumber, resourceSheetName);

						// Selecting the Status from the drop down menu
						click(locatorType, path);

						rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Status Option");
						rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Username");
						rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Username Check");

						rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Password");
						rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Confirm Password");
						rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Save");
						rowNo7 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Status");

						if (rowNo > 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0 && rowNo6 > 0
								&& rowNo7 > 0) {
							// Retrieving the Path of Status Option
							locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);
							path = replaceText(path, status);

							click(locatorType, path);

							// Retrieving the Path of Username
							locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

							// Entering the Username
							enterText(locatorType, path, newUsername);

							// Retrieving the Path of Username Check
							locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

							if (checkVisibility(locatorType, path)) {
								System.out.println("Username already exists : " + newUsername);

							} else {

								// Retrieving the Path of Password
								locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
								path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

								// Inserting the Passwords
								enterText(locatorType, path, newPassword);

								// Retrieving the Path of Confirm Password
								locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
								path = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

								enterText(locatorType, path, newPassword);

								// Retrieving the Path of Save
								locatorType = getExcelData(rowNo6, locatorTypeColumnNumber, resourceSheetName);
								path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);

								// Clicking the Save Button
								click(locatorType, path);

								// Retrieving the Path of Successfully Saved
								locatorType = getExcelData(rowNo7, locatorTypeColumnNumber, resourceSheetName);
								path = getExcelData(rowNo7, pathColumnNumber, resourceSheetName);

								// Waiting for Successful Message
								wait = new WebDriverWait(driver, Duration.ofSeconds(10));
								WebElement ele = driver.findElement(By.xpath(path));
								wait.until(ExpectedConditions.visibilityOf(ele));

								// Switching to Admin Tab
								click(locatorTypeAdmin, pathAdmin);

								System.out.println("System User Created Successfully");
							}
						} else {
							System.out.println("Xpath Not Found");
						}
					}
				}
			} else {
				System.out.println("Xpath Not Found");
			}

			writeToTextFile(" } ");
		} catch (IOException e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		} catch (NoSuchElementException e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		} catch (Exception e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		}
	}

	public static void disableSystemUser(String userName, String methodName) {

		try {

			writeToTextFile("~disableSystemUser(" + userName + "," + methodName + ")# steps { ");

			int rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Admin Module");
			int rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Username");
			int rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Search");

			int rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Record Found");
			int rowNo5 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "No Records Found");
			int rowNo6 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Disabled Check");
			int rowNo7 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Status");

			if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0 && rowNo5 > 0 && rowNo6 > 0 && rowNo7 > 0) {
				// Retrieving the Path of Admin Module
				String locatorTypeAdmin = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
				String pathAdmin = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

				// Switching to Admin Tab
				if (click(locatorTypeAdmin, pathAdmin)) {

					// Retrieving the Path of Username
					locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

					enterText(locatorType, path, userName);

					// Retrieving the Path of Search
					locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

					click(locatorType, path);

					// Retrieving the Path of Record Found and No Record Found
					locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

					locatorType = getExcelData(rowNo5, locatorTypeColumnNumber, resourceSheetName);
					String path2 = getExcelData(rowNo5, pathColumnNumber, resourceSheetName);

					wait = new WebDriverWait(driver, Duration.ofSeconds(10));
					wait.until(ExpectedConditions.or(ExpectedConditions.visibilityOfElementLocated(By.xpath(path)),
							ExpectedConditions.visibilityOfElementLocated(By.xpath(path2))));

					// Retrieving the Path of Disabled
					locatorType = getExcelData(rowNo6, locatorTypeColumnNumber, resourceSheetName);
					path = getExcelData(rowNo6, pathColumnNumber, resourceSheetName);

					if (checkVisibility(locatorType, path) == false) {
						System.out.println("User Already Disabled");
					} else {

						// Retrieving the Path of Status
						locatorType = getExcelData(rowNo7, locatorTypeColumnNumber, resourceSheetName);
						path = getExcelData(rowNo7, pathColumnNumber, resourceSheetName);

						click(locatorType, path);

						rowNo = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName,
								"Status Drop Down Menu");
						rowNo2 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Disabled");
						rowNo3 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName, "Save");
						rowNo4 = findTwoTextSameRowNumber(filePath, resourceSheetName, methodName,
								"Successfully Updated");

						if (rowNo >= 0 && rowNo2 > 0 && rowNo3 > 0 && rowNo4 > 0) {

							// Retrieving the Path of Status Drop Down Menu
							locatorType = getExcelData(rowNo, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo, pathColumnNumber, resourceSheetName);

							// Changing the Status to Disabled from the drop down menu
							click(locatorType, path);

							// Retrieving the Path of Disabled
							locatorType = getExcelData(rowNo2, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo2, pathColumnNumber, resourceSheetName);

							click(locatorType, path);

							// Retrieving the Path of Save
							locatorType = getExcelData(rowNo3, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo3, pathColumnNumber, resourceSheetName);

							// Clicking the Submit Button
							click(locatorType, path);

							// Retrieving the Path of Successfully Updated
							locatorType = getExcelData(rowNo4, locatorTypeColumnNumber, resourceSheetName);
							path = getExcelData(rowNo4, pathColumnNumber, resourceSheetName);

							// Waiting for Successful Message
							wait = new WebDriverWait(driver, Duration.ofSeconds(10));
							WebElement ele = driver.findElement(By.xpath(path));
							wait.until(ExpectedConditions.visibilityOf(ele));
						} else {
							System.out.println("Xpath Not Found");
						}

					}

				}
			} else {
				System.out.println("Xpath Not Found");
			}
			writeToTextFile(" } ");
		} catch (IOException e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		} catch (NoSuchElementException e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		} catch (Exception e) {
			e.printStackTrace();
			writeToTextFile(" } ");
		}

	}

	// Adding the Pass Case
	public static void reportPassCase(ExtentTest testCaseName, String message) {
		testCaseName.pass(message);
	}

	// Adding the Fail Case With Screenshots
	public static void reportFailCase(ExtentTest testCaseName, String message, String imageTitle) {
		writeToTextFile("reportFailCase(), ");
		String base64Code = captureScreenshot();
//		String path = captureScreenshot(imageTitle);
		testCaseName.fail(message,
				MediaEntityBuilder.createScreenCaptureFromBase64String(base64Code, imageTitle).build());
	}

	// Adding the Fail Case
	public static void reportFailCaseWitoutScreenshot(ExtentTest testCaseName, String message) {
		writeToTextFile("reportFailCaseWitoutScreenshot(), ");
		testCaseName.fail(message);
	}

	// Adding the Fail Case with Exception
	public static void reportFailCaseWitoutScreenshotException(ExtentTest testCaseName, Exception message) {
		writeToTextFile("reportFailCaseWitoutScreenshotException(), ");
		testCaseName.fail(message);
	}

	// Capturing the Image and Saving into the Local Folder
	public static String captureScreenshot(String fileName) {
		writeToTextFile("captureScreenshot(), ");
		TakesScreenshot takesScreenshot = (TakesScreenshot) driver;
		File sourceFile = takesScreenshot.getScreenshotAs(OutputType.FILE);
		File destFile = new File("./Screenshots/" + fileName);

		try {
			FileUtils.copyFile(sourceFile, destFile);
		} catch (IOException e) {
			e.printStackTrace();
		}

		System.out.println("Screenshots saved successfully");
		return destFile.getAbsolutePath();

	}

	// Capturing the Image and Saving into Base64Code
	public static String captureScreenshot() {
		writeToTextFile("captureScreenshot(), ");
		TakesScreenshot takesScreenshot = (TakesScreenshot) driver;
		String base64Code = takesScreenshot.getScreenshotAs(OutputType.BASE64);

		System.out.println("Screenshots saved successfully");
		return base64Code;

	}

}