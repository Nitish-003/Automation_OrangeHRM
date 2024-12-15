package orangeHRM;


import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.openqa.selenium.support.ui.WebDriverWait;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;


public class OrangeHRMMain extends UserDefinedMethods {

	static WebDriverWait wait;

	static String methodName;

	static ExtentReports extentReports;

	static String reportPath = "D:\\Eclipse\\eclipse-workspace\\afterBkp\\report";

	public static void main(String[] args) {

		extentReports = new ExtentReports();
		reportPath = reportPath + "_" + getCurrentTimestamp() + ".html";
		
		ExtentSparkReporter sparkReporter = new ExtentSparkReporter(reportPath);
		extentReports.attachReporter(sparkReporter);
		
		String testCaseName = "OrangeHRM";
		
		try {
			writeToTextFile(testCaseName + " | func");
			
			message(launchBrowser("Chrome"));
			fetchTasks();
			extentReports.flush();
			
//			Desktop.getDesktop().browse(new File(reportPath).toURI());
			
		} catch(Exception e) {
			System.out.println("Exception Caused in Fetch Task");
			e.printStackTrace();
		}
		System.out.println("Done.....");

		System.out.println("Execution Completed....Closing the Browser...Thank you");
		finishExecution();

	}
	
	public static String getCurrentTimestamp() {
        // Get the current timestamp
        LocalDateTime currentDateTime = LocalDateTime.now();

        // Define the desired format for the timestamp
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("ddMMyyyy_HHmmss");

        // Format the timestamp as a string
        String formattedTimestamp = currentDateTime.format(formatter);
        System.out.println(formattedTimestamp);

        return formattedTimestamp;
    }

	public static void fetchTasks() throws InterruptedException, IOException {

		writeToTextFile("~fetchTasks()#steps {\n");
		int lastTaskRowNumber = 1;
		int taskColumnNumber = 0;
		int additionalInfoColumnNumber = 0;

		try {
			// Retrieving the Last Row Number and the Column Number of the Task.
			lastTaskRowNumber = getLastRowNumber(filePath, "Task");
			taskColumnNumber = retrieveColumnNumber("Task", "Task");
			additionalInfoColumnNumber = retrieveColumnNumber("Task", "Additional Information");

			int i = 1;
			int flag = 0;
			while (i <= lastTaskRowNumber && flag == 0) {
				String tasks = getExcelData(i, taskColumnNumber, "Task");

				switch (tasks) {
				case "Login":

					// Getting the USer Choice whether to Log In as Admin, User or Validate Test
					// Case

					String addInfo = getExcelData(i, additionalInfoColumnNumber, "Task");
					if (addInfo.equals("Admin")) {
						// Logging In as a Admin
						ExtentTest loginTest = extentReports.createTest("Admin Login Test");
						boolean check = adminLogin("admin", "admin123", "login", loginTest);
						if (check == false) {
							flag = 1;

						}
					} else if (addInfo.equals("User")) {
						ExtentTest loginTest = extentReports.createTest("User Login Test");
						
						// Finding the Column Numbers
						try {
							// Row Number in Task Info
							int row = -1;
							try {
								row = Integer.parseInt(getExcelData(i, additionalInfoColumnNumber + 1, "Task"));
							
								float floatValue = Float
										.parseFloat(getExcelData(i, additionalInfoColumnNumber + 1, "Task"));
								row = (int) floatValue;
							} catch (NumberFormatException e) {
								reportFailCaseWitoutScreenshotException(loginTest, e);
							}
							
							// Performing Log In
							boolean check = login(row, "Sheet1", "login", loginTest);
							if (check == false)
								flag = 1;
						} catch (IOException e) {
							e.printStackTrace();
							reportFailCaseWitoutScreenshotException(loginTest, e);
							
						} catch (Exception e) {
							e.printStackTrace();
							reportFailCaseWitoutScreenshotException(loginTest, e);
							
						}

					} else if (addInfo.equals("Validate")) {
						
						ExtentTest loginTest = extentReports.createTest("Admin Login Test");

						// Performing all the Test Cases given in the Sheet
						int lastRowNo = 0;
						int currentRowNo = 1;
						int statusColumn = 0;

						// Finding the Column Numbers
						try {

							// Retrieving Column Number
//							int usernameColumn = retrieveColumnNumber("Sheet1", "Username");
//							int passwordColumn = retrieveColumnNumber("Sheet1", "Password");
							statusColumn = retrieveColumnNumber("Sheet1", "Status");
							currentRowNo = retrieveRowNumber("Sheet1", "Username") + 1;

							// Retrieving the Last Row Number from the Sheet
							lastRowNo = retrieveLastRow("Sheet1");
							System.out.println(lastRowNo);

							if (lastRowNo < currentRowNo) {
								System.out.println("No Test Data Found");
								reportPassCase(loginTest, "No Test Case Found.");
							} else {

								while (currentRowNo != lastRowNo + 1) {
									// Performing Log in function and storing the status.
									boolean check = login(currentRowNo, "Sheet1", "login", loginTest);

									if (check == true) {
										// Updating the Successful Login Status in the Excel Sheet
										updateExcelSheet(currentRowNo, statusColumn, "Sheet1", "Pass");

										// Logging Out user to check next test case
										currentUserLogOut("currentUserLogOut", "User Log Out Successfully", loginTest);

									} else {
										// Updating the Failed Login Status in the Excel Sheet
										updateExcelSheet(currentRowNo, statusColumn, "Sheet1", "Fail");
										

									}
									currentRowNo++;
								}
							}
						} catch (IOException e) {
							e.printStackTrace();
						}

					}
					break;

				case "Create Employee":
					// Creating a new PIM user
					int lastRowNo = 0;
					int currentRowNumber = 1;
					ExtentTest creationTest = extentReports.createTest("Create Employee Test");
					
					lastRowNo = retrieveLastRow("NewEmployee");

					if (lastRowNo < currentRowNumber) {
						System.out.println("No New Employee Data Found");
					} else {
						while (currentRowNumber != lastRowNo + 1) {
							createPIM(currentRowNumber, "NewEmployee", "createPIM", creationTest);
							currentRowNumber++;
						}
						System.out.println("All Employees Created Successfully");
					}
					break;

				case "Find Employee":
					// Retrieving the Employee Id to be searched
					ExtentTest searchEmployeeTest = extentReports.createTest("Create Employee Test");
					
					String empId = getExcelData(i, additionalInfoColumnNumber, "Task");
					searchEmployee(empId, "searchEmployee", "Employee Found Successfully", searchEmployeeTest);

					break;

				case "Delete Employee":

//					String deleteEmpId = getExcelData(i, additionalInfoColumnNumber, "Task");
					// Finding the Employee
//					if (searchEmployee(deleteEmpId, "searchEmployee")) {
//						// Deleting the Employee Based on Search Option
//						deleteEmployee(deleteEmpId, "deleteEmployee");
//					} else
//						System.out.println("You Don't have access to delete an Employee");

					break;

				case "Punch In":
					// Method to Punch In and Out
					ExtentTest punchInTest = extentReports.createTest("Punch In Test");
					
					lodgePunchInOut("lodgePunchInOut", "Punch In", punchInTest);
					break;

				case "Post News Feed":
					// Posting New News feed
					// Retrieving the text to be posted in the news feed.
					ExtentTest postNewsFeedTest = extentReports.createTest("Post News Feed Test");
					
					String postData = getExcelData(i, additionalInfoColumnNumber, "Task");
					postNewsfeed(postData, "postNewsfeed", "Newsfeed Posted Successfully", postNewsFeedTest);
					break;

				case "Delete News Feed":
					// Deleting News feed
					ExtentTest deleteNewsFeedTest = extentReports.createTest("Delete News Feed Test");
					
					deleteNewsfeed("deleteNewsfeed", deleteNewsFeedTest);
					break;

				case "Punch Out":
					// Method to Punch In and Out
					ExtentTest punchOutTest = extentReports.createTest("Punch Out Test");
					
					lodgePunchInOut("lodgePunchInOut", "Punch Out", punchOutTest);
					break;

				case "View Employee Time Sheet":
					// View Employee Time Sheet
					// Fetching the data from the excel sheet
					ExtentTest employeeTimeSheetTest = extentReports.createTest("View Employee Timesheet Test");
					
					String employeeName = getExcelData(i, additionalInfoColumnNumber, "Task");
					viewEmployeeTimeSheet(employeeName, "viewEmployeeTimeSheet", "Employee Timesheet is Visible", employeeTimeSheetTest);
					break;

				case "View My Time Sheet":
					// View My Time Sheet
					ExtentTest myTimeSheetTest = extentReports.createTest("View My Timesheet Test");
					
					viewMyTimeSheet("viewMyTimeSheet", "My Timesheet is Visible", myTimeSheetTest);
					break;

				case "Update Contact Details":
					// Method to Update Contact Details of Current Logged In User
					// Fetching the data from the excel sheet
					ExtentTest contactDetailsUpdateTest = extentReports.createTest("Contact Details Update Test");
					updateContactDetails(i, additionalInfoColumnNumber, "updateContactDetails", "Contact Details Updated Successfully", contactDetailsUpdateTest);
					break;

				case "Logout":
					// Method to Log Out Current User
					ExtentTest currentUserLogOutTest = extentReports.createTest("User Logout Test");
					
					currentUserLogOut("currentUserLogOut", "User Log Out Successfully", currentUserLogOutTest);
					break;

				case "Assign Leave":
					// Assigning Leave to the Employee
					// Fetching the data from the excel sheet
					ExtentTest assignLeaveTest = extentReports.createTest("Assign Leave Test");
					assignLeave(i, additionalInfoColumnNumber, "assignLeave",
							"Leave Assigned Successfully", assignLeaveTest);
					break;

				case "Add Vacancy":
					// Method to Add New Vacancy
					// Fetching the data from the excel sheet
					

					addVacancy(i, additionalInfoColumnNumber, "addVacancy");
					break;

				case "Apply Leave":
					// Method to Apply for Leave
					// Fetching the data from the excel sheet
					applyLeave("applyLeave");
					break;

				case "Add System User":
					// Method to Add New System User
					// Fetching the data from the excel sheet
					addSystemUser(i, additionalInfoColumnNumber, "addSystemUser");
					break;

				case "Disable System User":
					// Fetching the data from the excel sheet
					String usernameToDisable = getExcelData(i, additionalInfoColumnNumber, "Task");
					System.out.println("\n\nInside loop task is:- " + tasks);
					// Disable the System User Account
					disableSystemUser(usernameToDisable, "disableSystemUser");
					break;
				case "":
					break;

				default:
					System.out.println("Unsupported task: " + tasks);
				}
				i++;
			}
			writeToTextFile(" }");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}