try {
			//Create File to be Saved
			FileOutputStream fileOut = new FileOutputStream("C:\\Users\\rpa15\\Documents\\AgentStub\\workbook.xls");
			
			//Create a workbook object, XSSF -> Office 2007 Later, HSSF -> Office 2007 Below
			Workbook workBook = new XSSFWorkbook();
			
			//Creating a new excel sheet named Employee Details
			Sheet sheet1 = workBook.createSheet("Employee Details");
			
			//Creating a new excel sheet named Salary Details
			Sheet sheet2 = workBook.createSheet("Salary Details");
			
			workBook.write(fileOut);
			workBook.close();
			fileOut.close();
		} catch (FileNotFoundException e) {
			System.out.println(e.getMessage());
		} catch (IOException e) {
			System.out.println(e.getMessage());
		}
