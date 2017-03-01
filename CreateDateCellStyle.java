try {
			//Create File to be Saved
			FileOutputStream fileOut = new FileOutputStream("C:\\Users\\rpa15\\Documents\\AgentStub\\workbook.xls");
			
			//Create a workbook object, XSSF -> Office 2007 Later, HSSF -> Office 2007 Below
			Workbook workBook = new XSSFWorkbook();
			
			//Creating a workbookCreationHelper
			CreationHelper apacheHelper = workBook.getCreationHelper(); 
					
			//Creating a new excel sheet named Employee Details
			Sheet sheet1 = workBook.createSheet("Employee Details");
			
			//Creating a new excel sheet named Salary Details
			Sheet sheet2 = workBook.createSheet("Salary Details");
			
			//Creating a Row
			Row createdRow = sheet2.createRow(0);
			
			//Creating a Cell
			Cell createdCell_1 = createdRow.createCell(0);
			Cell createdCell_2 = createdRow.createCell(1);
			Cell dateCell = createdRow.createCell(2);
			
			//Creating a Cell Style
			
			CellStyle dateCellStyle = workBook.createCellStyle();
			dateCellStyle.setDataFormat(apacheHelper.createDataFormat().getFormat("m/d/yy h:mm"));
			
			//Setting Values in Created Cells
			createdCell_1.setCellValue("Cell One");
			createdCell_2.setCellValue("Cell Two");
			dateCell.setCellValue(new Date());
			dateCell.setCellStyle(dateCellStyle);
			
			workBook.write(fileOut);
			workBook.close();
			fileOut.close();
		} catch (FileNotFoundException e) {
			System.out.println(e.getMessage());
		} catch (IOException e) {
			System.out.println(e.getMessage());
		}
