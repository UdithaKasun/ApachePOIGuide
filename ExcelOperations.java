package com.virtusa.rpa;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperations {

	public static void main(String[] args) {
	
	}
	
	//Create a new Chart in a New Sheet with reference Data
	public static String createChart(String filePath, String sourceSheet,String destinationSheet, int chartStartRowIndex, int chartAxisRowIndex, int chartEndRowIndex, String chartTitle, int dataXRowIndex, int dataXRowStartIndex, int dataXRowEndIndex,int dataYRowIndex, int dataYRowStartIndex, int dataYRowEndIndex){
		try {
			
			FileInputStream excelSource = new FileInputStream(new File(filePath));
            Workbook sourceWorkbook = new XSSFWorkbook(excelSource);
			Sheet dataSheet = sourceWorkbook.getSheet(sourceSheet);

			XSSFSheet chartSheet = null;
			if(sourceWorkbook.getSheet(destinationSheet) == null){
				
				chartSheet = (XSSFSheet) sourceWorkbook.createSheet(destinationSheet);
			}
			else{
				chartSheet = (XSSFSheet) sourceWorkbook.getSheet(destinationSheet);
				sourceWorkbook.removeSheetAt(sourceWorkbook.getSheetIndex(destinationSheet));
				chartSheet = (XSSFSheet) sourceWorkbook.createSheet(destinationSheet);
			}
			
			XSSFDrawing drawing =  chartSheet.getDrawingPatriarch();
			
			if(drawing == null){
				drawing = chartSheet.createDrawingPatriarch();
			}
						
			ClientAnchor anchor = drawing.createAnchor(1, 1, 1, 1, 1, chartStartRowIndex, chartAxisRowIndex, chartEndRowIndex);
			
			XSSFChart chart = (XSSFChart) drawing.createChart(anchor);
			ChartLegend legend = chart.getOrCreateLegend();
			
			legend.setPosition(LegendPosition.BOTTOM);
	        
	        LineChartData chartData = chart.getChartDataFactory().createLineChartData();
	        
	        ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
	        
	        
	        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
	        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
	        

	        ChartDataSource<Number> xValues= DataSources.fromNumericCellRange(dataSheet, new CellRangeAddress(dataXRowIndex, dataXRowIndex, dataXRowStartIndex, dataYRowEndIndex));
	        ChartDataSource<Number> yValues = DataSources.fromNumericCellRange(dataSheet, new CellRangeAddress(dataYRowIndex, dataYRowIndex, dataYRowStartIndex, dataYRowEndIndex));
	        System.out.println(yValues.getFormulaString());
	        chartData.addSeries(xValues, yValues);
	        chartData.getSeries().get(0).setTitle(chartTitle);
	       	chart.plot(chartData, bottomAxis, leftAxis);
	       	FileOutputStream fileOut = new FileOutputStream(new File(filePath));
	        sourceWorkbook.write(fileOut);
	        fileOut.close();
	        sourceWorkbook.close();
	        excelSource.close();
	       	return "OPERATION_SUCCESS";
		} catch (IOException e) {
			return "OPERATION_FAILED";
		}
		
	}
	
	public static String setColumnValues(String filePath, String sheetName, String[] sourceCells, String[] cellValues){
		try(FileInputStream excelSource = new FileInputStream(new File(filePath));
				Workbook sourceWorkbook = new XSSFWorkbook(excelSource);FileOutputStream fileOut = new FileOutputStream(filePath);){
				Sheet activeSheet = sourceWorkbook.getSheet(sheetName);
				int index = 0;
				for (String sourceCell : sourceCells) {
					CellReference referenceCell = new CellReference(sourceCell);
					Cell ouputCell = activeSheet.getRow(referenceCell.getRow()).getCell(referenceCell.getCol());
					writeCellValue(ouputCell, cellValues[index]);
					index++;
				}
				
				sourceWorkbook.write(fileOut);
				sourceWorkbook.close();
			    fileOut.close();
			    return "OPERATION_SUCCESS";
		} catch (IOException e) {
			return "OPERATION_FAILED";
		}
	}
	
	private static void writeCellValue(Cell sourceCell, String cellValue) {
		CellStyle currentCellStyle = sourceCell.getCellStyle();
		if (sourceCell != null) {
			if (sourceCell.getCellTypeEnum() == CellType.STRING) {
				sourceCell.setCellValue(cellValue);
				sourceCell.setCellStyle(currentCellStyle);
			} else if (sourceCell.getCellTypeEnum() == CellType.NUMERIC) {
				if (DateUtil.isCellDateFormatted(sourceCell)) {
					sourceCell.setCellValue(new Date(cellValue));
					sourceCell.setCellStyle(currentCellStyle);
				} else {
					sourceCell.setCellValue(Double.parseDouble(cellValue));
					sourceCell.setCellStyle(currentCellStyle);
				}
			} else if (sourceCell.getCellTypeEnum() == CellType.BOOLEAN) {
				sourceCell.setCellValue(Boolean.parseBoolean(cellValue));
				sourceCell.setCellStyle(currentCellStyle);
			} else if (sourceCell.getCellTypeEnum() == CellType.FORMULA) {
				sourceCell.setCellFormula(cellValue);
				sourceCell.setCellStyle(currentCellStyle);
			} else if (sourceCell.getCellTypeEnum() == CellType._NONE) {
				sourceCell.setCellValue("");
				sourceCell.setCellStyle(currentCellStyle);
			} else if (sourceCell.getCellTypeEnum() == CellType.BLANK) {
				sourceCell.setCellValue("");
				sourceCell.setCellStyle(currentCellStyle);
			}
		} else {
			throw new NullPointerException();
		}
	}
	
	//Return cell address RowIndex, ColumnIndex
	public static String getCellAddressColumnAndRowIndex(String filePath, String sheetName, String cellAddress){
		Workbook sourceWorkbook;
		try {
			sourceWorkbook = ExcelOperations.getWorkbookInstance(filePath);
			Sheet sourceSheet = ExcelOperations.getSheetInstance(sourceWorkbook, sheetName);
			CellReference referenceCell = new CellReference(cellAddress);
			Cell referedCell = sourceSheet.getRow(referenceCell.getRow()).getCell(referenceCell.getCol());
			if(referedCell != null){
				return referenceCell.getRow() + "," + referenceCell.getCol();
			}
			else{
				return "CELL_NOT_FOUND";
			}
		} catch (FileNotFoundException e) {
			return "CELL_NOT_FOUND";
		}
		
		
	}
	
	//Copy paster cells to adjecent cells
	public static String copyPasteColumnsAdjecentCell(String filePath, String sheetName, String sourceStartCellAddress ,  int copyCellCount, String dateToBeUsed){
		try {
			FileInputStream excelSource = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(excelSource);
			Sheet workingSheet = workbook.getSheet(sheetName);
			
			Iterator<Row> rowIterator = workingSheet.iterator();

			ArrayList<Row> rowList = new ArrayList<Row>();
			while (rowIterator.hasNext()) {
				rowList.add(rowIterator.next());
			}
			
			
			CellReference sourceStartCell = new CellReference(sourceStartCellAddress);
			
			List<Row> filteredList = rowList.stream()
					.filter( row -> { 
						return (row.getRowNum() >= sourceStartCell.getRow() && row.getRowNum() <= (sourceStartCell.getRow() + (copyCellCount-1)));
					})
					.collect(Collectors.toList());
			
			for (Row row : filteredList) {
	        	   Cell cellSource = row.getCell(sourceStartCell.getCol());
	        	   if(cellSource != null)
	        	   {
	        		   Cell newCell;
	        		   Row destinationRow = workingSheet.getRow(cellSource.getRowIndex());
	        		   int columnIndex = cellSource.getColumnIndex() + 1;
	        		   switch (cellSource.getCellType()) {
		                case 1:
		                	newCell = destinationRow.createCell(columnIndex);
	                    	newCell.setCellValue(row.getCell(sourceStartCell.getCol()).getStringCellValue());
	                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());	   	
		                    break;
		                case 0:
		                    if (DateUtil.isCellDateFormatted(row.getCell(sourceStartCell.getCol()))) {
		                    	
		                    	newCell = destinationRow.createCell(columnIndex);
		                    	newCell.setCellValue(new Date(dateToBeUsed));
		                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());		                    	
		                    } else {
		                    	newCell = destinationRow.createCell(columnIndex);
		                    	newCell.setCellValue(row.getCell(sourceStartCell.getCol()).getNumericCellValue());
		                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());		
		                    }
		                    break;
		                case 4:
		                	newCell = destinationRow.createCell(columnIndex);
	                    	newCell.setCellValue(row.getCell(sourceStartCell.getCol()).getBooleanCellValue());
	                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());		
		                    break;
		                case 2:
		                   	newCell = destinationRow.createCell(columnIndex);
		                   	String previousCellAddress = row.getCell(sourceStartCell.getCol()).getAddress().formatAsString().replaceAll("\\d", "");
		                   	String newCellAddress = newCell.getAddress().formatAsString().replaceAll("\\d", "");
		                   	String newCellFormula = row.getCell(sourceStartCell.getCol()).getCellFormula().replaceAll(previousCellAddress, newCellAddress);
	                    	newCell.setCellFormula(newCellFormula);
	                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());
		                    break;
		                case 3:
		                	newCell = destinationRow.createCell(columnIndex);
		                	newCell.setCellValue("");
	                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());
		                    break;
		                default:
		                    System.out.println();
		            }
	        	   }
			}
			
			FileOutputStream fileOut = new FileOutputStream(filePath);
	        workbook.write(fileOut);
	        workbook.close();
	        fileOut.close();
	        return "OPERATION_SUCCESS";
		} catch (FileNotFoundException e) {
			return "OPERATION_FAILED";
		} catch (IOException e) {
			return "OPERATION_FAILED";
		}
		
	}
	
	@SuppressWarnings("deprecation")
	public static void copyPasteColumns(String filePath, String sheetName, String sourceStartCellAddress ,  int copyCellCount, int destinationCellRowIndex, int destinationCellColIndex,String dateToBeUsed){
		try {
			FileInputStream excelSource = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(excelSource);
			Sheet workingSheet = workbook.getSheet(sheetName);
			
			Iterator<Row> rowIterator = workingSheet.iterator();

			ArrayList<Row> rowList = new ArrayList<Row>();
			while (rowIterator.hasNext()) {
				rowList.add(rowIterator.next());
			}
			
			
			CellReference sourceStartCell = new CellReference(sourceStartCellAddress);
			
			List<Row> filteredList = rowList.stream()
					.filter( row -> { 
						return (row.getRowNum() >= sourceStartCell.getRow() && row.getRowNum() <= (sourceStartCell.getRow() + copyCellCount));
					})
					.collect(Collectors.toList());
			
			for (Row row : filteredList) {
	        	   Cell cellSource = row.getCell(sourceStartCell.getCol());
	        	   if(cellSource != null)
	        	   {
	        		   Cell newCell;
	        		   Row destinationRow = workingSheet.getRow(destinationCellRowIndex++);
	        		   switch (cellSource.getCellType()) {
		                case 1:
		                	newCell = destinationRow.createCell(destinationCellColIndex);
	                    	newCell.setCellValue(row.getCell(sourceStartCell.getCol()).getStringCellValue());
	                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());	   	
		                    break;
		                case 0:
		                    if (DateUtil.isCellDateFormatted(row.getCell(sourceStartCell.getCol()))) {
		                    	
		                    	newCell = destinationRow.createCell(destinationCellColIndex);
		                    	newCell.setCellValue(new Date());
		                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());		                    	
		                    } else {
		                    	newCell = destinationRow.createCell(destinationCellColIndex);
		                    	newCell.setCellValue(row.getCell(sourceStartCell.getCol()).getNumericCellValue());
		                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());		
		                    }
		                    break;
		                case 4:
		                	newCell = destinationRow.createCell(destinationCellColIndex);
	                    	newCell.setCellValue(row.getCell(sourceStartCell.getCol()).getBooleanCellValue());
	                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());		
		                    break;
		                case 2:
		                   	newCell = destinationRow.createCell(destinationCellColIndex);
		                   	String previousCellAddress = row.getCell(sourceStartCell.getCol()).getAddress().formatAsString().replaceAll("\\d", "");
		                   	String newCellAddress = newCell.getAddress().formatAsString().replaceAll("\\d", "");
		                   	String newCellFormula = row.getCell(sourceStartCell.getCol()).getCellFormula().replaceAll(previousCellAddress, newCellAddress);
	                    	newCell.setCellFormula(newCellFormula);
	                    	newCell.setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());
		                    break;
		                case 3:
		                	destinationRow.createCell(destinationCellColIndex).setCellValue("");
		                	destinationRow.createCell(destinationCellColIndex).setCellStyle(row.getCell(sourceStartCell.getCol()).getCellStyle());
		                    break;
		                default:
		                    System.out.println();
		            }
	        	   }
			}
			
			FileOutputStream fileOut = new FileOutputStream(filePath);
	        workbook.write(fileOut);
	        workbook.close();
	        fileOut.close();
		} catch (FileNotFoundException e) {
			System.out.println(e.getMessage());
		} catch (IOException e) {
			System.out.println(e.getMessage());
		}
	}
	
	@SuppressWarnings("deprecation")
	public static void copyPasteColumns(String filePath, String sheetName, int sourceCellRowIndex, int sourceCellColIndex , int rowCopyCount, int destinationCellRowIndex, int destinationCellColIndex,String dateToBeUsed){
		try {
			FileInputStream excelSource = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(excelSource);
			Sheet workingSheet = workbook.getSheet(sheetName);

			Iterator<Row> rowIterator = workingSheet.iterator();

			ArrayList<Row> rowList = new ArrayList<Row>();
			while (rowIterator.hasNext()) {
				rowList.add(rowIterator.next());
			}
			
			List<Row> filteredList = rowList.stream()
					.filter( row -> { 
						return (row.getRowNum() >= sourceCellRowIndex && row.getRowNum() <= (sourceCellRowIndex + rowCopyCount));
					})
					.collect(Collectors.toList());
			
			for (Row row : filteredList) {
	        	   Cell cellSource = row.getCell(sourceCellColIndex);
	        	   if(cellSource != null)
	        	   {
	        		   Cell newCell;
	        		   Row destinationRow = workingSheet.getRow(destinationCellRowIndex++);
	        		   switch (cellSource.getCellType()) {
		                case 1:
		                	newCell = destinationRow.createCell(destinationCellColIndex);
	                    	newCell.setCellValue(row.getCell(sourceCellColIndex).getStringCellValue());
	                    	newCell.setCellStyle(row.getCell(sourceCellColIndex).getCellStyle());	   	
		                    break;
		                case 0:
		                    if (DateUtil.isCellDateFormatted(row.getCell(sourceCellColIndex))) {
		                    	
		                    	newCell = destinationRow.createCell(destinationCellColIndex);
		                    	newCell.setCellValue(new Date());
		                    	newCell.setCellStyle(row.getCell(sourceCellColIndex).getCellStyle());		                    	
		                    } else {
		                    	newCell = destinationRow.createCell(destinationCellColIndex);
		                    	newCell.setCellValue(row.getCell(sourceCellColIndex).getNumericCellValue());
		                    	newCell.setCellStyle(row.getCell(sourceCellColIndex).getCellStyle());		
		                    }
		                    break;
		                case 4:
		                	newCell = destinationRow.createCell(destinationCellColIndex);
	                    	newCell.setCellValue(row.getCell(sourceCellColIndex).getBooleanCellValue());
	                    	newCell.setCellStyle(row.getCell(sourceCellColIndex).getCellStyle());		
		                    break;
		                case 2:
		                   	newCell = destinationRow.createCell(destinationCellColIndex);
		                   	String previousCellAddress = row.getCell(sourceCellColIndex).getAddress().formatAsString().replaceAll("\\d", "");
		                   	String newCellAddress = newCell.getAddress().formatAsString().replaceAll("\\d", "");
		                   	String newCellFormula = row.getCell(sourceCellColIndex).getCellFormula().replaceAll(previousCellAddress, newCellAddress);
	                    	newCell.setCellFormula(newCellFormula);
	                    	newCell.setCellStyle(row.getCell(sourceCellColIndex).getCellStyle());
		                    break;
		                case 3:
		                	destinationRow.createCell(destinationCellColIndex).setCellValue("");
		                	destinationRow.createCell(destinationCellColIndex).setCellStyle(row.getCell(sourceCellColIndex).getCellStyle());
		                    break;
		                default:
		                    System.out.println();
		            }
	        	   }
			}
			
			FileOutputStream fileOut = new FileOutputStream(filePath);
	        workbook.write(fileOut);
	        workbook.close();
	        fileOut.close();
		} catch (FileNotFoundException e) {
			System.out.println(e.getMessage());
		} catch (IOException e) {
			System.out.println(e.getMessage());
		}
	}
	
	//Given Row Keyword and Row Reference Cell Search the given row for the keyword and return Cell Address
	public static String rowBasedExcelSearch(String filePath, String sheetName,String rowKeyword, String filterRowAddress){
		CellReference filterRowCell = new CellReference(filterRowAddress);
		return rowBasedExcelSearch(filePath, sheetName, rowKeyword, filterRowCell.getRow());
	}
	
	//Given Row Keyword and Row Index Search the given row for the keyword and return Cell Address
	public static String rowBasedExcelSearch(String filePath, String sheetName,String rowKeyword, int filterRowIndex){
		try (Workbook sourceWorkbook = ExcelOperations.getWorkbookInstance(filePath)){
			Sheet workingSheet = ExcelOperations.getSheetInstance(sourceWorkbook, sheetName);
			Row filteredRow = workingSheet.getRow(filterRowIndex);
			
			for (Cell cell : filteredRow) {
				String cellValue = formatCellAsString(cell);
				if(cellValue.trim().equals(rowKeyword.trim())){
					return cell.getAddress().formatAsString();
				}	
			}
			
			return "CELL_NOT_FOUND";
			
		} catch (IOException e) {
			return "CELL_NOT_FOUND";
		}
		
	}
	// Search for a keyword row based given array of keywords and get the
	// relevant values in the resulted
	// rows in the given Column : Address Based
	public static String[] getAddressesSearchRowColumnKeywordsBased(String filePath, String sheetName,
			String rowKeywords[], String startRowCellAddress, String endRowCellAddress, String filterRowAddress,
			String filterColumnAddress) {
		ArrayList<String> searchResult = new ArrayList<>();
		for (String rowKeyword : rowKeywords) {
			String foundCellAddress = searchRowColumnKeywordBasedAddress(filePath, sheetName, rowKeyword,
					startRowCellAddress, endRowCellAddress, filterRowAddress, filterColumnAddress);
			searchResult.add(foundCellAddress);
		}
		return searchResult.toArray(new String[0]);
	}

	// Search for a keyword row based given array of keywords and get the
	// relevant values in the resulted
	// rows in the given Column : Address Based
	public static String[] getValuesSearchRowColumnKeywordsBased(String filePath, String sheetName,
			String rowKeywords[], String startRowCellAddress, String endRowCellAddress, String filterRowAddress,
			String filterColumnAddress) {
		ArrayList<String> searchResult = new ArrayList<>();
		for (String rowKeyword : rowKeywords) {
			String foundCellValue = searchRowColumnKeywordBased(filePath, sheetName, rowKeyword, startRowCellAddress,
					endRowCellAddress, filterRowAddress, filterColumnAddress);
			searchResult.add(foundCellValue);
		}
		return searchResult.toArray(new String[0]);
	}

	// Search for a keyword row based and get the relevant address in the
	// resulted
	// row in the given Column : Address Based
	public static String searchRowColumnKeywordBasedAddress(String filePath, String sheetName, String rowkeyword,
			String startRowCellAddress, String endRowCellAddress, String filterRowAddress, String filterColumnAddress) {

		CellReference filterColumnCell = new CellReference(filterColumnAddress);
		String searchResult = "";
		String foundCellAddress = crossSearchExcel(filePath, sheetName, rowkeyword, startRowCellAddress,
				endRowCellAddress, filterRowAddress);
		if (!foundCellAddress.equals("FILE_NOT_FOUND") && !foundCellAddress.equals("CELL_NOT_FOUND")
				&& !foundCellAddress.equals("OPERATION_FAILED")) {
			try (Workbook sourceWorkbook = ExcelOperations.getWorkbookInstance(filePath)) {
				Sheet workingSheet = ExcelOperations.getSheetInstance(sourceWorkbook, sheetName);
				CellReference referedCell = new CellReference(foundCellAddress);
				searchResult = workingSheet.getRow(referedCell.getRow()).getCell(filterColumnCell.getCol()).getAddress()
						.formatAsString();
			} catch (FileNotFoundException e) {
				return "OPERATION_FAILED";
			} catch (NullPointerException e) {
				return "OPERATION_FAILED";
			} catch (IOException e) {
				return "OPERATION_FAILED";
			}
		} else {
			return foundCellAddress;
		}
		return searchResult;
	}

	// Search for a keyword row based and get the relevant value in the resulted
	// row in the given Column : Address Based
	public static String searchRowColumnKeywordBased(String filePath, String sheetName, String rowkeyword,
			String startRowCellAddress, String endRowCellAddress, String filterRowAddress, String filterColumnAddress) {

		CellReference filterColumnCell = new CellReference(filterColumnAddress);
		String searchResult = "";
		String foundCellAddress = crossSearchExcel(filePath, sheetName, rowkeyword, startRowCellAddress,
				endRowCellAddress, filterRowAddress);
		if (!foundCellAddress.equals("FILE_NOT_FOUND") && !foundCellAddress.equals("CELL_NOT_FOUND")
				&& !foundCellAddress.equals("OPERATION_FAILED")) {
			try (Workbook sourceWorkbook = ExcelOperations.getWorkbookInstance(filePath)) {
				Sheet workingSheet = ExcelOperations.getSheetInstance(sourceWorkbook, sheetName);
				CellReference referedCell = new CellReference(foundCellAddress);
				searchResult = formatCellAsString(
						workingSheet.getRow(referedCell.getRow()).getCell(filterColumnCell.getCol()));
			} catch (FileNotFoundException e) {
				return "OPERATION_FAILED";
			} catch (NullPointerException e) {
				return "OPERATION_FAILED";
			} catch (IOException e) {
				return "OPERATION_FAILED";
			}
		} else {
			return foundCellAddress;
		}
		return searchResult;
	}

	// Search for a keyword row based and get the relevant value in the resulted
	// row in the given Column : Index Based
	public static String searchRowColumnKeywordBased(String filePath, String sheetName, String rowkeyword,
			int startRowIndex, int endRowIndex, int filterCell, int getColumnIndex) {
		String searchResult = "";
		String foundCellAddress = crossSearchExcel(filePath, sheetName, rowkeyword, startRowIndex, endRowIndex,
				filterCell);
		if (!foundCellAddress.equals("FILE_NOT_FOUND") && !foundCellAddress.equals("CELL_NOT_FOUND")
				&& !foundCellAddress.equals("OPERATION_FAILED")) {
			try (Workbook sourceWorkbook = ExcelOperations.getWorkbookInstance(filePath)) {
				Sheet workingSheet = ExcelOperations.getSheetInstance(sourceWorkbook, sheetName);
				CellReference referedCell = new CellReference(foundCellAddress);
				searchResult = formatCellAsString(workingSheet.getRow(referedCell.getRow()).getCell(getColumnIndex));
			} catch (FileNotFoundException e) {
				return "OPERATION_FAILED";
			} catch (NullPointerException e) {
				return "OPERATION_FAILED";
			} catch (IOException e) {
				return "OPERATION_FAILED";
			}
		} else {
			return foundCellAddress;
		}
		return searchResult;
	}

	// Search a Column for a specific keywords given as a row keywords Row Range
	// as Cell Address
	// and return Cell Address of results
	public static String[] crossSearchExcel(String filePath, String sheetName, String[] rowKeywords,
			String startRowCellAddress, String endRowCellAddress, String filterRowAddress) {

		CellReference startRowCell = new CellReference(startRowCellAddress);
		CellReference endRowCell = new CellReference(endRowCellAddress);
		CellReference filterCell = new CellReference(filterRowAddress);

		ArrayList<String> addressList = new ArrayList<>();

		for (String rowKeyword : rowKeywords) {
			addressList.add(crossSearchExcel(filePath, sheetName, rowKeyword, startRowCell.getRow(),
					endRowCell.getRow(), filterCell.getCol()));
		}

		return addressList.toArray(new String[0]);
	}

	// Search a Column for a specific keyword given a Row Range as Cell Address
	// and return Cell Address of result
	public static String crossSearchExcel(String filePath, String sheetName, String rowkeyword,
			String startRowCellAddress, String endRowCellAddress, String filterRowAddress) {

		CellReference startRowCell = new CellReference(startRowCellAddress);
		CellReference endRowCell = new CellReference(endRowCellAddress);
		CellReference filterCell = new CellReference(filterRowAddress);

		return crossSearchExcel(filePath, sheetName, rowkeyword, startRowCell.getRow(), endRowCell.getRow(),
				filterCell.getCol());
	}

	// Search a Column for a specific keyword given a Row Range as indexes and
	// return Cell Address of result
	public static String crossSearchExcel(String filePath, String sheetName, String rowkeyword, int startRowIndex,
			int endRowIndex, int filterRowIndex) {

		try (Workbook sourceWorkbook = ExcelOperations.getWorkbookInstance(filePath)) {
			Sheet workingSheet = ExcelOperations.getSheetInstance(sourceWorkbook, sheetName);

			Iterator<Row> rowIterator = workingSheet.iterator();
			List<Row> rowList = new ArrayList<Row>();

			while (rowIterator.hasNext()) {
				rowList.add(rowIterator.next());
			}

			Stream<Row> rowStream = rowList.stream();

			List<Row> filteredRowList = rowStream.filter(row -> {
				return (row.getRowNum() >= startRowIndex && row.getRowNum() <= endRowIndex);
			}).collect(Collectors.toList());

			for (Row filteredRow : filteredRowList) {
				for (Cell cell : filteredRow) {
					if (ExcelOperations.formatCellAsString(cell).trim().equals(rowkeyword.trim())
							&& cell.getColumnIndex() == filterRowIndex) {
						return cell.getAddress().formatAsString();
					}
				}
			}

		} catch (FileNotFoundException e) {
			return "OPERATION_FAILED";
		} catch (NullPointerException e) {
			return "OPERATION_FAILED";
		} catch (IOException e1) {
			return "OPERATION_FAILED";
		}

		return "OPERATION_FAILED";
	}

	// Get Cell Value Based on Given Cell Address
	public static String getCellValueAsString(String filePath, String sheetName, String cellAddress) {
		try (Workbook sourceWorkbook = ExcelOperations.getWorkbookInstance(filePath)) {
			Sheet workingSheet = ExcelOperations.getSheetInstance(sourceWorkbook, sheetName);
			CellReference cellReference = new CellReference(cellAddress);
			Cell referedCell = workingSheet.getRow(cellReference.getRow()).getCell(cellReference.getCol());
			String referedCellValue = ExcelOperations.formatCellAsString(referedCell);
			return referedCellValue;
		} catch (FileNotFoundException e) {
			return "OPERATION_FAILED";
		} catch (NullPointerException e) {
			return "OPERATION_FAILED";
		} catch (IOException e1) {
			return "OPERATION_FAILED";
		}
	}

	// Given a rowIndex and columnIndex return Cell Value as a String
	public static String getCellValueAsString(String filePath, String sheetName, int rowIndex, int columnIndex) {
		CellReference referedCell = new CellReference(rowIndex, columnIndex);
		return getCellValueAsString(filePath, sheetName, referedCell.formatAsString());
	}

	// Get Workbook Instance
	private static Workbook getWorkbookInstance(String filePath) throws FileNotFoundException {
		try (FileInputStream workbookSource = new FileInputStream(new File(filePath));
				Workbook workbook = new XSSFWorkbook(workbookSource);) {
			return workbook;
		} catch (FileNotFoundException e) {
			throw new FileNotFoundException();
		} catch (IOException e) {
			return null;
		}
	}

	// Get Sheet Instance
	private static Sheet getSheetInstance(Workbook sourceWorkbook, String sheetName) {
		Sheet selectedSheet = sourceWorkbook.getSheet(sheetName);
		if (selectedSheet == null) {
			throw new NullPointerException();
		}
		return selectedSheet;
	}

	// Given a cell address format the cell value as String
	private static String formatCellAsString(Cell sourceCell) {

		String formatedCellValue = "";
		if (sourceCell != null) {
			if (sourceCell.getCellTypeEnum() == CellType.STRING) {
				formatedCellValue = sourceCell.getStringCellValue();
			} else if (sourceCell.getCellTypeEnum() == CellType.NUMERIC) {
				if (DateUtil.isCellDateFormatted(sourceCell)) {
					formatedCellValue = new SimpleDateFormat("MM/dd/yyyy").format(sourceCell.getDateCellValue());
				} else {
					DecimalFormat df = new DecimalFormat("0");
					df.setMaximumFractionDigits(2);
					formatedCellValue = df.format(sourceCell.getNumericCellValue());
				}
			} else if (sourceCell.getCellTypeEnum() == CellType.BOOLEAN) {
				formatedCellValue = String.valueOf(sourceCell.getBooleanCellValue());
			} else if (sourceCell.getCellTypeEnum() == CellType.FORMULA) {
				DecimalFormat df = new DecimalFormat("0");
				df.setMaximumFractionDigits(2);
				formatedCellValue = df.format(sourceCell.getNumericCellValue());
			} else if (sourceCell.getCellTypeEnum() == CellType._NONE) {
				formatedCellValue = "";
			} else if (sourceCell.getCellTypeEnum() == CellType.BLANK) {
				formatedCellValue = "";
			}

			return formatedCellValue;
		} else {
			throw new NullPointerException();
		}
	}
	

}
