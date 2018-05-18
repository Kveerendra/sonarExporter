package com.sonar.export;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.sonarqube.ws.Common.Severity;
import org.sonarqube.ws.Issues.Issue;
import org.sonarqube.ws.Issues.SearchWsResponse;
import org.sonarqube.ws.client.HttpConnector;
import org.sonarqube.ws.client.WsClient;
import org.sonarqube.ws.client.WsClientFactories;
import org.sonarqube.ws.client.issue.SearchWsRequest;


public class MAIN {

	static String LOGIN_USERNAME = "";
	static String LOGIN_PASSWORD = "";
	static String SONARQUBE_URL = "";
	static String OUTPUT_FILENAME = "SonarIssues.xlsx";
	static String SHEET_NAME = "Sonar issues";
	static String SHEET1_NAME = "Graph 1";
	static String SHEET2_NAME = "Graph 2";

	public static void init(String fileName) {
		Properties prop = new Properties();
		InputStream input = null;
		try {
			input = new FileInputStream(fileName);
			prop.load(input);
			LOGIN_USERNAME = prop.getProperty("username");
			LOGIN_PASSWORD = prop.getProperty("password");
			SONARQUBE_URL = prop.getProperty("sonarqubeurl");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static void main(String[] args) {
		if (args.length == 0) {
			System.err.println("Program needs Severities(BLOCKER CRITICAL MAJOR MINOR INFO) as Arguments");
			System.exit(0);
		}
		init("config.properties");
		System.out.println("This program will fetch maximum of 10000 issues.");
		HttpConnector httpConnector = HttpConnector.newBuilder().url(SONARQUBE_URL)
				.credentials(LOGIN_USERNAME, LOGIN_PASSWORD).build();
		WsClient wsClient = WsClientFactories.getDefault().newClient(httpConnector);
		SearchWsRequest swr = new SearchWsRequest();
		List<Issue> list = new ArrayList<>();
		swr.setPageSize(500);
		boolean flag = true;
		for (int i = 1; i <= 20 && flag; i++) {
			if (list.size() > 0 && list.get(list.size() - 1).getLine() == 0) {
				flag = false;
			}
			swr.setPage(i);
			swr.setSeverities(Arrays.asList(args));
			SearchWsResponse searchWsResponse = wsClient.issues().search(swr);
			list.addAll(searchWsResponse.getIssuesList());
		}
		deleteExistingFile();

		createExcel(list);
	}

	private static void deleteExistingFile() {
		File file = new File(OUTPUT_FILENAME);
		if (file.delete()) {
			System.err.println("File Deleted");
		}
	}

	private static void createExcel(List<Issue> issueList) {
		System.out.println("Generating Excel");
		try {
			String filename = OUTPUT_FILENAME;
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFCellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setBorderBottom(BorderStyle.THICK);
			headerStyle.setBorderTop(BorderStyle.THICK);
			headerStyle.setBorderLeft(BorderStyle.THICK);
			headerStyle.setBorderRight(BorderStyle.THICK);
			XSSFFont headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerStyle.setFont(headerFont);

			XSSFSheet sheet = workbook.createSheet(SHEET_NAME);
			XSSFRow rowhead = sheet.createRow((short) 0);
			XSSFCell cell = rowhead.createCell(0);
			cell.setCellValue("Project");
			cell.setCellStyle(headerStyle);
			cell = rowhead.createCell(1);
			cell.setCellValue("Track");
			cell.setCellStyle(headerStyle);
			cell = rowhead.createCell(2);
			cell.setCellValue("Path");
			cell.setCellStyle(headerStyle);
			cell = rowhead.createCell(3);
			cell.setCellValue("Line");
			cell.setCellStyle(headerStyle);
			cell = rowhead.createCell(4);
			cell.setCellValue("Severity");
			cell.setCellStyle(headerStyle);
			cell = rowhead.createCell(5);
			cell.setCellValue("Message");
			cell.setCellStyle(headerStyle);
			cell = rowhead.createCell(6);
			cell.setCellValue("Status");
			cell.setCellStyle(headerStyle);
			cell = rowhead.createCell(7);
			cell.setCellValue("Assignee");
			cell.setCellStyle(headerStyle);

			XSSFCellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setBorderBottom(BorderStyle.THIN);
			cellStyle.setBorderTop(BorderStyle.THIN);
			cellStyle.setBorderLeft(BorderStyle.THIN);
			cellStyle.setBorderRight(BorderStyle.THIN);
			int i = 0;
			for (i = 0; i < issueList.size() && issueList.get(i).getLine() != 0; i++) {
				XSSFRow row = sheet.createRow((short) i + 1);

				cell = row.createCell(0);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(issueList.get(i).getSubProject().split(":")[1]);
				cell = row.createCell(1);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(getTrack(issueList.get(i).getSubProject().split(":")[1]));
				cell = row.createCell(2);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(issueList.get(i).getComponent().split(":")[2]);
				cell = row.createCell(3);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(String.valueOf(issueList.get(i).getLine()));
				cell = row.createCell(4);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(getSeverity(issueList.get(i).getSeverity()));
				cell = row.createCell(5);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(issueList.get(i).getMessage());
				cell = row.createCell(6);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(issueList.get(i).getStatus());
				cell = row.createCell(7);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(issueList.get(i).getAssignee());
			}
			sheet.autoSizeColumn(0);
			sheet.autoSizeColumn(1);
			sheet.autoSizeColumn(2);
			sheet.autoSizeColumn(3);
			sheet.autoSizeColumn(4);
			sheet.autoSizeColumn(5);
			sheet.autoSizeColumn(6);
			sheet.autoSizeColumn(7);
			sheet.setDisplayGridlines(false);
			sheet.setAutoFilter(new CellRangeAddress(0, i, 0, 7));
			generatePivotSheet1(workbook);
			// generatePivotSheet2(workbook);
			FileOutputStream fileOut = new FileOutputStream(filename);
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
			System.out.printf("Your excel file(%s) has been generated!", OUTPUT_FILENAME);

		} catch (Exception ex) {
			ex.printStackTrace();
			System.out.println(ex);

		}
	}

	private static void generatePivotSheet1(XSSFWorkbook workbook) {
		System.out.println("Generating Pivot");
		XSSFSheet sheet = workbook.createSheet(SHEET1_NAME);
		XSSFSheet dataSheet = workbook.getSheet(SHEET_NAME);
		sheet.setDisplayGridlines(false);
		AreaReference pivotSource = new AreaReference("'Sonar issues'!A1:H" + dataSheet.getLastRowNum(),
				workbook.getSpreadsheetVersion());
		XSSFPivotTable pivotTable = sheet.createPivotTable(pivotSource, new CellReference(0, 0), dataSheet);
		pivotTable.addRowLabel(5);
		pivotTable.addRowLabel(1);
		long count = pivotTable.getCTPivotTableDefinition().getPivotFields().getCount();
		for(int j=0;j<count;j++)
		pivotTable.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(j).setOutline(false);

		pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, 3);
	}

	private static void generatePivotSheet2(XSSFWorkbook workbook) {
		XSSFSheet sheet = workbook.createSheet(SHEET2_NAME);
		XSSFSheet dataSheet = workbook.getSheet(SHEET_NAME);
		sheet.setDisplayGridlines(false);
		AreaReference pivotSource = new AreaReference("'Sonar issues'!A1:H" + (dataSheet.getLastRowNum() + 1),
				workbook.getSpreadsheetVersion());
		XSSFPivotTable pivotTable = sheet.createPivotTable(pivotSource, new CellReference(0, 0), dataSheet);
		pivotTable.addRowLabel(1);
		pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, 1);

	}

	private static String getSeverity(Severity severity) {
		switch (severity.ordinal()) {
		case Severity.BLOCKER_VALUE:
			return "Blocker";
		case Severity.CRITICAL_VALUE:
			return "Critical";
		case Severity.INFO_VALUE:
			return "Info";
		case Severity.MAJOR_VALUE:
			return "Major";
		case Severity.MINOR_VALUE:
			return "Minor";
		}
		return null;
	}

	private static String getTrack(String module) {
		switch (module) {
		case "AH":
		case "BI":
		case "BV":
		case "DM":
		case "TaskManagementEJB":
		case "IV":
			return "BO";
		case "CO":
			return "CO";
		case "CV":
			return "Conversion";
		case "ED":
		case "SH":
		case "AR":
			return "FO";
		case "RL":
			return "FO";
		case "IN":
			return "IN";
		case "RP":
			return "Reports";
		case "SE":
			return "Security";
		case "AL":
		case "FW":
		case "QC":
		case "RD":
		case "PM":
			return "Support";
		case "EDM":
		case "BatchOpsUtil":
		case "ST":
		case "RT":
		case "IEWebApp":
		case "Common":
		case "DA":
		case "HE":
		case "IQ":
		case "MO":
		case "CR":
		case "HM":
		case "HP":
		case "RM":
		case "TP":
		case "UI ":
		default:
			return "N/A";

		}
	}

}
