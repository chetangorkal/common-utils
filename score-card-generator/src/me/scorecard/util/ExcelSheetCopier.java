package me.scorecard.util;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import me.scorecard.common.ScoreCardConstants;

public class ExcelSheetCopier {

	public static void main(String[] args) throws IOException, InvalidFormatException {
		ExcelSheetCopier copier = new ExcelSheetCopier();

		File directory = new File(ScoreCardConstants.SOURCE_FOLDER);
		String fileName = "1DS13TE059.xlsx";
		copier.process(directory.getAbsolutePath() + "\\" + fileName,
				new File(directory.getAbsolutePath() + "/test_bkp/"));

	}

	public void process(String absoluteFileName, File copyFolder) throws IOException, InvalidFormatException {
		BufferedInputStream bis = new BufferedInputStream(new FileInputStream(absoluteFileName));
		XSSFWorkbook sourceWorkbook = new XSSFWorkbook(bis);

		for (String copyFileName : copyFolder.list()) {
			System.out.println(copyFolder + "/" + copyFileName);

			OPCPackage pkg = OPCPackage.open(copyFolder + "/" + copyFileName);
			XSSFWorkbook destinationWorkBook = new XSSFWorkbook(pkg);

			// XSSFWorkbook destinationWorkBook = new XSSFWorkbook();
			for (int iSheet = 1; iSheet < sourceWorkbook.getNumberOfSheets(); iSheet++) {

				XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(iSheet);
				if (sourceSheet != null) {
					XSSFSheet destinationSheet = destinationWorkBook.createSheet(sourceSheet.getSheetName());

					int fRow = sourceSheet.getFirstRowNum();
					int lRow = sourceSheet.getLastRowNum();
					for (int iRow = fRow; iRow <= lRow; iRow++) {
						XSSFRow sourceRow = sourceSheet.getRow(iRow);
						if (sourceRow != null) {
							XSSFRow destinationRow = destinationSheet.createRow(iRow);
							int firstCellIndex = sourceRow.getFirstCellNum();
							int lastCellIndex = sourceRow.getLastCellNum();
							for (int iCell = firstCellIndex; iCell < lastCellIndex; iCell++) {
								XSSFCell sourceCell = sourceRow.getCell(iCell);
								XSSFCell destinationCell = destinationRow.createCell(iCell);
								if (sourceCell != null) {
									destinationCell.setCellType(sourceCell.getCellType());
									switch (sourceCell.getCellType()) {
									case Cell.CELL_TYPE_BLANK:
										destinationCell.setCellValue("");
										break;

									case Cell.CELL_TYPE_BOOLEAN:
										destinationCell.setCellValue(sourceCell.getBooleanCellValue());
										break;

									case Cell.CELL_TYPE_ERROR:
										destinationCell.setCellErrorValue(sourceCell.getErrorCellValue());
										break;

									case Cell.CELL_TYPE_FORMULA:
										destinationCell.setCellFormula(sourceCell.getCellFormula());
										break;

									case Cell.CELL_TYPE_NUMERIC:
										destinationCell.setCellValue(sourceCell.getNumericCellValue());
										break;

									case Cell.CELL_TYPE_STRING:
										destinationCell.setCellValue(sourceCell.getStringCellValue());
										break;
									default:
										destinationCell.setCellFormula(sourceCell.getCellFormula());
									}
								}
							}
						}
					}
				}

				bis.close();
				sourceWorkbook.close();
				FileOutputStream stream = new FileOutputStream(copyFolder + "/" + "temp-" + copyFileName);
				destinationWorkBook.write(stream);
				destinationWorkBook.close();
				File file = new File(copyFolder + "/" + "temp-" + copyFileName);
				file.delete();
			}
		}
	}
}