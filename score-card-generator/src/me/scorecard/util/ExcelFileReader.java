package me.scorecard.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import me.scorecard.bean.ScoreCardRecord;
import me.scorecard.bean.Student;

/**
 * Class to read excel sheet
 * 
 * @author Chetan Gorkal
 *
 */
public class ExcelFileReader {

	public List<ScoreCardRecord> readScoreCardExcelFile(String excelFilePath) throws IOException {
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		Workbook workbook = new XSSFWorkbook(inputStream);
		List<ScoreCardRecord> scoreCards = readScoreCardSheet(workbook.getSheetAt(0));
		workbook.close();
		inputStream.close();
		return scoreCards;
	}

	/**
	 * Method to read Students data from excel sheet.<br>
	 * Used by email service
	 * 
	 * @param excelFilePath
	 * @return
	 * @throws IOException
	 */
	public List<Student> readStudentExcelFile(String excelFilePath) throws IOException {
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		Workbook workbook = new XSSFWorkbook(inputStream);
		List<Student> scoreCards = readStudentSheet(workbook.getSheetAt(0));
		workbook.close();
		inputStream.close();
		return scoreCards;
	}

	private List<Student> readStudentSheet(Sheet studentSheet) {
		Iterator<Row> iterator = studentSheet.iterator();
		List<Student> studentList = new ArrayList<>();
		int cellCount;
		Student student = new Student();
		while (iterator.hasNext()) {

			cellCount = 1;

			student = new Student();
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					switch (cellCount) {
					case 1:
						student.setFullName(cell.getStringCellValue().trim());
						break;
					case 2:
						student.setUsn(cell.getStringCellValue().trim());
						break;
					case 3:
						student.setEmail(cell.getStringCellValue().trim());
						break;
					default:
						break;

					}
					break;
				default:
					break;
				}
				cellCount++;
			}
			studentList.add(student);
		}
		// System.out.println(studentList);
		return studentList;

	}

	public List<ScoreCardRecord> readScoreCardSheet(Sheet scoreCardSheet) {
		Iterator<Row> iterator = scoreCardSheet.iterator();
		List<ScoreCardRecord> scoreCardList = new ArrayList<>();
		int cellCount;
		ScoreCardRecord scoreCard = new ScoreCardRecord();
		while (iterator.hasNext()) {

			cellCount = 1;

			scoreCard = new ScoreCardRecord();
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					switch (cellCount) {
					case 1:
						scoreCard.setTopic(cell.getStringCellValue().trim());
						break;
					case 2:
						break;
					case 3:
						scoreCard.setComment(cell.getStringCellValue().trim());
						break;
					default:
						break;

					}
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					break;
				case Cell.CELL_TYPE_NUMERIC:
					switch (cellCount) {
					case 2:
						scoreCard.setStatus(cell.getNumericCellValue());
						break;
					}
					break;
				default:
					break;
				}
				cellCount++;
			}
			scoreCardList.add(scoreCard);
		}
		return scoreCardList;

	}
}