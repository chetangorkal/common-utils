package me.scorecard.client;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.mail.MessagingException;

import me.scorecard.bean.ScoreCardRecord;
import me.scorecard.common.ScoreCardConstants;
import me.scorecard.util.EmailDispatcherUtil;
import me.scorecard.util.ExcelFileReader;
import me.scorecard.util.ScoreCardExcelToPdfConvertor;

public class ScoreCardGenerator {

	public static void main(String[] args) throws IOException, MessagingException {
		ScoreCardGenerator scoreCardGenerator = new ScoreCardGenerator();
		// scoreCardGenerator.createScoreCards();
		// scoreCardGenerator.convertScoreCardsFromExcelToPDF();
		scoreCardGenerator.dispatchEmails();

	}

	/**
	 * Method to generate copies of feedback reports.<br>
	 * Note: This method should be run only once. Other wise, marks will be over
	 * written
	 * 
	 * @throws IOException
	 */
	@SuppressWarnings("unused")
	private void createScoreCards() throws IOException {
		int count = 2;
		while (count < 10) {
			copyFile(ScoreCardConstants.SOURCE_FOLDER + File.separator + "1DS14TE001.xlsx",
					ScoreCardConstants.DESTINATION_FOLDER + "/1DS14TE00" + count + ".xlsx");
			count++;
		}

		while (count < 100) {
			copyFile(ScoreCardConstants.SOURCE_FOLDER + File.separator + "1DS14TE001.xlsx",
					ScoreCardConstants.DESTINATION_FOLDER + "/1DS14TE0" + count + ".xlsx");
			count++;
		}
	}

	private void convertScoreCardsFromExcelToPDF() throws IOException {
		ScoreCardExcelToPdfConvertor excelToPdfConvertor = new ScoreCardExcelToPdfConvertor();
		ExcelFileReader scoreCardExcelFileReader = new ExcelFileReader();
		File directory = new File(ScoreCardConstants.PDF_SOURCE_FOLDER);
		java.util.List<ScoreCardRecord> scoreCardsList = null;
		if (directory.isDirectory()) {
			for (String fileName : directory.list()) {
				scoreCardsList = scoreCardExcelFileReader
						.readScoreCardExcelFile(directory.getAbsolutePath() + "\\" + fileName);
				excelToPdfConvertor.generatePDFReports(scoreCardsList, fileName.substring(0, 10) + ".pdf");
			}
		}
	}

	private void dispatchEmails() throws IOException, MessagingException {
		new EmailDispatcherUtil().send();

	}

	private void copyFile(String sourceFile, String destinationFile) throws IOException {
		File inputFile = new File(sourceFile);
		FileInputStream fileInputStream = new FileInputStream(inputFile);
		FileOutputStream fileOutputStream = new FileOutputStream(destinationFile);
		int c;
		while ((c = fileInputStream.read()) != -1) {
			fileOutputStream.write(c);
		}
		fileOutputStream.close();
		fileInputStream.close();
	}

}
