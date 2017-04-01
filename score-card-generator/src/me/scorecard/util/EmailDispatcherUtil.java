package me.scorecard.util;

import java.io.IOException;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import me.scorecard.bean.Student;
import me.scorecard.common.ScoreCardConstants;

public class EmailDispatcherUtil {

	public void send() throws IOException, MessagingException {

		List<Student> studentList = readStudentDataFromExcelFile();
		for (Student student : studentList) {
			Message mimeMessage = prepareEmail(student);
			// sends the e-mail
			System.out.println("Sending report to: " + student.getFullName());
			Transport.send(mimeMessage);
			System.out.println("Report has been successfully send to: " + student.getFullName());
		}
	}

	private Message prepareEmail(Student student) throws MessagingException, IOException {
		Properties smtpProperties = getSmtpProperties();

		Session session = Session.getDefaultInstance(smtpProperties, new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(ScoreCardConstants.SENDER_EMAIL, ScoreCardConstants.SENDER_PASSWORD);
			}
		});
		Multipart multipart = prepareMultipartEmail(student);
		Message mimeMessage = composeMessage(session, ScoreCardConstants.SENDER_EMAIL, student.getEmail(),
				ScoreCardConstants.EMAIL_SUBJECT);
		mimeMessage.setContent(multipart);

		return mimeMessage;
	}

	private static List<Student> readStudentDataFromExcelFile() throws IOException {
		return new ExcelFileReader().readStudentExcelFile(ScoreCardConstants.USER_DATA_FILE);
	}

	private Multipart prepareMultipartEmail(Student student) throws MessagingException, IOException {

		MimeBodyPart messageBodyPart = prepareMessagePart(student);
		MimeBodyPart attachPart = prepareAttachmentPart(student);
		Multipart multipart = new MimeMultipart();
		multipart.addBodyPart(messageBodyPart);
		multipart.addBodyPart(attachPart);
		System.out.println("-------------------------------------------------------------------------------");
		System.out.println("MessageBodyPart:\n" + messageBodyPart.getContent());
		System.out.println("*******");
		System.out.println("MessageBodyPart:\n" + attachPart.getFileName());
		System.out.println("-------------------------------------------------------------------------------");

		// printEmailDetailsToConsole(attachPart, attachPart);
		return multipart;
	}

	private void printEmailDetailsToConsole(MimeBodyPart messageBodyPart, MimeBodyPart attachPart)
			throws IOException, MessagingException {
		System.out.println("-------------------------------------------------------------------------------");
		System.out.println("MessageBodyPart:\n" + messageBodyPart.getContent());
		System.out.println("*******");
		System.out.println("MessageBodyPart:\n" + attachPart.getFileName());
		System.out.println("-------------------------------------------------------------------------------");

	}

	/**
	 * Method to create body part for the attachment
	 * 
	 * @param student
	 * 
	 * @return {@link MimeBodyPart} with report as attachment
	 */
	private MimeBodyPart prepareAttachmentPart(Student student) {
		MimeBodyPart attachPart = new MimeBodyPart();
		String attachmentFilePath = ScoreCardConstants.PDF_DESTINATION_FOLDER + student.getUsn().toUpperCase() + ".pdf";
		try {
			attachPart.attachFile(attachmentFilePath);
		} catch (Exception e) {
			System.out.println("Exception occurred while attaching the file: " + attachmentFilePath);
			e.printStackTrace();
		}
		return attachPart;
	}

	/**
	 * Method to create email body
	 * 
	 * @param student
	 * 
	 * @return {@link MimeBodyPart} containing email body
	 * @throws MessagingException
	 */
	private MimeBodyPart prepareMessagePart(Student student) throws MessagingException {
		MimeBodyPart messageBodyPart = new MimeBodyPart();

		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.append("Dear " + student.getFullName() + ",<br/>");
		stringBuilder.append(ScoreCardConstants.EMAIL_HTML_TEXT);
		stringBuilder.append("Thank you,<br/>" + "Chetan");
		messageBodyPart.setContent(stringBuilder.toString(), "text/html");
		return messageBodyPart;
	}

	private static Message composeMessage(Session session, String senderEmail, String receiverEmail,
			String emailSubject) throws AddressException, MessagingException {
		Message mimeMessage = new MimeMessage(session);
		mimeMessage.setFrom(new InternetAddress(senderEmail));
		InternetAddress[] toAddresses = { new InternetAddress(receiverEmail) };
		mimeMessage.setRecipients(Message.RecipientType.TO, toAddresses);
		mimeMessage.setSubject(emailSubject);
		mimeMessage.setSentDate(new Date());
		return mimeMessage;

	}

	private static Properties getSmtpProperties() {
		Properties smtpProperties = new Properties();
		smtpProperties.put("mail.smtp.host", "smtp.gmail.com");
		smtpProperties.put("mail.smtp.socketFactory.port", "465");
		smtpProperties.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
		smtpProperties.put("mail.smtp.auth", "true");
		smtpProperties.put("mail.smtp.port", "465");
		smtpProperties.put("mail.smtp.starttls.enable", "true");
		smtpProperties.put("mail.smtp.socketFactory.fallback", "false");
		smtpProperties.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
		return smtpProperties;
	}
}
