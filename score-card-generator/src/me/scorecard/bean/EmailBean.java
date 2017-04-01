package me.scorecard.bean;

public class EmailBean {

	private Student student;
	private String emailSubject;
	private String emailMessage;

	public EmailBean() {

	}

	public EmailBean(String emailSubject, String emailMessage) {
		super();
		this.emailSubject = emailSubject;
		this.emailMessage = emailMessage;
	}

	public Student getStudent() {
		return student;
	}

	public void setStudent(Student student) {
		this.student = student;
	}

	public String getEmailSubject() {
		return emailSubject;
	}

	public void setEmailSubject(String emailSubject) {
		this.emailSubject = emailSubject;
	}

	public String getEmailMessage() {
		return emailMessage;
	}

	public void setEmailMessage(String emailMessage) {
		this.emailMessage = emailMessage;
	}

}
