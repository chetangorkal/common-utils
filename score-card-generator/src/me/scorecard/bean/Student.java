package me.scorecard.bean;

public class Student {

	private String fullName;
	private String usn;
	private String email;

	public String getFullName() {
		return fullName;
	}

	public void setFullName(String fullName) {
		this.fullName = fullName;
	}

	public String getUsn() {
		return usn;
	}

	public void setUsn(String usn) {
		this.usn = usn;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	@Override
	public String toString() {
		return "Full Name: " + this.fullName + " USN: " + this.usn + " e-mail: " + this.email;
	}
}
