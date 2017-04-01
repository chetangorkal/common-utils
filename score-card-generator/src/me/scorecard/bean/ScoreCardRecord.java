package me.scorecard.bean;

/**
 * Bean class to store details of a record read from excel report
 * 
 * @author Chetan Gorkal
 *
 */
public class ScoreCardRecord {

	private String topic;
	private double status;
	private String comment = "No Comments";

	public String getTopic() {
		return topic;
	}

	public void setTopic(String topic) {
		this.topic = topic;
	}

	public double getStatus() {
		return status;
	}

	public void setStatus(double status) {
		this.status = status;
	}

	public String getComment() {
		return comment;
	}

	public void setComment(String comment) {
		this.comment = comment;
	}

}
