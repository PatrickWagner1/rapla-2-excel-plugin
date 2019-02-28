package semesterTimeTable.excel;

import java.util.Date;

public class Lecture {
	
	private String name;
	private Date startDate, endDate;
	private String[] resources, lecturers;

	public Lecture (String name, Date startDate, Date endDate, String[] resources, String[] lecturers) {
		this.name = name;
		this.startDate = startDate;
		this.endDate = endDate;
		this.resources = resources;
		this.lecturers = lecturers;
	}
	
	public Lecture (String name, Date startDate, Date endDate, String[] resources, String lecturer) {
		this.name = name;
		this.startDate = startDate;
		this.endDate = endDate;
		this.resources = resources;
		this.lecturers = new String[] {lecturer};
	}
	
	public Lecture (String name, Date startDate, Date endDate, String resource, String[] lecturers) {
		this.name = name;
		this.startDate = startDate;
		this.endDate = endDate;
		this.resources = new String[] {resource};
		this.lecturers = lecturers;
	}
	
	public Lecture (String name, Date startDate, Date endDate, String resource, String lecturer) {
		this.name = name;
		this.startDate = startDate;
		this.endDate = endDate;
		this.resources = new String[] {resource};
		this.lecturers = new String[] {lecturer};
	}

	public String getName() {
		return this.name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Date getStartDate() {
		return this.startDate;
	}

	public void setStartDate(Date startDate) {
		this.startDate = startDate;
	}

	public Date getEndDate() {
		return this.endDate;
	}

	public void setEndDate(Date endDate) {
		this.endDate = endDate;
	}

	public String[] getResources() {
		return this.resources;
	}
	
	public String getResource() {
		return this.resources[0];
	}

	public void setResources(String[] resources) {
		this.resources = resources;
	}
	
	public void setResources(String resource) {
		this.resources = new String[] {resource};
	}

	public String[] getLecturers() {
		return this.lecturers;
	}
	
	public String getLecturer() {
		return this.lecturers[0];
	}

	public void setLecturers(String[] lecturers) {
		this.lecturers = lecturers;
	}
	
	public void setLecturers(String lecturer) {
		this.lecturers = new String[] {lecturer};
	}
}
