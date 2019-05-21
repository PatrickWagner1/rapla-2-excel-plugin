package semesterTimeTable.excel;

import java.util.Calendar;

/**
 * Class representing a lecture.
 */
public class Lecture {

	/** Name of the lecture. */
	private String name;
	
	/** Short name of the lecture */
	private String shortName;

	/** Start date and time of the lecture. */
	private Calendar startDate;

	/** End date and time of the lecture. */
	private Calendar endDate;

	/** The resources which the lecture uses. */
	private String[] resources;

	/** The lecturers giving the lecture. */
	private String[] lecturers;

	/**
	 * Constructor method taking multiple resources and lecturers.
	 * 
	 * @param name
	 * @param startDate
	 * @param endDate
	 * @param resources
	 * @param lecturers
	 */
	public Lecture(String name, Calendar startDate, Calendar endDate, String[] resources, String[] lecturers) {
		this.setName(name);
		this.setStartDate(startDate);
		this.setEndDate(endDate);
		this.setResources(resources);
		this.setLecturers(lecturers);
	}

	/**
	 * Constructor method taking multiple resources and a single lecturer.
	 * 
	 * @param name
	 * @param startDate
	 * @param endDate
	 * @param resources
	 * @param lecturer
	 */
	public Lecture(String name, Calendar startDate, Calendar endDate, String[] resources, String lecturer) {
		this.setName(name);
		this.setStartDate(startDate);
		this.setEndDate(endDate);
		this.setResources(resources);
		this.setLecturers(lecturer);
	}

	/**
	 * Constructor method taking a single resource and multiple lecturers.
	 * 
	 * @param name
	 * @param startDate
	 * @param endDate
	 * @param resources
	 * @param lecturer
	 */
	public Lecture(String name, Calendar startDate, Calendar endDate, String resource, String[] lecturers) {
		this.setName(name);
		this.setStartDate(startDate);
		this.setEndDate(endDate);
		this.setResources(resource);
		this.setLecturers(lecturers);
	}

	/**
	 * Constructor method taking a single resource and lecturer.
	 * 
	 * @param name
	 * @param startDate
	 * @param endDate
	 * @param resources
	 * @param lecturer
	 */
	public Lecture(String name, Calendar startDate, Calendar endDate, String resource, String lecturer) {
		this.setName(name);
		this.setStartDate(startDate);
		this.setEndDate(endDate);
		this.setResources(resource);
		this.setLecturers(lecturer);
	}

	/**
	 * Getter method for the name of the lecture.
	 * 
	 * @return Name of the current lecture object.
	 */
	public String getName() {
		return this.name;
	}

	/**
	 * Setter method for the name of the lecture.
	 * 
	 * @param name
	 */
	private void setName(String name) {
		this.name = name;
	}
	
	public String getShortName() {
		return this.shortName;
	}
	
	public void setShortName(String shortName) {
		this.shortName = shortName;
	}

	/**
	 * Getter method for the start date of the lecture.
	 * 
	 * @return Start date of the current lecture object.
	 */
	public Calendar getStartDate() {
		return this.startDate;
	}

	/**
	 * Setter method for the start date of the lecture.
	 * 
	 * @param startDate
	 */
	private void setStartDate(Calendar startDate) {
		this.startDate = startDate;
	}

	/**
	 * Getter method for the end date of the lecture
	 * 
	 * @return End date of the current lecture object.
	 */
	public Calendar getEndDate() {
		return this.endDate;
	}

	/**
	 * Setter method for the end date of the lecture.
	 * 
	 * @param endDate
	 */
	private void setEndDate(Calendar endDate) {
		this.endDate = endDate;
	}

	/**
	 * Getter method for multiple resources of the lecture.
	 * 
	 * @return The resources of the current lecture object as an array.
	 */
	public String[] getResources() {
		return this.resources;
	}

	/**
	 * Getter method for the single resource of the lecture.
	 * 
	 * @return The resource of the current lecture object.
	 */
	public String getResource() {
		return this.resources[0];
	}

	/**
	 * Setter method for multiple resources of the lecture.
	 * 
	 * @param resources
	 */
	private void setResources(String[] resources) {
		this.resources = resources;
	}

	/**
	 * Setter method for a single resource of the lecture.
	 * 
	 * @param resource
	 */
	private void setResources(String resource) {
		this.resources = new String[] { resource };
	}

	/**
	 * Getter method for multiple lecturers of the lecture.
	 * 
	 * @return The lecturers of the current lecture object as an array.
	 */
	public String[] getLecturers() {
		return this.lecturers;
	}

	/**
	 * Getter method for a single lecturer of the lecture.
	 * 
	 * @return The lecturer of the current lecture object.
	 */
	public String getLecturer() {
		return this.lecturers[0];
	}

	/**
	 * Setter method for multiple lecturers of the lecture.
	 * 
	 * @param lecturers
	 */
	private void setLecturers(String[] lecturers) {
		this.lecturers = lecturers;
	}

	/**
	 * Setter method for a single lecturer of the lecture.
	 * 
	 * @param lecturer
	 */
	private void setLecturers(String lecturer) {
		this.lecturers = new String[] { lecturer };
	}
}
