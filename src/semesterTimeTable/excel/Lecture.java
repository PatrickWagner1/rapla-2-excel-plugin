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
	 * @return The name of the lecture
	 */
	public String getName() {
		return this.name;
	}

	/**
	 * Setter method for the name of the lecture.
	 * 
	 * @param The name for the lecture
	 */
	private void setName(String name) {
		this.name = name;
	}

	/**
	 * Setter method for the short name of the lecture.
	 * 
	 * @return The short name of the lecture
	 */
	public String getShortName() {
		return this.shortName;
	}

	/**
	 * Getter method for the short name of the lecture.
	 * 
	 * @param shortName The short name for the lecture
	 */
	public void setShortName(String shortName) {
		this.shortName = shortName;
	}

	/**
	 * Getter method for the start date of the lecture.
	 * 
	 * @return The start date of the lecture
	 */
	public Calendar getStartDate() {
		return this.startDate;
	}

	/**
	 * Setter method for the start date of the lecture.
	 * 
	 * @param startDate The start date for the lecture
	 */
	private void setStartDate(Calendar startDate) {
		this.startDate = startDate;
	}

	/**
	 * Getter method for the end date of the lecture
	 * 
	 * @return The end date of the lecture
	 */
	public Calendar getEndDate() {
		return this.endDate;
	}

	/**
	 * Setter method for the end date of the lecture.
	 * 
	 * @param endDate The end date for the lecture
	 */
	private void setEndDate(Calendar endDate) {
		this.endDate = endDate;
	}

	/**
	 * Getter method for multiple resources of the lecture.
	 * 
	 * @return The resources of the lectures
	 */
	public String[] getResources() {
		return this.resources;
	}

	/**
	 * Getter method for the single resource of the lecture.
	 * 
	 * @return The (first) resource of the lectures
	 */
	public String getResource() {
		return this.resources[0];
	}

	/**
	 * Setter method for multiple resources of the lecture.
	 * 
	 * @param resources The resources for the lecture
	 */
	private void setResources(String[] resources) {
		this.resources = resources;
	}

	/**
	 * Setter method for a single resource of the lecture.
	 * 
	 * @param resource The resource for the lecture
	 */
	private void setResources(String resource) {
		this.resources = new String[] { resource };
	}

	/**
	 * Getter method for multiple lecturers of the lecture.
	 * 
	 * @return The lecturers of the lecture
	 */
	public String[] getLecturers() {
		return this.lecturers;
	}

	/**
	 * Getter method for a single lecturer of the lecture.
	 * 
	 * @return The (first) lecturer of the lecture
	 */
	public String getLecturer() {
		return this.lecturers[0];
	}

	/**
	 * Setter method for multiple lecturers of the lecture.
	 * 
	 * @param lecturers The lecturers for the lecture
	 */
	private void setLecturers(String[] lecturers) {
		this.lecturers = lecturers;
	}

	/**
	 * Setter method for a single lecturer of the lecture.
	 * 
	 * @param lecturer The lecturer for the lecture
	 */
	private void setLecturers(String lecturer) {
		this.lecturers = new String[] { lecturer };
	}
}
