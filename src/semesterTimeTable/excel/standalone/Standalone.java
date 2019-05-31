package semesterTimeTable.excel.standalone;

import java.awt.FileDialog;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;
import java.util.Map.Entry;

import semesterTimeTable.excel.Lecture;
import semesterTimeTable.excel.LectureWorkbook;

public class Standalone {

	/** CSV cell break character */
	private static final String CELL_BREAK = ";";

	/** File extension for excel files */
	private static final String FILE_EXTENSION = "xlsx";

	/** Default lecture file name */
	private static final String LECTURE_FILE_TITLE = "Vorlesungsplan";

	/** Default time zone */
	private static final TimeZone TIME_ZONE = TimeZone.getTimeZone("GMT");

	/** Column position of lecture name in CSV file */
	private static final int LECTURE_NAME_POSITION = 0;
	/** Column position of start date in CSV file */
	private static final int LECTURE_START_DATE_POSITION = 1;
	/** Column position of end date in CSV file */
	private static final int LECTURE_END_DATE_POSITION = 2;
	/** Column position of resources in CSV file */
	private static final int LECTURE_RESOURCES_POSITION = 3;
	/** Column position of lecturers in CSV file */
	private static final int LECTURE_LECTURERS_POSITION = 4;

	/** Frame for stand alone GUI */
	private StandaloneFrame standaloneFrame;

	/**
	 * Opens the rapla to excel converter as stand alone GUI.
	 */
	public Standalone() {
		this.standaloneFrame = new StandaloneFrame();
	}

	/**
	 * Returns the stand alone frame.
	 * 
	 * @return The stand alone frame.
	 */
	private StandaloneFrame getStandaloneFrame() {
		return this.standaloneFrame;
	}

	/**
	 * Method implementing the actual exporting functionality.
	 * 
	 * @throws IOException If reading the CSV file or reading/writing the excel file
	 *                     failed
	 */
	public void export() throws IOException {
		StandaloneFrame standaloneFrame = this.getStandaloneFrame();
		String csvFileName = this.loadFile();
		if (csvFileName != null && csvFileName != "") {
			standaloneFrame.println("Scanning CSV File...");
			List<String[]> rawLectureList = Standalone.CSVToTwoDimList(csvFileName, Standalone.CELL_BREAK, true);
			String mostCommonClassName = this.getMostCommonClassName(rawLectureList);
			String filename = this.getDefaultFileName(mostCommonClassName);
			standaloneFrame.println("Waiting for excel file selection...");
			String path = this.saveFile(filename);
			if (path != null) {
				standaloneFrame.println("Converting CSV file to excel file...");
				LectureWorkbook lectureWorkbook = new LectureWorkbook(path, standaloneFrame);

				List<Lecture> lectures = this.getLecturesFromRawLectureList(rawLectureList);

				if ((lectureWorkbook.getQuarterStartDate() == null || lectureWorkbook.getQuarterEndDate() == null)
						&& lectures.size() > 0) {
					Calendar dateInQuarter = (Calendar) lectures.get(0).getStartDate().clone();
					lectureWorkbook.setBorderDatesWithDateInQuarter(dateInQuarter);
				}

				lectureWorkbook.setLectures(lectures);
				lectureWorkbook.saveToFile(path);
				String errorOutput = lectureWorkbook.getErrorOutput().getErrorOutput();
				if (errorOutput == null || errorOutput == "") {
					standaloneFrame.println("Export completed successfully!");
					int secondsToWait = 2;
					try {
						Thread.sleep(secondsToWait * 1000);
					} catch (InterruptedException e) {
						System.err.println("Waiting for" + secondsToWait + "seconds failed");
					}
					standaloneFrame.close();
				} else {
					standaloneFrame.println("Export completed with some failures!");
				}
			}
		}
	}

	/**
	 * Open a FileDialog for loading a file and return the full path of the entered
	 * filename. If the FileDialog was closed or canceled, then null is returned.
	 * 
	 * @return The full path of the entered filename, or null if closed or canceled
	 *         the FileDialog
	 */
	public String loadFile() {
		final StandaloneFrame frame = this.getStandaloneFrame();
		final FileDialog fd = new FileDialog(frame, "Load CSV File", FileDialog.LOAD);

		fd.setLocation(50, 50);
		fd.setVisible(true);
		final String loadFileName = fd.getFile();

		String path = null;
		if (loadFileName != null) {
			path = this.createFullPath(fd);
		}

		return path;
	}

	/**
	 * Open a FileDialog for saving a file and return the full path of the entered
	 * filename. If the FileDialog was closed or canceled, then null is returned.
	 * 
	 * @param filename The default filename for the FileDialog
	 * @return The full path of the entered filename, or null if closed or canceled
	 *         the FileDialog
	 */
	public String saveFile(String filename) {
		final StandaloneFrame frame = this.getStandaloneFrame();
		final FileDialog fd = new FileDialog(frame, "Save (and load) Excel File", FileDialog.SAVE);

		fd.setFile(filename);

		fd.setLocation(50, 50);
		fd.setVisible(true);
		final String savedFileName = fd.getFile();

		String path = null;
		if (savedFileName != null) {
			path = this.createFullPath(fd);
		}

		return path;
	}

	/**
	 * Returns the full path of a file dialog.
	 * 
	 * @param fd The file dialog
	 * @return The full path of the file dialog
	 */
	private String createFullPath(final FileDialog fd) {
		String filename = fd.getFile();
		return fd.getDirectory() + filename;
	}

	/**
	 * Returns the default file name for lectures.
	 * 
	 * @return The default file name
	 */
	private String getDefaultFileName(String className) {
		String lecturesTitle = Standalone.LECTURE_FILE_TITLE;
		if (className != null && className != "") {
			lecturesTitle += "_" + className;
		}

		return lecturesTitle + "." + Standalone.FILE_EXTENSION;
	}

	/**
	 * Converts a raw list of lectures into a list of lectures.
	 * 
	 * @param rawLectureList the raw list of lectures
	 * @return A list of lectures
	 */
	private List<Lecture> getLecturesFromRawLectureList(List<String[]> rawLectureList) {

		TimeZone timeZone = Standalone.TIME_ZONE;
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		sdf.setTimeZone(timeZone);

		List<Lecture> lectures = new ArrayList<Lecture>();

		for (String[] row : rawLectureList) {
			String lectureName = row[Standalone.LECTURE_NAME_POSITION];

			Calendar lectureStartDate = null;
			Calendar startDate = new GregorianCalendar();
			startDate.setTimeZone(timeZone);
			try {
				startDate.setTime(sdf.parse(row[Standalone.LECTURE_START_DATE_POSITION]));
				lectureStartDate = startDate;
			} catch (ParseException e) {
				System.err.println("Cannot parse start date of the lecture \"" + lectureName + "\"");
			}

			Calendar lectureEndDate = null;
			Calendar endDate = new GregorianCalendar();
			endDate.setTimeZone(timeZone);
			try {
				endDate.setTime(sdf.parse(row[Standalone.LECTURE_END_DATE_POSITION]));
				lectureEndDate = endDate;
			} catch (ParseException e) {
				System.err.println("Cannot parse end date of the lecture \"" + lectureName + "\"");
			}

			String[] lectureResources = null;
			String resourcesString = row[Standalone.LECTURE_RESOURCES_POSITION];
			if (resourcesString != null) {
				String[] resources = resourcesString.split(", ");
				List<String> rooms = new ArrayList<String>();
				for (String resource : resources) {
					if (Standalone.resourceIsRoom(resource)) {
						rooms.add(resource);
					}
				}
				lectureResources = rooms.toArray(new String[rooms.size()]);
			}

			String[] lectureLecturers = null;
			String lecturers = row[Standalone.LECTURE_LECTURERS_POSITION];
			if (lecturers != null) {
				lectureLecturers = lecturers.split(", ");
			}

			Lecture lecture = new Lecture(lectureName, lectureStartDate, lectureEndDate, lectureResources,
					lectureLecturers);
			lectures.add(lecture);
		}

		return lectures;
	}

	/**
	 * Returns the most common class name from a raw list of lectures.
	 * 
	 * @param rawLectureList The raw list of lectures
	 * @return The most common class name from the raw lecture list
	 */
	private String getMostCommonClassName(List<String[]> rawLectureList) {

		Map<String, Integer> classNames = new HashMap<String, Integer>();

		for (String[] row : rawLectureList) {
			String element = row[Standalone.LECTURE_RESOURCES_POSITION];
			if (element != null) {
				String[] resources = element.split(", ");
				for (String resource : resources) {
					if (!Standalone.resourceIsRoom(resource)) {
						resource = resource.replaceAll(" \\(.*\\)", "");
						int count = 0;
						if (classNames.containsKey(resource)) {
							count = classNames.get(resource);
						}
						count++;
						classNames.put(resource, count);
					}
				}
			}
		}
		return getHighestCountKey(classNames);
	}

	/**
	 * Converts a CSV file to a two dimensional raw list of lectures.
	 * 
	 * @param filename     The filename of the CSV file
	 * @param separator    The separator for columns in the CSV file
	 * @param hasTitleLine True if the first line of the CSV file contains the
	 *                     column titles, otherwise false.
	 * @return A raw list of lectures
	 * @throws IOException If reading CSV file failed
	 */
	public static List<String[]> CSVToTwoDimList(String filename, String separator, boolean hasTitleLine)
			throws IOException {
		BufferedReader bufferedReader = null;
		String line;

		List<String[]> list = new ArrayList<String[]>();

		try {
			bufferedReader = new BufferedReader(new FileReader(filename));
			if (hasTitleLine) {
				bufferedReader.readLine();
			}
			while ((line = bufferedReader.readLine()) != null) {
				String[] rawElements = line.split(separator);
				String[] elements = new String[5];
				int maxSize = Math.min(rawElements.length, elements.length);
				for (int i = 0; i < maxSize; i++) {
					elements[i] = rawElements[i];
				}
				list.add(elements);
			}
		} finally {
			if (bufferedReader != null) {
				bufferedReader.close();
			}
		}
		return list;
	}

	/**
	 * Checks if the given resource name matches the regular expression for a room.
	 * If the name does NOT contains 3 upper case letters followed by two digits.
	 * 
	 * @param resourceName The name to check if it is a room name
	 * @return True if the given name is a room name, otherwise false.
	 */
	public static boolean resourceIsRoom(String resourceName) {
		return !resourceName.matches(".*\\p{Upper}{3}\\d{2}.*");
	}

	/**
	 * Returns the key with the highest value number in the given map.
	 * 
	 * @param map The map for searching the highest number
	 * @return The key of the highest value number
	 */
	private static String getHighestCountKey(Map<String, Integer> map) {
		Entry<String, Integer> highestEntry = null;
		for (Entry<String, Integer> entry : map.entrySet()) {
			if (highestEntry == null || entry.getValue() > highestEntry.getValue()) {
				highestEntry = entry;
			}
		}
		return highestEntry == null ? null : highestEntry.getKey();
	}

	/**
	 * Starts the stand alone GUI of the rapla 2 excel converter.
	 * 
	 * @param args The arguments are not used
	 */
	public static void main(String[] args) {
		Standalone standalone = new Standalone();
		try {
			standalone.export();
		} catch (Exception e) {
			standalone.getStandaloneFrame().println(e.toString());
		}
	}
}
