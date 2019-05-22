package semesterTimeTable.excel;

import java.awt.FileDialog;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;
import java.util.Map.Entry;

import org.rapla.framework.RaplaContextException;
import org.rapla.framework.RaplaException;
import org.rapla.plugin.export2excel.Export2ExcelMenu;

public class Standalone {

	private static final String CELL_BREAK = ";";

	private static final String FILE_EXTENSION = "xlsx";

	private static final String LECTURE_FILE_TITLE = "Vorlesungsplan";

	private static final TimeZone TIME_ZONE = TimeZone.getTimeZone("GMT");

	private static final int LECTURE_NAME_POSITION = 0;
	private static final int LECTURE_START_DATE_POSITION = 1;
	private static final int LECTURE_END_DATE_POSITION = 2;
	private static final int LECTURE_RESSOURCES_POSITION = 3;
	private static final int LECTURE_LECTURERS_POSITION = 4;

	private StandaloneFrame standaloneFrame;

	public Standalone() {
		this.standaloneFrame = new StandaloneFrame();
	}

	private StandaloneFrame getStandaloneFrame() {
		return this.standaloneFrame;
	}

	/**
	 * Method implementing the actual exporting functionality. It is called by the
	 * event handler for clicking on the export menu entry.
	 * 
	 * @param model The calendar selection model
	 * @throws Exception
	 */
	public void export() throws Exception {
		StandaloneFrame standaloneFrame = this.getStandaloneFrame();
		String csvFileName = this.loadFile();
		if (csvFileName != null && csvFileName != "") {
			standaloneFrame.addTextLine("Scanning CSV File...");
			List<List<String>> rawLectureList = Standalone.CSVToTwoDimList(csvFileName, Standalone.CELL_BREAK, true);
			String mostCommonClassName = this.getMostCommonClassName(rawLectureList);
			String filename = this.getDefaultFileName(mostCommonClassName);
			standaloneFrame.addTextLine("Waiting for excel file selection...");
			String path = this.saveFile(filename);
			if (path != null) {
				standaloneFrame.addTextLine("Converting CSV file to excel file...");
				LectureWorkbook lectureWorkbook = new LectureWorkbook(path);

				List<Lecture> lectures = this.getLecturesFromRawLectureList(rawLectureList);

				if ((lectureWorkbook.getQuarterStartDate() == null || lectureWorkbook.getQuarterEndDate() == null)
						&& lectures.size() > 0) {
					Calendar dateInQuarter = (Calendar) lectures.get(0).getStartDate().clone();
					lectureWorkbook.setBorderDatesWithDateInQuarter(dateInQuarter);
				}

				lectureWorkbook.setLectures(lectures);
				this.saveFile(lectureWorkbook, path);
				standaloneFrame.addTextLine("Export completed successfully!");
				Thread.sleep(2000);
				standaloneFrame.close();
			}
		}
	}

	/**
	 * Saves the lectures as a workbook under the given filename. If there exists a
	 * file with the given filename, it is used as custom template workbook for the
	 * lectures. Otherwise the standard template workbook is used.
	 * 
	 * @param filename         The name for the file
	 * @param quarterStartDate The included start date of the quarter
	 * @param lectures         a list of lectures
	 * @throws RaplaException If loading or saving the file fails
	 */
	public void saveFile(LectureWorkbook lectureWorkbook, String filename) throws RaplaException {
		try {
			lectureWorkbook.saveToFile(filename);
		} catch (IOException e) {
			throw new RaplaException(e.getMessage(), e);
		}
	}

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
	 * Open a FileDialog and return the full path of the entered filename. If the
	 * FileDialog was closed or canceled, then null is returned.
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
	 * Converts a list of rapla objects into a list of lectures and sets the most
	 * common class name from the lectures.
	 * 
	 * @param objects The rapla objects
	 * @param columns The rapla columns
	 * @return A list of lectures
	 * @throws RaplaException
	 * @throws RaplaContextException
	 */
	private List<Lecture> getLecturesFromRawLectureList(List<List<String>> rawLectureList) {

		TimeZone timeZone = Standalone.TIME_ZONE;
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		sdf.setTimeZone(timeZone);

		List<Lecture> lectures = new ArrayList<Lecture>();

		for (List<String> row : rawLectureList) {
			String lectureName = row.get(Standalone.LECTURE_NAME_POSITION);

			Calendar lectureStartDate = null;
			Calendar startDate = new GregorianCalendar();
			startDate.setTimeZone(timeZone);
			try {
				startDate.setTime(sdf.parse(row.get(Standalone.LECTURE_START_DATE_POSITION)));
				lectureStartDate = startDate;
			} catch (ParseException e) {
				System.err.println("Cannot parse start date of the lecture \"" + lectureName + "\"");
			}

			Calendar lectureEndDate = null;
			Calendar endDate = new GregorianCalendar();
			endDate.setTimeZone(timeZone);
			try {
				endDate.setTime(sdf.parse(row.get(Standalone.LECTURE_END_DATE_POSITION)));
				lectureEndDate = endDate;
			} catch (ParseException e) {
				System.err.println("Cannot parse end date of the lecture \"" + lectureName + "\"");
			}

			String[] lectureResources = null;
			String resourcesString = row.get(Standalone.LECTURE_RESSOURCES_POSITION);
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
			String lecturers = row.get(Standalone.LECTURE_LECTURERS_POSITION);
			if (lecturers != null) {
				lectureLecturers = lecturers.split(", ");
			}

			Lecture lecture = new Lecture(lectureName, lectureStartDate, lectureEndDate, lectureResources,
					lectureLecturers);
			lectures.add(lecture);
		}

		return lectures;
	}

	private String getMostCommonClassName(List<List<String>> rawLectureList) {

		Map<String, Integer> classNames = new HashMap<String, Integer>();

		for (List<String> row : rawLectureList) {
			String element = row.get(Standalone.LECTURE_RESSOURCES_POSITION);
			if (element != null) {
				String[] resources = element.split(", ");
				for (String resource : resources) {
					if (!Export2ExcelMenu.resourceIsRoom(resource)) {
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

	public static List<List<String>> CSVToTwoDimList(String filename, String seperator, boolean hasTitleLine)
			throws IOException {
		BufferedReader bufferedReader = null;
		String line;

		List<List<String>> list = new ArrayList<List<String>>();

		try {
			bufferedReader = new BufferedReader(new FileReader(filename));
			if (hasTitleLine) {
				bufferedReader.readLine();
			}
			while ((line = bufferedReader.readLine()) != null) {
				List<String> elements = Arrays.asList(line.split(seperator));
				list.add(elements);
			}
		} finally {
			if (bufferedReader != null) {
				bufferedReader.close();
			}
		}
		return list;
	}

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

	public static void main(String[] args) throws Exception {
		Standalone standalone = new Standalone();
		standalone.export();
	}
}
