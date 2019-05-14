package org.rapla.plugin.export2excel;

import java.awt.Component;
import java.awt.FileDialog;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TimeZone;

import javax.swing.JMenuItem;
import javax.swing.SwingUtilities;

import org.rapla.entities.User;
import org.rapla.entities.domain.AppointmentBlock;
import org.rapla.facade.CalendarSelectionModel;
import org.rapla.framework.RaplaContext;
import org.rapla.framework.RaplaException;
import org.rapla.gui.RaplaGUIComponent;
import org.rapla.gui.toolkit.DialogUI;
import org.rapla.gui.toolkit.IdentifiableMenuEntry;
import org.rapla.plugin.tableview.RaplaTableColumn;
import org.rapla.plugin.tableview.TableViewExtensionPoints;
import org.rapla.plugin.tableview.internal.TableConfig;

import semesterTimeTable.excel.LectureWorkbook;
import semesterTimeTable.excel.Lecture;

/**
 * Class representing the export to excel menu entry and its functionality.
 */
public class Export2ExcelMenu extends RaplaGUIComponent implements IdentifiableMenuEntry, ActionListener {

	String id = "export_file_text";
	JMenuItem item;

	private static final String LINE_BREAK = "\n";
	private static final String CELL_BREAK = ";";

	private static final String FILE_EXTENSION = "xlsx";

	private String mostCommonClassName;

	public Export2ExcelMenu(RaplaContext sm) {
		super(sm);
		setChildBundleName(Export2ExcelPlugin.RESOURCE_FILE);
		this.item = new JMenuItem(getString(this.id));
		this.item.setIcon(getIcon("icon.export"));
		this.item.addActionListener(this);
		this.setMostCommonClassName("");
	}

	/**
	 * Event handler for clicking on the export to excel menu entry.
	 * 
	 * @param evt
	 */
	public void actionPerformed(ActionEvent evt) {
		try {
			CalendarSelectionModel model = getService(CalendarSelectionModel.class);

			TimeZone timeZone = getRaplaLocale().getTimeZone();
			Calendar calendar = new GregorianCalendar();

			calendar.setTime(model.getStartDate());
			calendar.setTimeZone(timeZone);

			int weekOfYear = calendar.get(Calendar.WEEK_OF_YEAR);
			int year = calendar.get(Calendar.YEAR);

			Date startDate = weekOfYearToDate(getQuarterStartWeekForAWeek(weekOfYear), Calendar.MONDAY, year, timeZone);
			Date endDate = weekOfYearToDate(getQuarterEndWeekForAWeek(weekOfYear), Calendar.SATURDAY, year, timeZone);

			model.setStartDate(startDate);
			model.setEndDate(endDate);

			export(model);
		} catch (Exception ex) {
			showException(ex, getMainComponent());
		}
	}

	/**
	 * Getter method for the id of the menu.
	 * 
	 * @return id
	 */
	public String getId() {
		return this.id;
	}

	/**
	 * Getter method for the menu item.
	 * 
	 * @return item
	 */
	public JMenuItem getMenuElement() {
		return this.item;
	}

	/**
	 * Method implementing the actual exporting functionality. It is called by the
	 * event handler for clicking on the export menu entry.
	 * 
	 * @param model The calendar selection model
	 * @throws Exception
	 */
	public void export(final CalendarSelectionModel model) throws Exception {
		Collection<? extends RaplaTableColumn<?>> columns;
		List<Object> objects = new ArrayList<Object>();
		User user = model.getUser();
		columns = TableConfig.loadColumns(getContainer(), "appointments",
				TableViewExtensionPoints.APPOINTMENT_TABLE_COLUMN, user);
		final List<AppointmentBlock> blocks = model.getBlocks();
		objects.addAll(blocks);

		List<Lecture> lectures = this.raplaObjectsToLectures(objects, columns);

		TimeZone timeZone = getRaplaLocale().getTimeZone();

		Calendar quarterStartDate = new GregorianCalendar();
		quarterStartDate.setTime(model.getStartDate());
		quarterStartDate.setTimeZone(timeZone);

		Calendar quarterEndDate = new GregorianCalendar();
		quarterEndDate.setTime(model.getEndDate());
		quarterEndDate.setTimeZone(timeZone);

		String filename = this.getDefaultFileName();
		String path = loadFile(filename);
		if (path != null) {
			saveFile(path, quarterStartDate, quarterEndDate, lectures);
			exportFinished(getMainComponent());
		}
	}

	/**
	 * Shows dialog window with export finished message.
	 * 
	 * @param topLevel
	 * @return
	 */
	protected boolean exportFinished(Component topLevel) {
		try {
			DialogUI dlg = DialogUI.create(getContext(), topLevel, true, getString("export"), getString("file_saved"),
					new String[] { getString("ok") });
			dlg.setIcon(getIcon("icon.export"));
			dlg.setDefault(0);
			dlg.start();
			return (dlg.getSelectedIndex() == 0);
		} catch (RaplaException e) {
			return true;
		}

	}

	/**
	 * Saves the lectures as a workbook under the given filename. If there exists a
	 * file with the given filename, it is used as custom template workbook for the
	 * lectures. Otherwise the standard template workbook is used.
	 * 
	 * @param filename         The name for the file
	 * @param quarterStartDate The included start date of the quarter
	 * @param quarterEndDate   The excluded end date of the quarter
	 * @param lectures         a list of lectures
	 * @throws RaplaException If loading or saving the file fails
	 */
	public void saveFile(String filename, Calendar quarterStartDate, Calendar quarterEndDate, List<Lecture> lectures)
			throws RaplaException {
		try {
			LectureWorkbook excelGenerator = new LectureWorkbook(filename, quarterStartDate, quarterEndDate, lectures);
			excelGenerator.saveToFile(filename);
		} catch (IOException e) {
			throw new RaplaException(e.getMessage(), e);
		}
	}

	/**
	 * Open a FileDialog and return the full path of the entered filename. If the
	 * FileDialog was closed or canceled, then null is returned.
	 * 
	 * @param filename The default filename for the FileDialog
	 * @return The full path of the entered filename, or null if closed or canceled
	 *         the FileDialog
	 */
	public String loadFile(String filename) {
		final Frame frame = (Frame) SwingUtilities.getRoot(getMainComponent());
		final FileDialog fd = new FileDialog(frame, "Save (and load) Excel File", FileDialog.SAVE);

		fd.setFile(filename);

		fd.setLocation(50, 50);
		fd.setVisible(true);
		final String savedFileName = fd.getFile();

		String path = null;
		if (savedFileName != null) {
			path = createFullPath(fd);
		}

		return path;
	}

	/**
	 * Returns the plain string of a cell without line breaks and cell breaks.
	 * 
	 * @param cell
	 * @return The plain string of a cell
	 */
	private String escape(Object cell) {
		return cell.toString().replace(LINE_BREAK, " ").replace(CELL_BREAK, " ");
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
	private String getDefaultFileName() {
		String lecturesTitle = getString("lectures_file_name").replace(' ', '_');
		String className = this.getMostCommonClassName();
		if (className != null && className != "") {
			lecturesTitle += "_" + className;
		}

		return lecturesTitle + "." + Export2ExcelMenu.FILE_EXTENSION;
	}

	/**
	 * Converts a list of rapla objects into a list of lectures and sets the most
	 * common class name from the lectures.
	 * 
	 * @param objects The rapla objects
	 * @param columns The rapla columns
	 * @return A list of lectures
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	private List<Lecture> raplaObjectsToLectures(List<Object> objects,
			Collection<? extends RaplaTableColumn<?>> columns) {

		Map<String, Integer> classNames = new HashMap<String, Integer>();
		TimeZone timeZone = getRaplaLocale().getTimeZone();
		List<Lecture> lectures = new ArrayList<Lecture>();
		for (Object row : objects) {
			String lectureName = null;
			Calendar lectureStartDate = null;
			Calendar lectureEndDate = null;
			String[] lectureResources = null;
			String[] lectureLecturers = null;

			for (RaplaTableColumn column : columns) {
				Object value = column.getValue(row);
				Class columnClass = column.getColumnClass();
				String columnName = column.getColumnName();
				boolean isDate = columnClass.isAssignableFrom(java.util.Date.class);

				if (value != null) {
					if (columnName == getString("name")) {
						lectureName = escape(value);
					} else if (columnName == getString("start_date") && isDate) {
						lectureStartDate = new GregorianCalendar();
						lectureStartDate.setTime((Date) value);
						lectureStartDate.setTimeZone(timeZone);
					} else if (columnName == getString("end_date") && isDate) {
						lectureEndDate = new GregorianCalendar();
						lectureEndDate.setTime((Date) value);
						lectureEndDate.setTimeZone(timeZone);
					} else if (columnName == getString("resources")) {
						String[] resources = escape(value).split(", ");
						ArrayList<String> rooms = new ArrayList<String>();
						for (String resource : resources) {
							if (resource.endsWith(")")) {
								resource = resource.replaceAll(" \\(.*\\)", "");
								int count = 0;
								if (classNames.containsKey(resource)) {
									count = classNames.get(resource);
								}
								count++;
								classNames.put(resource, count);
							} else {
								rooms.add(resource);
							}
						}
						lectureResources = new String[rooms.size()];
						lectureResources = rooms.toArray(lectureResources);

					} else if (columnName == getString("persons")) {
						lectureLecturers = escape(value).split(", ");
					}
				}
			}

			Lecture lecture = new Lecture(lectureName, lectureStartDate, lectureEndDate, lectureResources,
					lectureLecturers);
			lectures.add(lecture);
		}

		this.setMostCommonClassName(getHighestCountKey(classNames));
		return lectures;
	}

	/**
	 * Returns the most common class name of the current export.
	 * 
	 * @return The most common class name
	 */
	private String getMostCommonClassName() {
		return this.mostCommonClassName;
	}

	/**
	 * Sets the most common class name of the current export.
	 * 
	 * @param mostCommonClassName The most common class name
	 */
	private void setMostCommonClassName(String mostCommonClassName) {
		this.mostCommonClassName = mostCommonClassName;
	}

	/**
	 * Converts the given week of the year, day of the week and year to a Date.
	 * 
	 * @param weekOfYear The week of the year
	 * @param dayOfWeek  The day of the week
	 * @param year       The year
	 * @param timeZone   The time zone
	 * @return The date created from the given parameters
	 */
	private static Date weekOfYearToDate(int weekOfYear, int dayOfWeek, int year, TimeZone timeZone) {
		Calendar calendar = new GregorianCalendar();
		calendar.set(Calendar.DAY_OF_WEEK, dayOfWeek);
		calendar.set(Calendar.WEEK_OF_YEAR, weekOfYear);
		calendar.set(Calendar.YEAR, year);
		calendar.setTimeZone(timeZone);
		calendar.set(Calendar.HOUR_OF_DAY, 0);
		calendar.set(Calendar.MINUTE, 0);
		calendar.set(Calendar.SECOND, 0);
		calendar.set(Calendar.MILLISECOND, 0);
		return calendar.getTime();
	}

	/**
	 * Returns the first week of a quarter from a random week in this quarter.
	 * 
	 * @param weekOfYear A random week in a quarter
	 * @return The first week of the quarter
	 */
	private static int getQuarterStartWeekForAWeek(int weekOfYear) {
		int quarter = weekOfYear / 13;
		if (quarter < 2) {
			return quarter * 13 + 2;
		} else {
			return quarter * 13 + 1;
		}
	}

	/**
	 * Returns the last week of a quarter from a random week in this quarter.
	 * 
	 * @param weekOfYear A random week in a quarter
	 * @return The last week of the quarter
	 */
	private static int getQuarterEndWeekForAWeek(int weekOfYear) {
		int quarter = weekOfYear / 13;
		if (quarter < 2) {
			return quarter * 13 + 13;
		} else {
			return quarter * 13 + 12;
		}
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
			if (highestEntry == null) {
				highestEntry = entry;
			} else if (entry.getValue() > highestEntry.getValue()) {
				highestEntry = entry;
			}
		}
		if (highestEntry == null) {
			return null;
		} else {
			return highestEntry.getKey();
		}
	}
}