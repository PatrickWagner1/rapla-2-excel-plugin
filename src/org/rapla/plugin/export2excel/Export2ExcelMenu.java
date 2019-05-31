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
import org.rapla.framework.RaplaContextException;
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

	/** Line break character */
	private static final String LINE_BREAK = "\n";

	/** Cell break character */
	private static final String CELL_BREAK = ";";

	/** File extension for excel files */
	private static final String FILE_EXTENSION = "xlsx";

	/**
	 * Creates the export menu for the export to excel plugin.
	 * 
	 * @param sm The rapla context
	 */
	public Export2ExcelMenu(RaplaContext sm) {
		super(sm);
		setChildBundleName(Export2ExcelPlugin.RESOURCE_FILE);
		this.item = new JMenuItem(getString(this.id));
		this.item.setIcon(getIcon("icon.export"));
		this.item.addActionListener(this);
	}

	/**
	 * Event handler for clicking on the export to excel menu entry.
	 * 
	 * @param evt The action event
	 */
	public void actionPerformed(ActionEvent evt) {
		try {
			this.export();
		} catch (Exception ex) {
			showException(ex, getMainComponent());
		}
	}

	/**
	 * Getter method for the id of the menu.
	 * 
	 * @return id The id for the menu
	 */
	public String getId() {
		return this.id;
	}

	/**
	 * Getter method for the menu item.
	 * 
	 * @return The menu item
	 */
	public JMenuItem getMenuElement() {
		return this.item;
	}

	/**
	 * Method implementing the actual exporting functionality. It is called by the
	 * event handler for clicking on the export menu entry.
	 * 
	 * @param model The calendar selection model
	 * @throws IOException           If saving or loading a workbook fails
	 * @throws RaplaException
	 * @throws RaplaContextException
	 */
	public void export() throws IOException, RaplaContextException, RaplaException {
		CalendarSelectionModel preModel = getService(CalendarSelectionModel.class);
		String mostCommonClassName = this.getMostCommonClassName(preModel);
		String filename = this.getDefaultFileName(mostCommonClassName);
		String path = this.loadFile(filename);
		if (path != null) {
			LectureWorkbook lectureWorkbook = new LectureWorkbook(path);

			CalendarSelectionModel model = getService(CalendarSelectionModel.class);

			if (lectureWorkbook.getQuarterStartDate() == null || lectureWorkbook.getQuarterEndDate() == null) {
				TimeZone timeZone = getRaplaLocale().getTimeZone();
				Calendar dateInQuarter = new GregorianCalendar();
				dateInQuarter.setTime(model.getStartDate());
				dateInQuarter.setTimeZone(timeZone);

				lectureWorkbook.setBorderDatesWithDateInQuarter(dateInQuarter);
			}

			Date startDate = lectureWorkbook.getQuarterStartDate().getTime();
			model.setStartDate(startDate);
			Date endDate = lectureWorkbook.getQuarterEndDate().getTime();
			model.setEndDate(endDate);

			List<Lecture> lectures = this.getLecturesFromRaplaModel(model);

			lectureWorkbook.setLectures(lectures);
			lectureWorkbook.saveToFile(path);
			this.exportFinished(getMainComponent(), lectureWorkbook.getErrorOutput().getErrorOutput());
		}
	}

	/**
	 * Shows dialog window with export finished message.
	 * 
	 * @param topLevel
	 * @param errorLogging The error logging of the export (empty string, if no
	 *                     errors occurred)
	 * @return
	 */
	protected boolean exportFinished(Component topLevel, String errorLogging) {
		if (errorLogging == null) {
			errorLogging = "";
		}
		try {
			errorLogging = errorLogging.replace("\n", "<br>");
			String dialogText = getString("file_saved") + "<br>" + errorLogging;
			DialogUI dlg = DialogUI.create(getContext(), topLevel, true, getString("export"), dialogText,
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
	private String getDefaultFileName(String className) {
		String lecturesTitle = getString("lectures_file_name").replace(' ', '_');
		if (className != null && className != "") {
			lecturesTitle += "_" + className;
		}

		return lecturesTitle + "." + Export2ExcelMenu.FILE_EXTENSION;
	}

	/**
	 * Extracts a collection of rapla table column objects from the given calendar
	 * selection model.
	 * 
	 * @param model The calendar selection model
	 * @return A collection of rapla table column objects
	 * @throws RaplaContextException
	 * @throws RaplaException
	 */
	private Collection<? extends RaplaTableColumn<?>> getColumnsFromModel(CalendarSelectionModel model)
			throws RaplaContextException, RaplaException {
		Collection<? extends RaplaTableColumn<?>> columns;
		User user = model.getUser();
		columns = TableConfig.loadColumns(getContainer(), "appointments",
				TableViewExtensionPoints.APPOINTMENT_TABLE_COLUMN, user);
		return columns;
	}

	/**
	 * Extracts a list of raw lecture objects from the given calendar selection
	 * model.
	 * 
	 * @param model The calendar selection model
	 * @return A list of raw lecture objects
	 * @throws RaplaException
	 */
	private List<Object> getObjectsFromModel(CalendarSelectionModel model) throws RaplaException {
		final List<AppointmentBlock> blocks = model.getBlocks();
		List<Object> objects = new ArrayList<Object>();
		objects.addAll(blocks);
		return objects;
	}

	/**
	 * Extracts a list of lectures from the given calendar selection model.
	 * 
	 * @param model The calendar selection model
	 * @return A list of lectures
	 * @throws RaplaContextException
	 * @throws RaplaException
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	private List<Lecture> getLecturesFromRaplaModel(CalendarSelectionModel model)
			throws RaplaContextException, RaplaException {
		Collection<? extends RaplaTableColumn<?>> columns = this.getColumnsFromModel(model);
		List<Object> objects = this.getObjectsFromModel(model);

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
							if (Export2ExcelMenu.resourceIsRoom(resource)) {
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

		return lectures;
	}

	/**
	 * Extracts the most common class name from the given calendar selection model.
	 * 
	 * The class names are all resources except rooms.
	 * 
	 * @param model The calendar selection model
	 * @return The most common class name
	 * @throws RaplaContextException
	 * @throws RaplaException
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	private String getMostCommonClassName(CalendarSelectionModel model) throws RaplaContextException, RaplaException {
		Collection<? extends RaplaTableColumn<?>> columns = this.getColumnsFromModel(model);
		List<Object> objects = this.getObjectsFromModel(model);

		Map<String, Integer> classNames = new HashMap<String, Integer>();

		for (Object row : objects) {
			for (RaplaTableColumn column : columns) {
				Object value = column.getValue(row);
				String columnName = column.getColumnName();

				if (columnName == getString("resources") && value != null) {
					String[] resources = this.escape(value).split(", ");
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
		}
		return getHighestCountKey(classNames);
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
}