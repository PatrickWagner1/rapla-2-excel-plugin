package org.rapla.plugin.export2excel;

import java.awt.Component;
import java.awt.FileDialog;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
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
	 * @param model
	 * @throws Exception
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void export(final CalendarSelectionModel model) throws Exception {
		Collection<? extends RaplaTableColumn<?>> columns;
		List<Object> objects = new ArrayList<Object>();
		User user = model.getUser();
		columns = TableConfig.loadColumns(getContainer(), "appointments",
				TableViewExtensionPoints.APPOINTMENT_TABLE_COLUMN, user);
		final List<AppointmentBlock> blocks = model.getBlocks();
		objects.addAll(blocks);

		List<Lecture> lectures = new ArrayList<Lecture>();

		TimeZone timeZone = getRaplaLocale().getTimeZone();

		String lecturesTitle = "Vorlesungsplan";
		Map<String, Integer> classNames = new HashMap<String, Integer>();

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
								resource = resource.replaceAll("\\(.*\\)", "");
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

		Lecture.generateLecturesGroupId(lectures);

		Calendar quarterStartDate = new GregorianCalendar();
		quarterStartDate.setTime(model.getStartDate());
		quarterStartDate.setTimeZone(getRaplaLocale().getTimeZone());

		String className = getHighestCountKey(classNames);
		if (className != null && className != "") {
			lecturesTitle += "_" + className;
		}

		DateFormat sdfyyyyMMdd = new SimpleDateFormat("yyyy-MM-dd");
		String currentDate = sdfyyyyMMdd.format(new Date());
		final String extension = "xlsx";
		String filename = lecturesTitle + "_" + currentDate + "." + extension;
		String path = loadFile(filename);
		if (path != null) {
			saveFile(path, quarterStartDate, lectures);
			exportFinished(getMainComponent());
		}
	}

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

	public void saveFile(String filename, Calendar quarterStartDate, List<Lecture> lectures) throws RaplaException {
		try {
			LectureWorkbook excelGenerator = new LectureWorkbook(filename, quarterStartDate, lectures);
			excelGenerator.saveToFile(filename);
		} catch (IOException e) {
			throw new RaplaException(e.getMessage(), e);
		}
	}

	public String loadFile(String filename) {
		final Frame frame = (Frame) SwingUtilities.getRoot(getMainComponent());
		final FileDialog fd = new FileDialog(frame, "Save (and load) Excel File", FileDialog.SAVE);

		fd.setFile(filename);

		fd.setLocation(50, 50);
		fd.setVisible(true);
		final String savedFileName = fd.getFile();

		if (savedFileName == null) {
			return null;
		} else {
			String path = createFullPath(fd);
			return path;
		}
	}

	private String escape(Object cell) {
		return cell.toString().replace(LINE_BREAK, " ").replace(CELL_BREAK, " ");
	}

	private String createFullPath(final FileDialog fd) {
		String filename = fd.getFile();
		return fd.getDirectory() + filename;
	}

	private Date weekOfYearToDate(int weekOfYear, int dayOfWeek, int year, TimeZone timeZone) {
		Calendar calendar = new GregorianCalendar();
		calendar.set(Calendar.DAY_OF_WEEK, dayOfWeek);
		calendar.set(Calendar.WEEK_OF_YEAR, weekOfYear);
		calendar.set(Calendar.YEAR, year);
		calendar.setTimeZone(timeZone);
		calendar.set(Calendar.HOUR_OF_DAY, 0);
		calendar.set(Calendar.MINUTE, 0);
		calendar.set(Calendar.SECOND, 0);
		return calendar.getTime();
	}

	private int getQuarterStartWeekForAWeek(int weekOfYear) {
		int quarter = weekOfYear / 13;
		if (quarter < 2) {
			return quarter * 13 + 2;
		} else {
			return quarter * 13 + 1;
		}
	}

	private int getQuarterEndWeekForAWeek(int weekOfYear) {
		int quarter = weekOfYear / 13;
		if (quarter < 2) {
			return quarter * 13 + 13;
		} else {
			return quarter * 13 + 12;
		}
	}

	private String getHighestCountKey(Map<String, Integer> map) {
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