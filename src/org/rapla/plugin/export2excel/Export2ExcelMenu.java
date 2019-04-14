package org.rapla.plugin.export2excel;

import java.awt.Component;
import java.awt.FileDialog;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
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

import semesterTimeTable.excel.ExcelGenerator;
import semesterTimeTable.excel.Lecture;;

/**
 * Class representing the export to excel menu entry and its functionality.
 */
public class Export2ExcelMenu extends RaplaGUIComponent implements IdentifiableMenuEntry, ActionListener {

	String id = "export_file_text";
	JMenuItem item;

	public Export2ExcelMenu(RaplaContext sm) {
		super(sm);
		setChildBundleName(Export2ExcelPlugin.RESOURCE_FILE);
		item = new JMenuItem(getString(id));
		item.setIcon(getIcon("icon.export"));
		item.addActionListener(this);
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
			Date endDate = weekOfYearToDate(getQuarterEndWeekForAWeek(weekOfYear), Calendar.FRIDAY, year, timeZone);
			
			System.out.println(startDate);
			System.out.println(endDate);
			
			model.setStartDate(startDate);
			model.setEndDate(endDate);

			export(model);
		} catch (Exception ex) {
			showException(ex, getMainComponent());
		}
	}

	public String getId() {
		return id;
	}

	public JMenuItem getMenuElement() {
		return item;
	}

	private static final String LINE_BREAK = "\n";
	private static final String CELL_BREAK = ";";

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
						lectureResources = new String[resources.length / 2];
						for (int i = 0; i < lectureResources.length; i++) {
							lectureResources[i] = resources[2 * i + 1];
						}
					} else if (columnName == getString("persons")) {
						lectureLecturers = escape(value).split(", ");
					}
				}
			}

			Lecture lecture = new Lecture(lectureName, lectureStartDate, lectureEndDate, lectureResources,
					lectureLecturers);
			lectures.add(lecture);
		}
		ExcelGenerator excelGenerator = new ExcelGenerator(lectures);
		
		File excelTemplate = new File("Vorlesungsplan Beispiel.xlsx");
		excelGenerator.createTemplateFromFile(excelTemplate);
		
		excelGenerator.saveNewFile("test1.xlsx");
		excelGenerator.saveNewFile("test2.xlsx");

		/*
		 * byte[] bytes = buf.toString().getBytes();
		 * 
		 * DateFormat sdfyyyyMMdd = new SimpleDateFormat("yyyyMMdd"); final String
		 * calendarName =
		 * getQuery().getSystemPreferences().getEntryAsString(RaplaMainContainer.TITLE,
		 * getString("rapla.title")); String filename = calendarName + "-" +
		 * sdfyyyyMMdd.format(model.getStartDate()) + "-" +
		 * sdfyyyyMMdd.format(model.getEndDate()) + ".xlsx";
		 */
		String filename = "testname";
		final String extension = "xlsx";
		if (saveFile(loadFile(extension, filename))) {
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

	private String escape(Object cell) {
		return cell.toString().replace(LINE_BREAK, " ").replace(CELL_BREAK, " ");
	}

	public boolean saveFile(String filename) throws IOException {
		final File savedFile = new File(filename);
		byte[] content = new byte[] {};
		writeFile(savedFile, content);

		// Open

		return false;
	}

	public String loadFile(final String fileExtension, String filename) {
		final Frame frame = (Frame) SwingUtilities.getRoot(getMainComponent());
		final FileDialog fd = new FileDialog(frame, "Save (and load) Excel File", FileDialog.SAVE);

		fd.setFile(filename);

		fd.setLocation(50, 50);
		fd.setVisible(true);
		final String savedFileName = fd.getFile();

		if (savedFileName == null) {
			return null;
		} else {
			String path = createFullPath(fd, fileExtension);
			return path;
		}
	}

	/*
	 * public FileContent openFile(Frame frame,String dir, String[] fileExtensions)
	 * throws IOException { final FileDialog fd = new FileDialog(frame, "Open File",
	 * FileDialog.LOAD);
	 * 
	 * fd.setDirectory(dir); fd.setLocation(50, 50); fd.setVisible( true); final
	 * String openFileName = fd.getFile();
	 * 
	 * if (openFileName == null) { return null; } String path = createFullPath(fd);
	 * final FileInputStream openFile = new FileInputStream( path); FileContent
	 * content = new FileContent(); content.setName( openFileName);
	 * content.setInputStream( openFile ); return content; }
	 */

	private String createFullPath(final FileDialog fd, String extension) {
		String filename = fd.getFile();
		if (!filename.endsWith(extension)) {
			filename = filename + "." + extension;
		}
		return fd.getDirectory() + filename;
	}

	private void writeFile(final File savedFile, byte[] content) throws IOException {
		final FileOutputStream out;
		out = new FileOutputStream(savedFile);
		out.write(content);
		out.flush();
		out.close();
	}

	private Date weekOfYearToDate(int weekOfYear, int dayOfWeek, int year, TimeZone timeZone) {
		Calendar calendar = new GregorianCalendar();
		calendar.set(Calendar.DAY_OF_WEEK, dayOfWeek);
		calendar.set(Calendar.WEEK_OF_YEAR, weekOfYear);
		calendar.set(Calendar.YEAR, year);
		calendar.setTimeZone(timeZone);
		return calendar.getTime();
	}

	private int getQuarterStartWeekForAWeek(int weekOfYear) {
		int quarter = (int) (weekOfYear / 13.0);
		if (quarter < 2) {
			return quarter * 13 + 2;
		} else {
			return quarter * 13 + 1;
		}
	}

	private int getQuarterEndWeekForAWeek(int weekOfYear) {
		int quarter = (int) (weekOfYear / 13.0);
		if (quarter < 2) {
			return quarter * 13 + 13;
		} else {
			return quarter * 13 + 12;
		}
	}
}