package org.rapla.plugin.export2excel;

import java.awt.Component;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
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

import org.rapla.components.iolayer.IOInterface;
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
		// generates a text file from all filtered events;

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
				boolean isDate = columnClass.isAssignableFrom(java.util.Date.class);

				if (value != null) {
					if (column.getColumnName() == getString("name")) {
						lectureName = escape(value);
					} else if (column.getColumnName() == getString("start_date")) {
						if (isDate) {
							lectureStartDate = new GregorianCalendar();
							lectureStartDate.setTime((Date) value);
							lectureStartDate.setTimeZone(timeZone);
						}
					} else if (column.getColumnName() == getString("end_date")) {
						if (isDate) {
							lectureEndDate = new GregorianCalendar();
							lectureEndDate.setTime((Date) value);
							lectureEndDate.setTimeZone(timeZone);
						}
					} else if (column.getColumnName() == getString("resources")) {
						String[] resources = escape(value).split(", ");
						lectureResources = new String[resources.length / 2];
						for (int i = 0; i < lectureResources.length; i++) {
							lectureResources[i] = resources[2 * i + 1];
						}
					} else if (column.getColumnName() == getString("persons")) {
						lectureLecturers = escape(value).split(", ");
					}
				}
			}

			Lecture lecture = new Lecture(lectureName, lectureStartDate, lectureEndDate, lectureResources,
					lectureLecturers);
			lectures.add(lecture);
		}
		ExcelGenerator excelGenearator = new ExcelGenerator(lectures);

		/*
		 * byte[] bytes = buf.toString().getBytes();
		 * 
		 * DateFormat sdfyyyyMMdd = new SimpleDateFormat("yyyyMMdd"); final String
		 * calendarName =
		 * getQuery().getSystemPreferences().getEntryAsString(RaplaMainContainer.TITLE,
		 * getString("rapla.title")); String filename = calendarName + "-" +
		 * sdfyyyyMMdd.format(model.getStartDate()) + "-" +
		 * sdfyyyyMMdd.format(model.getEndDate()) + ".xlsx"; if (saveFile(bytes,
		 * filename, "xlsx")) { exportFinished(getMainComponent()); }
		 */
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

	public boolean saveFile(byte[] content, String filename, String extension) throws RaplaException {
		final Frame frame = (Frame) SwingUtilities.getRoot(getMainComponent());
		IOInterface io = getService(IOInterface.class);
		try {
			String file = io.saveFile(frame, null, new String[] { extension }, filename, content);
			return file != null;
		} catch (IOException e) {
			throw new RaplaException(e.getMessage(), e);
		}
	}
}