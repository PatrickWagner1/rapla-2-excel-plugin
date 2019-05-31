package semesterTimeTable.excel;

import java.io.File;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class containing style (especially color) configurations for a workbook.
 *
 */
public class ConfigWorkbook {

	/** Filename of configuration template */
	private final static String TEMPLATE_FILENAME = "rapla_config.xlsx";

	private Map<String, LectureProperties> lecturePropertiesMap;

	/** Map of values and their font to highlight the value with */
	private Map<String, XSSFFont> highlightedFonts;

	/**
	 * Array of all prefixes, which should be ignored by the grouping of lectures
	 */
	private String[] ignorePrefixes;

	/** The locale for the holidays */
	private Locale holidayLocale;

	/** An array of start weeks of quarters */
	private int[] quarterStartWeeks;

	/** The length of the exam week */
	private int examWeekLength;

	/** The start date of the quarter */
	private Calendar quarterStartDate;

	/** The text for the box on the bottom of the exam week */
	private XSSFRichTextString examWeekText;

	/** The font for the box on the bottom of the exam week */
	private XSSFFont examWeekFont;

	/** The fill color for the box on the bottom of the exam week */
	private XSSFColor examWeekFillColor;

	/** The workbook containing the configurations */
	private XSSFWorkbook workbook;

	/** The file of the workbook containing the configurations */
	private File file;

	/** The status, if the configuration file was recreated or not */
	private boolean isNewConfig;

	/**
	 * Loads the style configuration of the configuration template.
	 * 
	 * @throws IOException If loading the template file fails
	 */
	public ConfigWorkbook() throws IOException {
		this.setWorkbook(ApachePOIWrapper
				.loadWorkbookFromInputStream(LectureWorkbook.getTemplateInputStream(ConfigWorkbook.TEMPLATE_FILENAME)));
		this.initConfigWorkbook();
		this.setIsNewConfig(false);
	}

	/**
	 * Loads the custom style configuration from the given path, if there is a file
	 * in the path with the {@link ConfigWorkbook#TEMPLATE_FILENAME} as filename.
	 * Otherwise the style configuration of the template loads and creates a new
	 * configuration file in the given path with the
	 * {@link ConfigWorkbook#TEMPLATE_FILENAME} as filename.
	 * 
	 * @param pathToTemplate The path to the custom template
	 * @throws IOException If reading one file failed
	 */
	public ConfigWorkbook(String pathToTemplate) throws IOException {
		File file = new File(pathToTemplate, ConfigWorkbook.TEMPLATE_FILENAME);
		if (file.exists()) {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromFile(file));
			this.initConfigWorkbook();
			this.setIsNewConfig(false);
		} else {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromInputStream(
					LectureWorkbook.getTemplateInputStream(ConfigWorkbook.TEMPLATE_FILENAME)));
			this.initConfigWorkbook(file);
			this.setIsNewConfig(true);
		}
	}

	/**
	 * Loads the custom style configuration from the given path, if there is a file
	 * in the path with the {@link ConfigWorkbook#TEMPLATE_FILENAME} as filename.
	 * Otherwise the style configuration of the template loads and creates a new
	 * configuration file containing the given lecture names in the given path with
	 * the {@link ConfigWorkbook#TEMPLATE_FILENAME} as filename.
	 * 
	 * All lecture names starts with the {@link ConfigWorkbook#REMOVE_PREFIX} are
	 * not added to the configuration file.
	 * 
	 * All lecture names starts with the {@link ConfigWorkbook#ignorePrefixes} are
	 * added without the prefix. If there exists a lecture name with this new name,
	 * it is not added.
	 * 
	 * @param pathToTemplate The path to the custom template
	 * @param lectureNames   A list of lecture names
	 * @throws IOException If reading one file or saving the configuration file
	 *                     failed
	 */
	public ConfigWorkbook(String pathToTemplate, List<String> lectureNames) throws IOException {
		File file = new File(pathToTemplate, ConfigWorkbook.TEMPLATE_FILENAME);
		if (file.exists()) {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromFile(file));
			this.initConfigWorkbook();
			this.setIsNewConfig(false);
		} else {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromInputStream(
					LectureWorkbook.getTemplateInputStream(ConfigWorkbook.TEMPLATE_FILENAME)));
			this.initConfigWorkbook(file, lectureNames);
			this.setIsNewConfig(true);
		}
	}

	/**
	 * Loads the custom style configuration from the given path, if there is a file
	 * in the path with the given template filename. Otherwise the style
	 * configuration of the template loads and creates a new configuration file in
	 * the given path with the given template filename.
	 * 
	 * @param pathToTemplate   The path to the custom template
	 * @param templateFilename The name of the custom template
	 * @throws IOException If reading one file failed
	 */
	public ConfigWorkbook(String pathToTemplate, String templateFilename) throws IOException {
		File file = new File(pathToTemplate, templateFilename);
		if (file.exists()) {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromFile(file));
			this.initConfigWorkbook();
			this.setIsNewConfig(false);
		} else {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromInputStream(
					LectureWorkbook.getTemplateInputStream(ConfigWorkbook.TEMPLATE_FILENAME)));
			this.initConfigWorkbook(file);
			this.setIsNewConfig(true);
		}
	}

	/**
	 * Loads the custom style configuration from the given path, if there is a file
	 * in the path with the given template filename. Otherwise the style
	 * configuration of the template loads and creates a new configuration file
	 * containing the given lecture names in the given path with the given filename.
	 * 
	 * All lecture names starts with the {@link ConfigWorkbook#REMOVE_PREFIX} are
	 * not added to the configuration file.
	 * 
	 * All lecture names starts with the {@link ConfigWorkbook#ignorePrefixes} are
	 * added without the prefix. If there exists a lecture name with this new name,
	 * it is not added.
	 * 
	 * @param pathToTemplate   The path to the custom template
	 * @param templateFilename The name of the custom template
	 * @param lectureNames     A list of lecture names
	 * @throws IOException If reading one file or saving the configuration file
	 *                     failed
	 */
	public ConfigWorkbook(String pathToTemplate, String templateFilename, List<String> lectureNames)
			throws IOException {
		File file = new File(pathToTemplate, templateFilename);
		if (file.exists()) {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromFile(file));
			this.initConfigWorkbook();
			this.setIsNewConfig(false);
		} else {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromInputStream(
					LectureWorkbook.getTemplateInputStream(ConfigWorkbook.TEMPLATE_FILENAME)));
			this.initConfigWorkbook(file, lectureNames);
			this.setIsNewConfig(true);
		}
	}

	/**
	 * Returns the workbook containing the configurations.
	 * 
	 * @return The configuration workbook
	 */
	private XSSFWorkbook getWorkbook() {
		return this.workbook;
	}

	/**
	 * Sets the workbook containing the configurations.
	 * 
	 * @param workbook The configuration workbook
	 */
	private void setWorkbook(XSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	/**
	 * Returns true if the configuration file was recreated, otherwise false.
	 * 
	 * @return True if the configuration file was recreated, otherwise false
	 */
	public boolean isNewConfig() {
		return this.isNewConfig;
	}

	/**
	 * Sets the status if the configuration file was recreated or not.
	 * 
	 * @param isNewConfig True if the configuration file was recreated, otherwise
	 *                    false
	 */
	private void setIsNewConfig(boolean isNewConfig) {
		this.isNewConfig = isNewConfig;
	}

	/**
	 * Returns a map of the lecture properties.
	 * 
	 * @return A map of the lecture properties
	 */
	public Map<String, LectureProperties> getLecturePropertiesMap() {
		return this.lecturePropertiesMap;
	}

	/**
	 * Returns a map of values and their font to highlight the value with.
	 * 
	 * @return A map of value font pairs
	 */
	public Map<String, XSSFFont> getHighlightedFonts() {
		return this.highlightedFonts;
	}

	/**
	 * Returns an array of prefixes, which should be ignored.
	 * 
	 * @return An array of ignore prefixes
	 */
	public String[] getIgnorePrefixes() {
		return this.ignorePrefixes;
	}

	/**
	 * Sets the ignore prefixes contained in the zero based column 3 in the first
	 * sheet of the workbook.
	 */
	private void setIgnorePrefixes() {
		XSSFSheet sheet = ApachePOIWrapper.getSheet(this.getWorkbook());
		int lastRowNum = sheet.getLastRowNum();
		this.ignorePrefixes = ApachePOIWrapper.getStringValuesFromWorkbook(sheet,
				new CellRangeAddress(2, lastRowNum, 4, 4));
	}

	/**
	 * Returns the locale for the holidays.
	 * 
	 * @return The locale for the holidays
	 */
	public Locale getHolidayLocale() {
		return this.holidayLocale;
	}

	/**
	 * Sets the locale for the holidays contained in the zero based column 4 in the
	 * first sheet of the workbook.
	 */
	private void setHolidayLocale() {
		XSSFSheet sheet = ApachePOIWrapper.getSheet(this.getWorkbook());
		int lastRowNum = sheet.getLastRowNum();
		String[] localeValues = ApachePOIWrapper.getStringValuesFromWorkbook(sheet,
				new CellRangeAddress(2, lastRowNum, 5, 5));
		if (localeValues.length > 0) {
			String country = localeValues[0];
			String variant = "";
			for (int i = 1; i < localeValues.length; i++) {
				variant += "_" + localeValues[i];
			}
			if (variant == "") {
				this.holidayLocale = new Locale("de", country);
			} else {
				variant = variant.substring(1);
				this.holidayLocale = new Locale("de", country, variant);
			}
		} else {
			this.holidayLocale = new Locale("de", "de", "bw");
		}
	}

	/**
	 * Returns an array of all start weeks of quarters.
	 * 
	 * @return The start weeks of quarters
	 */
	public int[] getQuarterStartWeeks() {
		return this.quarterStartWeeks;
	}

	/**
	 * Sets the start weeks of quarters contained in the zero based column 5 in the
	 * first sheet of the workbook.
	 */
	private void setQuarterStartWeeks() {
		XSSFSheet sheet = ApachePOIWrapper.getSheet(this.getWorkbook());
		int lastRowNum = sheet.getLastRowNum();
		int[] quarterStartWeeks = ApachePOIWrapper.getIntegerValuesFromWorkbook(sheet,
				new CellRangeAddress(2, lastRowNum, 6, 6));
		if (quarterStartWeeks.length < 1) {
			quarterStartWeeks = new int[] { 2, 15, 27, 40 };
		}
		this.quarterStartWeeks = quarterStartWeeks;
	}

	/**
	 * Returns the length of the exam week. Only Monday to Friday counts as days.
	 * 
	 * @return The length of the exam week
	 */
	public int getExamWeekLength() {
		return this.examWeekLength;
	}

	/**
	 * Sets the length of the exam week contained in the zero based row 2 in the
	 * zero based column 6 in the first sheet of the workbook.
	 * 
	 * If the value is smaller than zero, the length will be zero. If the value is
	 * higher than ten, the length will be ten. If the value is not a number, the
	 * default length of six will be used.
	 */
	private void setExamWeekLength() {
		XSSFSheet sheet = ApachePOIWrapper.getSheet(this.getWorkbook());
		int[] lengths = ApachePOIWrapper.getIntegerValuesFromWorkbook(sheet, new CellRangeAddress(2, 2, 7, 7));
		if (lengths.length == 1) {
			if (lengths[0] < 0) {
				this.examWeekLength = 0;
			} else if (lengths[0] > 10) {
				this.examWeekLength = 10;
			} else {
				this.examWeekLength = lengths[0];
			}
		} else {
			this.examWeekLength = 6;
		}
	}

	/**
	 * Returns the start date of the quarter.
	 * 
	 * @return The start date of the quarter
	 */
	public Calendar getQuarterStartDate() {
		return this.quarterStartDate;
	}

	/**
	 * Sets the start date of the quarter contained in the zero based column 7 in
	 * the first sheet of the workbook.
	 * 
	 * The value in the zero based row 2 is the start date for the quarter and the
	 * zero based row 3 is the ID for the time zone which should be used for the
	 * start date. If no time zone is set, the time zone with the id "GMT" will be
	 * used as default.
	 */
	private void setQuarterStartDate() {
		XSSFSheet sheet = ApachePOIWrapper.getSheet(this.getWorkbook());
		Date[] dates = ApachePOIWrapper.getDateValuesFromWorkbook(sheet, new CellRangeAddress(2, 2, 8, 8));
		String[] timeZoneCodes = ApachePOIWrapper.getStringValuesFromWorkbook(sheet, new CellRangeAddress(3, 3, 8, 8));

		TimeZone timeZone;
		Calendar quarterStartDate;
		if (timeZoneCodes.length == 1) {
			timeZone = TimeZone.getTimeZone(timeZoneCodes[0]);
		} else {
			timeZone = TimeZone.getTimeZone("GMT");
		}

		if (dates.length == 1) {
			quarterStartDate = new GregorianCalendar();
			quarterStartDate.setTime(dates[0]);
			quarterStartDate.setTimeZone(timeZone);
			this.quarterStartDate = quarterStartDate;
		}
	}

	/**
	 * Returns the text for the box at the bottom of the exam week.
	 * 
	 * @return The text for the exam week box
	 */
	public XSSFRichTextString getExamWeekText() {
		return this.examWeekText;
	}

	/**
	 * Returns the font for the box at the bottom of the exam week.
	 * 
	 * @return The font for the exam week box
	 */
	public XSSFFont getExamWeekFont() {
		return this.examWeekFont;
	}

	/**
	 * Returns the fill color for the box at the bottom of the exam week.
	 * 
	 * @return The fill color for the exam week box
	 */
	public XSSFColor getExamWeekFillColor() {
		return this.examWeekFillColor;
	}

	/**
	 * Sets the text, font and fill color for the box at the bottom of the exam
	 * week. The text and styles are contained in the zero based row 2 in the zero
	 * based column 8 in the first sheet of the workbook.
	 */
	private void setExamWeekTextAndStyle() {
		XSSFSheet sheet = ApachePOIWrapper.getSheet(this.getWorkbook());
		XSSFRow row = sheet.getRow(2);
		if (row != null) {
			XSSFCell cell = row.getCell(9);
			if (cell != null && cell.getCellType() == CellType.STRING) {
				XSSFRichTextString richText = cell.getRichStringCellValue();
				XSSFCellStyle cellStyle = cell.getCellStyle();
				XSSFColor fillColor = cellStyle.getFillForegroundColorColor();
				fillColor = fillColor == null ? cellStyle.getFillBackgroundColorColor() : fillColor;
				XSSFFont font = cellStyle.getFont();

				this.examWeekText = richText;
				this.examWeekFont = font;
				this.examWeekFillColor = fillColor;
			}
		}
	}

	/**
	 * Closes the workbook containing the configurations.
	 * 
	 * @throws IOException If closing the workbook failed
	 */
	public void close() throws IOException {
		this.getWorkbook().close();
	}

	/**
	 * Returns the file of the workbook containing the configurations.
	 * 
	 * @return The file of the configuration workbook
	 */
	private File getFile() {
		return this.file;
	}

	/**
	 * Sets the file of the workbook containing the configurations.
	 * 
	 * @param file The file for the configuration workbook
	 */
	private void setFile(File file) {
		this.file = file;
	}

	/**
	 * Creates a new configuration file with the configuration template and the
	 * given lectures and sets all the configurable variables.
	 * 
	 * Only new Lecture names are added to the configuration file.
	 * 
	 * All lecture names starts with the {@link ConfigWorkbook#ignorePrefixes} are
	 * added without the prefix. If there exists a lecture name with this new name,
	 * it is not added.
	 * 
	 * @param file         The file to save the configuration workbook into
	 * @param lectureNames A list of lecture names
	 * @throws IOException If saving the configuration workbook or closing the
	 *                     workbook failed
	 */
	private void initConfigWorkbook(File file, List<String> lectureNames) throws IOException {
		this.addLectureNames(lectureNames);
		this.initConfigWorkbook(file);
	}

	/**
	 * Creates a new configuration file with the configuration template and the
	 * given lectures and sets all the configurable variables.<
	 * 
	 * @param file The file to save the configuration workbook into
	 * @throws IOException If saving the configuration workbook or closing the
	 *                     workbook failed
	 */
	private void initConfigWorkbook(File file) throws IOException {
		this.setFile(file);
		ApachePOIWrapper.saveWorkbookToFile(this.getWorkbook(), file);
		this.initConfigWorkbook();
	}

	/**
	 * Sets all the configurable variables.
	 * 
	 * @param workbook The template configuration workbook
	 * @throws IOException If closing the workbook failed
	 */
	private void initConfigWorkbook() throws IOException {
		int rowNum = 2;
		XSSFSheet sheet = this.getWorkbook().getSheetAt(0);
		int lastRowNum = sheet.getLastRowNum();
		this.lecturePropertiesMap = ConfigWorkbook.getMappedLectureProperties(sheet,
				new CellRangeAddress(rowNum, lastRowNum, 1, 2));
		this.highlightedFonts = ApachePOIWrapper.getMappedFontColor(sheet,
				new CellRangeAddress(rowNum, lastRowNum, 3, 3));
		this.setIgnorePrefixes();
		this.setHolidayLocale();
		this.setQuarterStartWeeks();
		this.setExamWeekLength();
		this.setQuarterStartDate();
		this.setExamWeekTextAndStyle();
	}

	/**
	 * Adds all new lecture names of the given list in the zero based row 0 in the
	 * first sheet of the workbook.
	 * 
	 * All lecture names starts with the {@link ConfigWorkbook#ignorePrefixes} are
	 * added without the prefix. If there exists a lecture name with this new name,
	 * it is not added.
	 * 
	 * @param lectureNames List of lecture names to add to the workbook
	 * @throws IOException If saving the workbook file failed
	 */
	public void addLectureNames(List<String> lectureNames) throws IOException {
		int rowNum = 2;
		XSSFSheet sheet = ApachePOIWrapper.getSheet(this.getWorkbook());
		int lastRowNum = sheet.getLastRowNum();
		String[] namesToColor = ApachePOIWrapper.getStringValuesFromWorkbook(sheet,
				new CellRangeAddress(rowNum, lastRowNum, 1, 1));
		this.setIgnorePrefixes();

		for (String lectureName : lectureNames) {
			String rawLectureName = LectureWorkbook.removePrefixFromString(lectureName, this.getIgnorePrefixes());
			if (!ConfigWorkbook.arrayContainsWithWildcard(rawLectureName, namesToColor)
					&& (rawLectureName.equals(lectureName) || !lectureNames.contains(rawLectureName))) {

				rowNum = ConfigWorkbook.addValueToNextEmptyCellInARow(sheet, rawLectureName, rowNum, 1);

				rowNum++;
			}
		}
		lastRowNum = sheet.getLastRowNum();
		ApachePOIWrapper.saveWorkbookToFile(this.getWorkbook(), this.getFile());
		this.lecturePropertiesMap = ConfigWorkbook.getMappedLectureProperties(sheet,
				new CellRangeAddress(2, lastRowNum, 1, 2));
	}

	/**
	 * Returns a map of color pairs from a cell range. All cells with a value in the
	 * cell range will be added to the map with their color pair. A color pair is an
	 * array with to values. The first value is the font color of the cell and the
	 * second value is the fill color of the cell.
	 * 
	 * @param sheet     The sheet, which will be scanned
	 * @param cellRange The cell range, where first column will be used for the
	 *                  lecture name and the colors and the last column will be used
	 *                  for the short lecture name
	 * @return A map of cell values and color array pairs (array always and only
	 *         contains font color and fill color)
	 */
	public static Map<String, LectureProperties> getMappedLectureProperties(XSSFSheet sheet,
			CellRangeAddress cellRange) {
		Map<String, LectureProperties> lecturePropertiesMap = new HashMap<String, LectureProperties>();
		int firstColumnNum = cellRange.getFirstColumn();
		int lastColumnNum = cellRange.getLastColumn();

		for (int rowNum = cellRange.getFirstColumn(); rowNum <= cellRange.getLastRow(); rowNum++) {
			XSSFRow row = sheet.getRow(rowNum);
			XSSFCell firstCell = row.getCell(firstColumnNum);
			if (firstCell != null && firstCell.getCellType() == CellType.STRING) {
				String lectureName = firstCell.getStringCellValue();
				if (lectureName != null && lectureName != "") {
					XSSFCellStyle cellStyle = firstCell.getCellStyle();
					XSSFColor fillColor = cellStyle.getFillForegroundColorColor();
					fillColor = fillColor == null ? cellStyle.getFillBackgroundColorColor() : fillColor;
					XSSFCell lastCell = row.getCell(lastColumnNum);

					String shortLectureName = lastCell != null && lastCell.getCellType() == CellType.STRING
							? lastCell.getStringCellValue()
							: "";
					shortLectureName = shortLectureName == null ? "" : shortLectureName;

					LectureProperties lectureProperties = new LectureProperties(lectureName, shortLectureName,
							cellStyle.getFont().getXSSFColor(), fillColor);
					lecturePropertiesMap.put(lectureName, lectureProperties);
				}
			}
		}
		return lecturePropertiesMap;
		// TODO move to ApachePOIWrapper? Generalize variable names
	}

	/**
	 * Adds a value in the given sheet in the given column in the first cell, which
	 * is empty or is not a string value.
	 * 
	 * @param sheet       The sheet, where the value is added to
	 * @param value       The value to add
	 * @param startRowNum The first possible row, where the value is added to
	 * @param columnNum   The column, where the value is added to
	 * @return The row number, where the value was inserted
	 */
	public static int addValueToNextEmptyCellInARow(XSSFSheet sheet, String value, int startRowNum, int columnNum) {
		boolean added = false;
		while (!added) {
			XSSFRow row = sheet.getRow(startRowNum);
			row = row == null ? sheet.createRow(startRowNum) : row;
			XSSFCell cell = row.getCell(columnNum);
			cell = cell == null ? row.createCell(columnNum) : cell;
			if (cell.getCellType() != CellType.STRING || cell.getStringCellValue() == "") {
				cell.setCellValue(value);
				added = true;
			} else {
				startRowNum++;
			}
		}
		return startRowNum;
	}

	/**
	 * Returns True if the name matches one of the strings in the matchers array,
	 * otherwise false. The matchers strings can contain wildcards ('*').
	 * 
	 * @param name     The name to check for possible matches
	 * @param matchers The array of matchers
	 * @return True if the name matches one of the matchers, otherwise false
	 */
	public static boolean arrayContainsWithWildcard(String name, String[] matchers) {
		boolean matches = false;
		for (String matcher : matchers) {
			String matcherKey = "\\Q" + matcher.replace("*", "\\E.*\\Q") + "\\E";
			if (name.matches(matcherKey)) {
				matches = true;
				break;
			}
		}
		return matches;
		// TODO move method to helper class
	}
}
