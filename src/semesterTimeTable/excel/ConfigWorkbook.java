package semesterTimeTable.excel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
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

	private Locale holidayLocale;

	private int[] quarterStartWeeks;

	private int examWeekLength;

	private Calendar quarterStartDate;

	private XSSFRichTextString examWeekText;

	private XSSFFont examWeekFont;

	private XSSFColor examWeekFillColor;

	private XSSFWorkbook workbook;

	private File file;

	private boolean isNewConfig;

	/**
	 * Loads the style configuration of the configuration template.
	 * 
	 * @throws IOException If loading the template file fails
	 */
	public ConfigWorkbook() throws IOException {
		this.setWorkbook(LectureWorkbook
				.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ConfigWorkbook.TEMPLATE_FILENAME)));
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
			this.setWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
			this.initConfigWorkbook();
			this.setIsNewConfig(false);
		} else {
			this.setWorkbook(LectureWorkbook
					.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ConfigWorkbook.TEMPLATE_FILENAME)));
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
			this.setWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
			this.initConfigWorkbook();
			this.setIsNewConfig(false);
		} else {
			this.setWorkbook(LectureWorkbook
					.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ConfigWorkbook.TEMPLATE_FILENAME)));
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
			this.setWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
			this.initConfigWorkbook();
			this.setIsNewConfig(false);
		} else {
			this.setWorkbook(LectureWorkbook
					.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ConfigWorkbook.TEMPLATE_FILENAME)));
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
			this.setWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
			this.initConfigWorkbook();
			this.setIsNewConfig(false);
		} else {
			this.setWorkbook(LectureWorkbook
					.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ConfigWorkbook.TEMPLATE_FILENAME)));
			this.initConfigWorkbook(file, lectureNames);
			this.setIsNewConfig(true);
		}
	}

	private XSSFWorkbook getWorkbook() {
		return this.workbook;
	}

	private void setWorkbook(XSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	public boolean isNewConfig() {
		return this.isNewConfig;
	}

	private void setIsNewConfig(boolean isNewConfig) {
		this.isNewConfig = isNewConfig;
	}

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

	private void setIgnorePrefixes(XSSFSheet sheet) {
		int lastRowNum = sheet.getLastRowNum();
		this.ignorePrefixes = ConfigWorkbook.getStringValuesFromWorkbook(sheet,
				new CellRangeAddress(2, lastRowNum, 3, 3));
	}

	public Locale getHolidayLocale() {
		return this.holidayLocale;
	}

	private void setHolidayLocale(XSSFSheet sheet) {
		int lastRowNum = sheet.getLastRowNum();
		String[] localeValues = ConfigWorkbook.getStringValuesFromWorkbook(sheet,
				new CellRangeAddress(2, lastRowNum, 4, 4));
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

	public int[] getQuarterStartWeeks() {
		return this.quarterStartWeeks;
	}

	private void setQuarterStartWeeks(XSSFSheet sheet) {
		int lastRowNum = sheet.getLastRowNum();
		int[] quarterStartWeeks = ConfigWorkbook.getIntegerValuesFromWorkbook(sheet,
				new CellRangeAddress(2, lastRowNum, 5, 5));
		if (quarterStartWeeks.length < 1) {
			quarterStartWeeks = new int[] { 2, 15, 27, 40 };
		}
		this.quarterStartWeeks = quarterStartWeeks;
	}

	public int getExamWeekLength() {
		return this.examWeekLength;
	}

	private void setExamWeekLength(XSSFSheet sheet) {
		int[] lengths = ConfigWorkbook.getIntegerValuesFromWorkbook(sheet, new CellRangeAddress(2, 2, 6, 6));
		if (lengths.length == 1) {
			if (lengths[0] < 0) {
				this.examWeekLength = 0;
			} else if(lengths[0] > 10) {
				this.examWeekLength = 10;
			} else {
				this.examWeekLength = lengths[0];
			}
		} else {
			this.examWeekLength = 6;
		}
	}

	public Calendar getQuarterStartDate() {
		return this.quarterStartDate;
	}

	private void setQuarterStartDate(XSSFSheet sheet) {
		Date[] dates = ConfigWorkbook.getDateValuesFromWorkbook(sheet, new CellRangeAddress(2, 2, 7, 7));
		String[] timeZoneCodes = ConfigWorkbook.getStringValuesFromWorkbook(sheet, new CellRangeAddress(3, 3, 7, 7));

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

	public XSSFRichTextString getExamWeekText() {
		return this.examWeekText;
	}

	public XSSFFont getExamWeekFont() {
		return this.examWeekFont;
	}

	public XSSFColor getExamWeekFillColor() {
		return this.examWeekFillColor;
	}

	private void setExamWeekTextAndStyle(XSSFSheet sheet) {
		XSSFRow row = sheet.getRow(2);
		if (row != null) {
			XSSFCell cell = row.getCell(8);
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

	public void close() throws IOException {
		this.getWorkbook().close();
	}

	private File getFile() {
		return this.file;
	}

	private void setFile(File file) {
		this.file = file;
	}

	/**
	 * Creates a new configuration file with the configuration template and the
	 * given lectures and sets all the configurable variables.
	 * 
	 * All lecture names starts with the {@link ConfigWorkbook#REMOVE_PREFIX} are
	 * not added to the configuration file.
	 * 
	 * All lecture names starts with the {@link ConfigWorkbook#ignorePrefixes} are
	 * added without the prefix. If there exists a lecture name with this new name,
	 * it is not added.
	 * 
	 * @param workbook     The template configuration workbook
	 * @param file         The file to save the configuration workbook into
	 * @param lectureNames A list of lecture names
	 * @throws IOException If saving the configuration workbook or closing the
	 *                     workbook failed
	 */
	private void initConfigWorkbook(File file, List<String> lectureNames) throws IOException {
		this.addLectureNames(lectureNames);
		this.initConfigWorkbook(file);
	}

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
				new CellRangeAddress(rowNum, lastRowNum, 0, 1));
		this.highlightedFonts = ConfigWorkbook.getMappedFontColor(sheet,
				new CellRangeAddress(rowNum, lastRowNum, 2, 2));
		this.setIgnorePrefixes(sheet);
		this.setHolidayLocale(sheet);
		this.setQuarterStartWeeks(sheet);
		this.setExamWeekLength(sheet);
		this.setQuarterStartDate(sheet);
		this.setExamWeekTextAndStyle(sheet);
	}

	public void addLectureNames(List<String> lectureNames) throws IOException {
		int rowNum = 2;
		XSSFSheet sheet = this.getWorkbook().getSheetAt(0);
		int lastRowNum = sheet.getLastRowNum();
		String[] namesToColor = ConfigWorkbook.getStringValuesFromWorkbook(sheet,
				new CellRangeAddress(rowNum, lastRowNum, 0, 0));
		this.setIgnorePrefixes(sheet);

		for (String lectureName : lectureNames) {
			String rawLectureName = LectureWorkbook.removePrefixFromString(lectureName, this.getIgnorePrefixes());
			if (!ConfigWorkbook.arrayContainsWithWildcard(rawLectureName, namesToColor)
					&& (rawLectureName.equals(lectureName) || !lectureNames.contains(rawLectureName))) {

				rowNum = ConfigWorkbook.addValueToNextEmptyCellInARow(sheet, rawLectureName, rowNum, 0);

				rowNum++;
			}
		}
		lastRowNum = sheet.getLastRowNum();
		ApachePOIWrapper.saveWorkbookToFile(this.getWorkbook(), this.getFile());
		this.lecturePropertiesMap = ConfigWorkbook.getMappedLectureProperties(sheet,
				new CellRangeAddress(2, lastRowNum, 0, 1));
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
	}

	/**
	 * Returns a map of font colors from a cell range. All cells with a value in the
	 * cell range will be added to the map with their font color.
	 * 
	 * @param sheet     The sheet, which will be scanned
	 * @param cellRange The cell range, which will be scanned
	 * @return A map of cell values and font color pairs
	 */
	public static Map<String, XSSFFont> getMappedFontColor(XSSFSheet sheet, CellRangeAddress cellRange) {
		Map<String, XSSFFont> fontMap = new HashMap<String, XSSFFont>();
		for (int rowNum = cellRange.getFirstRow(); rowNum <= cellRange.getLastRow(); rowNum++) {
			XSSFRow row = sheet.getRow(rowNum);
			for (int columnNum = cellRange.getFirstColumn(); columnNum <= cellRange.getLastColumn(); columnNum++) {
				XSSFCell cell = row.getCell(columnNum);
				if (cell != null) {
					String key = cell.getStringCellValue();
					if (key != null && key != "") {
						XSSFFont font = new XSSFFont();
						font.setFontHeight((short) 200);
						font.setFontName("Arial");
						font.setColor(cell.getCellStyle().getFont().getXSSFColor());
						fontMap.put(key, font);
					}
				}
			}

		}
		return fontMap;
	}

	/**
	 * Returns an array of all values from a cell range. Cells with no value will
	 * not be added to the array.
	 * 
	 * @param sheet     The sheet, which will be scanned
	 * @param cellRange The cell range, which will be scanned
	 * @return An array of cell values
	 */
	public static String[] getStringValuesFromWorkbook(XSSFSheet sheet, CellRangeAddress cellRange) {
		List<String> valueList = new ArrayList<String>();
		for (int rowNum = cellRange.getFirstRow(); rowNum <= cellRange.getLastRow(); rowNum++) {
			XSSFRow row = sheet.getRow(rowNum);
			for (int columnNum = cellRange.getFirstColumn(); columnNum <= cellRange.getLastColumn(); columnNum++) {
				XSSFCell cell = row.getCell(columnNum);
				if (cell != null && cell.getCellType() == CellType.STRING) {
					String value = cell.getStringCellValue();
					if (value != null && value != "") {
						valueList.add(value);
					}
				}
			}
		}
		return valueList.toArray(new String[valueList.size()]);
	}

	public static int[] getIntegerValuesFromWorkbook(XSSFSheet sheet, CellRangeAddress cellRange) {
		List<Integer> valueList = new ArrayList<Integer>();
		for (int rowNum = cellRange.getFirstRow(); rowNum <= cellRange.getLastRow(); rowNum++) {
			XSSFRow row = sheet.getRow(rowNum);
			for (int columnNum = cellRange.getFirstColumn(); columnNum <= cellRange.getLastColumn(); columnNum++) {
				XSSFCell cell = row.getCell(columnNum);
				if (cell != null) {
					if (cell.getCellType() == CellType.NUMERIC) {
						int value = (int) cell.getNumericCellValue();
						valueList.add(value);
					} else if (cell.getCellType() == CellType.STRING) {
						try {
							int value = (int) Double.parseDouble(cell.getStringCellValue());
							valueList.add(value);
						} catch (NumberFormatException e) {

						}
					}
				}
			}
		}
		int[] valueArray = new int[valueList.size()];
		for (int i = 0; i < valueArray.length; i++) {
			valueArray[i] = valueList.get(i);
		}
		return valueArray;
	}

	public static Date[] getDateValuesFromWorkbook(XSSFSheet sheet, CellRangeAddress cellRange) {
		List<Date> valueList = new ArrayList<Date>();
		for (int rowNum = cellRange.getFirstRow(); rowNum <= cellRange.getLastRow(); rowNum++) {
			XSSFRow row = sheet.getRow(rowNum);
			for (int columnNum = cellRange.getFirstColumn(); columnNum <= cellRange.getLastColumn(); columnNum++) {
				XSSFCell cell = row.getCell(columnNum);
				if (cell != null && cell.getCellType() == CellType.NUMERIC) {
					Date value = cell.getDateCellValue();
					if (value != null) {
						valueList.add(value);
					}
				}
			}
		}
		return valueList.toArray(new Date[valueList.size()]);
	}

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
	}
}
