package semesterTimeTable.excel;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TimeZone;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class representing a workbook, which contains lectures.
 *
 */
public class LectureWorkbook {

	/** List of normal start times of a lecture */
	public final static List<String> START_TIMES = new ArrayList<String>(Arrays.asList("00:00", "08:00", "08:45",
			"09:45", "10:30", "11:30", "12:15", "14:00", "14:45", "15:45", "16:30", "17:30", "18:15"));
	/** List of normal end times of a lecture */
	public final static List<String> END_TIMES = new ArrayList<String>(Arrays.asList("00:00", "08:45", "09:30", "10:30",
			"11:15", "12:15", "13:00", "14:45", "15:30", "16:30", "17:15", "18:15", "19:00"));

	/** Group name of all holidays */
	public final static String HOLIDAY = "holiday";

	/** Name of the template workbook file */
	private final static String TEMPLATE_FILENAME = "template.xlsx";

	/** String representing a line break inside a workbook cell */
	public final static String LINE_BREAK = "\n";

	/** Error Output for printing errors */
	private ErrorOutput errorOutput;

	/**
	 * Object containing style (especially color) configurations from a workbook for
	 * the lectures
	 */
	private ConfigWorkbook configWorkbook;

	/** Workbook for the lectures */
	private XSSFWorkbook workbook;

	/** Included start date of the quarter */
	private Calendar quarterStartDate;

	/** Excluded end date of the quarter */
	private Calendar quarterEndDate;

	/** Map of lectures grouped by their names */
	private Map<String, List<Lecture>> groupedLectures;

	/**
	 * Lists of columns for different types of left/right borders in last block
	 * expect the exam week. For more details see {@link #setBorderLists()}
	 */
	private List<List<Integer>> borderColumnsLastBlock;
	/**
	 * Lists of columns for different types of left/right borders in the exam week.
	 * For more details see {@link #setBorderLists()}
	 */
	private List<List<Integer>> borderColumnsExamWeek;
	/**
	 * Lists of columns for different types of left/right borders in first two
	 * blocks. For more details see {@link #setBorderLists()}
	 */
	private List<List<Integer>> borderColumns;
	/**
	 * Lists of rows for different types of top borders. For more details see
	 * {@link #setBorderLists()}
	 */
	private List<List<Integer>> borderRows;

	/**
	 * 3x3 Matrix of border Styles with white fill color for all empty cells in
	 * lecture area
	 * <table>
	 * <tr>
	 * <td>left border</td>
	 * <td>no border</td>
	 * <td>right border</td>
	 * </tr>
	 * <tr>
	 * <td>left border + top gray border</td>
	 * <td>top gray border</td>
	 * <td>right border + top gray border</td>
	 * </tr>
	 * <tr>
	 * <td>left border + top black border</td>
	 * <td>top black border</td>
	 * <td>right border + top black border</td>
	 * </tr>
	 * </table>
	 */
	private XSSFCellStyle[][] borderStyle;
	/**
	 * Same border Styles like {@link #borderColumns #borderStyle} but with other
	 * fill color for empty cells in exam week area
	 */
	private XSSFCellStyle[][] borderStyleExamWeek;

	/**
	 * Loads the workbook of the given filename, if the file exists. If not a
	 * workbook template will be loaded. Insert the given lectures into the workbook
	 * by using the styles of the colorWorkbook. If there is an excel file called
	 * colorMap.xlsx in the directory of the file, it will be loaded as custom
	 * colorWorkbook.
	 * 
	 * @param filename The path to the workbook file
	 * @throws IOException If reading one workbook file failed
	 */
	public LectureWorkbook(String filename) throws IOException {
		this.errorOutput = new ErrorOutput();
		File file = new File(filename);
		this.setConfigWorkbook(new ConfigWorkbook(file.getParent()));
		if (file.exists()) {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromFile(file));
		} else {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromInputStream(
					LectureWorkbook.getTemplateInputStream(LectureWorkbook.TEMPLATE_FILENAME)));
		}
		this.setBorderLists();
		this.createBorderStyles();
		Calendar quarterStartDate = this.getConfigWorkbook().getQuarterStartDate();
		if (quarterStartDate != null) {
			this.setBorderDatesWithDateInFirstWeek(quarterStartDate);
		}
	}

	/**
	 * Loads the workbook of the given filename, if the file exists. If not a
	 * workbook template will be loaded. Insert the given lectures into the workbook
	 * by using the styles of the colorWorkbook. If there is an excel file called
	 * colorMap.xlsx in the directory of the file, it will be loaded as custom
	 * colorWorkbook.
	 * 
	 * @param filename    The path to the workbook file
	 * @param errorOutput The object for error outputs
	 * @throws IOException If reading one workbook file failed
	 */
	public LectureWorkbook(String filename, Output output) throws IOException {
		this.errorOutput = new ErrorOutput(output);
		File file = new File(filename);
		this.setConfigWorkbook(new ConfigWorkbook(file.getParent()));
		if (file.exists()) {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromFile(file));
		} else {
			this.setWorkbook(ApachePOIWrapper.loadWorkbookFromInputStream(
					LectureWorkbook.getTemplateInputStream(LectureWorkbook.TEMPLATE_FILENAME)));
		}
		this.setBorderLists();
		this.createBorderStyles();
		Calendar quarterStartDate = this.getConfigWorkbook().getQuarterStartDate();
		if (quarterStartDate != null) {
			this.setBorderDatesWithDateInFirstWeek(quarterStartDate);
		}
	}

	/**
	 * Returns the error output object.
	 * 
	 * You can use the error output as a kind of error logging object. If for
	 * example a lecture cannot be added to the excel workbook, the error message
	 * will be added to the error output.
	 * 
	 * @return The error output object
	 */
	public ErrorOutput getErrorOutput() {
		return this.errorOutput;
	}

	/**
	 * Returns the workbook containing the color styles for the lectures.
	 * 
	 * @return The color workbook
	 */
	public ConfigWorkbook getConfigWorkbook() {
		return this.configWorkbook;
	}

	/**
	 * Sets the workbook containing the color styles for the lectures.
	 * 
	 * @param colorWorkbook The workbook with the color styles
	 */
	private void setConfigWorkbook(ConfigWorkbook colorWorkbook) {
		this.configWorkbook = colorWorkbook;
	}

	/**
	 * Returns the lecture workbook.
	 * 
	 * @return The workbook
	 */
	public XSSFWorkbook getWorkbook() {
		return this.workbook;
	}

	/**
	 * Sets the lecture workbook. The workbook is used as a template and all
	 * lectures inside the template could be removed.
	 * 
	 * @param workbook The workbook for inserting lectures
	 */
	private void setWorkbook(XSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	/**
	 * Deletes all lectures in the workbook and insert the lectures from
	 * groupedLectures into the workbook.
	 */
	public void fillWorkbook() {
		XSSFSheet sheet = ApachePOIWrapper.getSheet(this.getWorkbook());
		this.resetLectureAreaInWorkbook();

		ConfigWorkbook configWorkbook = this.getConfigWorkbook();
		int firstColumn = 22 - configWorkbook.getExamWeekLength();
		if (firstColumn < 22) {
			sheet.addMergedRegion(new CellRangeAddress(139, 146, firstColumn, 21));
			XSSFCell cell = sheet.getRow(139).getCell(firstColumn);
			XSSFWorkbook workbook = this.getWorkbook();

			XSSFCellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setFillForegroundColor(configWorkbook.getExamWeekFillColor());
			cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cellStyle.setWrapText(true);
			cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			cellStyle.setAlignment(HorizontalAlignment.CENTER);

			XSSFFont font = workbook.createFont();
			XSSFFont examWeekFont = configWorkbook.getExamWeekFont();
			ApachePOIWrapper.copyFont(font, examWeekFont);
			cellStyle.setFont(font);

			cell.setCellStyle(cellStyle);
			cell.setCellValue(configWorkbook.getExamWeekText());
		}

		this.addLecturesToWorkbook();

		sheet.getRow(2).getCell(1).setCellValue(this.getQuarterStartDate());
		sheet.getRow(1).getCell(0).setCellValue(new GregorianCalendar());
		XSSFFormulaEvaluator.evaluateAllFormulaCells(this.getWorkbook());
	}

	/**
	 * Checks if the cell at the given position is part of the lecture area.
	 * 
	 * @param rowNum    The (0 based) row number of the cell
	 * @param columnNum The (0 based) column number of the cell
	 * @return true if cell is part of the lecture area, false otherwise
	 */
	private boolean isLectureCell(int rowNum, int columnNum) {
		return ((rowNum > 2 && rowNum < 49) || (rowNum > 51 && rowNum < 98) || (rowNum > 100 && rowNum < 147))
				&& (columnNum > 0 && columnNum != 11 && columnNum < 22);
	}

	/**
	 * Creates lists with row numbers and their top border style and with column
	 * numbers and their left/right border style.
	 * 
	 * Each border columns list ({@link #borderColumnsLastBlock},
	 * {@link #borderColumnsExamWeek}, {@link #borderColumns}) contains three inner
	 * lists. Each inner list contains column numbers for different left/right
	 * border types.
	 * 
	 * <table>
	 * <tr>
	 * <th>list index</th>
	 * <th>list content</th>
	 * </tr>
	 * <tr>
	 * <td>0</td>
	 * <td>column numbers with left border</td>
	 * </tr>
	 * <tr>
	 * <td>1</td>
	 * <td>column numbers with no left or right border</td>
	 * </tr>
	 * <tr>
	 * <td>2</td>
	 * <td>column numbers with right border</td>
	 * </tr>
	 * </table>
	 * 
	 * Each border rows list ({@link #borderRows}) contains three lists. Each inner
	 * list contains row numbers for different top border types.
	 * 
	 * <table>
	 * <tr>
	 * <th>list index</th>
	 * <th>list content</th>
	 * </tr>
	 * <tr>
	 * <td>0</td>
	 * <td>rows with no top border</td>
	 * </tr>
	 * <tr>
	 * <td>1</td>
	 * <td>rows with gray top border</td>
	 * </tr>
	 * <tr>
	 * <td>2</td>
	 * <td>rows with black top border</td>
	 * </tr>
	 * </table>
	 * 
	 */
	private void setBorderLists() {
		int firstExamWeekColumn = 22 - this.getConfigWorkbook().getExamWeekLength();

		List<Integer> leftBorderColumns = Arrays.asList(1, 6, 12, 17);
		List<Integer> noBorderColumns = Arrays.asList(2, 3, 4, 7, 8, 9, 13, 14, 15, 18, 19, 20);
		List<Integer> rightBorderColumns = Arrays.asList(5, 10, 16, 21);

		this.borderColumnsLastBlock = new ArrayList<List<Integer>>(3);
		this.borderColumnsExamWeek = new ArrayList<List<Integer>>(3);
		this.borderColumns = new ArrayList<List<Integer>>(3);

		this.borderColumnsLastBlock.add(0, LectureWorkbook.getSubColumnList(leftBorderColumns, 0, firstExamWeekColumn));
		this.borderColumnsExamWeek.add(0, LectureWorkbook.getSubColumnList(leftBorderColumns, firstExamWeekColumn, 22));
		this.borderColumns.add(0, leftBorderColumns);

		this.borderColumnsLastBlock.add(1, LectureWorkbook.getSubColumnList(noBorderColumns, 0, firstExamWeekColumn));
		this.borderColumnsExamWeek.add(1, LectureWorkbook.getSubColumnList(noBorderColumns, firstExamWeekColumn, 22));
		this.borderColumns.add(1, noBorderColumns);

		this.borderColumnsLastBlock.add(2,
				LectureWorkbook.getSubColumnList(rightBorderColumns, 0, firstExamWeekColumn));
		this.borderColumnsExamWeek.add(2,
				LectureWorkbook.getSubColumnList(rightBorderColumns, firstExamWeekColumn, 22));
		this.borderColumns.add(2, rightBorderColumns);

		this.borderRows = new ArrayList<List<Integer>>(3);
		this.borderRows.add(0, Arrays.asList(5, 6, 8, 9, 12, 13, 15, 16, 19, 20, 22, 23, 25, 26, 27, 29, 30, 32, 33, 36,
				37, 39, 40, 43, 44, 46, 47));
		this.borderRows.add(1, Arrays.asList(4, 7, 10, 11, 14, 17, 18, 21, 31, 34, 35, 38, 41, 42, 45, 48));
		this.borderRows.add(2, Arrays.asList(3, 24, 28));
	}

	private static List<Integer> getSubColumnList(List<Integer> columns, int startColumn, int endColumn) {
		List<Integer> subColumns = new ArrayList<Integer>();
		for (int column : columns) {
			if (column >= startColumn && column < endColumn) {
				subColumns.add(column);
			}
		}
		return subColumns;
	}

	/**
	 * Creates border styles for each kind of cell.
	 */
	private void createBorderStyles() {

		XSSFWorkbook workbook = this.getWorkbook();

		// leftNoTop middleNoTop rightNoTop
		// leftGrey middleGrey rightGrey
		// leftBlack middleBlack rightBlack
		this.borderStyle = new XSSFCellStyle[3][3];

		this.borderStyle[0][1] = workbook.createCellStyle();
		this.borderStyle[0][1].setBorderLeft(BorderStyle.THIN);
		this.borderStyle[0][1].setBorderRight(BorderStyle.THIN);
		this.borderStyle[0][1].setFillForegroundColor(ApachePOIWrapper.colorToXSSFColor(Color.WHITE));
		this.borderStyle[0][1].setFillPattern(FillPatternType.SOLID_FOREGROUND);

		this.borderStyle[0][0] = workbook.createCellStyle();
		this.borderStyle[0][0].cloneStyleFrom(this.borderStyle[0][1]);
		this.borderStyle[0][0].setBorderLeft(BorderStyle.THICK);

		this.borderStyle[0][2] = workbook.createCellStyle();
		this.borderStyle[0][2].cloneStyleFrom(this.borderStyle[0][1]);
		this.borderStyle[0][2].setBorderRight(BorderStyle.THICK);

		this.borderStyle[2][1] = workbook.createCellStyle();
		this.borderStyle[2][1].cloneStyleFrom(this.borderStyle[0][1]);
		this.borderStyle[2][1].setBorderTop(BorderStyle.THIN);

		this.borderStyle[2][0] = workbook.createCellStyle();
		this.borderStyle[2][0].cloneStyleFrom(this.borderStyle[2][1]);
		this.borderStyle[2][0].setBorderLeft(BorderStyle.THICK);

		this.borderStyle[2][2] = workbook.createCellStyle();
		this.borderStyle[2][2].cloneStyleFrom(this.borderStyle[2][1]);
		this.borderStyle[2][2].setBorderRight(BorderStyle.THICK);

		this.borderStyle[1][1] = workbook.createCellStyle();
		this.borderStyle[1][1].cloneStyleFrom(this.borderStyle[2][1]);
		this.borderStyle[1][1].setTopBorderColor(ApachePOIWrapper.colorToXSSFColor(Color.LIGHT_GRAY));

		this.borderStyle[1][0] = workbook.createCellStyle();
		this.borderStyle[1][0].cloneStyleFrom(this.borderStyle[1][1]);
		this.borderStyle[1][0].setBorderLeft(BorderStyle.THICK);

		this.borderStyle[1][2] = workbook.createCellStyle();
		this.borderStyle[1][2].cloneStyleFrom(this.borderStyle[1][1]);
		this.borderStyle[1][2].setBorderRight(BorderStyle.THICK);

		this.borderStyleExamWeek = new XSSFCellStyle[3][3];
		for (int kindOfRow = 0; kindOfRow < 3; kindOfRow++) {
			for (int kindOfColumn = 0; kindOfColumn < 3; kindOfColumn++) {
				this.borderStyleExamWeek[kindOfRow][kindOfColumn] = workbook.createCellStyle();
				this.borderStyleExamWeek[kindOfRow][kindOfColumn]
						.cloneStyleFrom(this.borderStyle[kindOfRow][kindOfColumn]);
				this.borderStyleExamWeek[kindOfRow][kindOfColumn]
						.setFillForegroundColor(ApachePOIWrapper.colorToXSSFColor(Color.CYAN));
				this.borderStyleExamWeek[kindOfRow][kindOfColumn].setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}
		}
		for (int kindOfColumn = 0; kindOfColumn < 3; kindOfColumn++) {
			this.borderStyleExamWeek[1][kindOfColumn].setTopBorderColor(ApachePOIWrapper.colorToXSSFColor(Color.BLACK));
		}
	}

	/**
	 * Resets each cell in the lecture area in the sheet of the workbook by setting
	 * the cells blank and adding the correct cellStyle.
	 */
	private void resetLectureAreaInWorkbook() {
		XSSFSheet sheet = ApachePOIWrapper.getSheet(this.getWorkbook());
		List<CellRangeAddress> ranges = sheet.getMergedRegions();
		for (int index = sheet.getNumMergedRegions() - 1; index > -1; index--) {
			CellRangeAddress range = ranges.get(index);
			if (this.isLectureCell(range.getFirstRow(), range.getFirstColumn())) {
				sheet.removeMergedRegion(index);
			}
		}

		// First two blocks
		for (int block = 0; block < 2; block++) {
			int blockStartRow = block * 49;
			int kindOfRowIndex = 0;
			for (List<Integer> kindOfRow : this.borderRows) {
				for (int rawRowNum : kindOfRow) {
					int rowNum = rawRowNum + blockStartRow;
					int kindOfColumnIndex = 0;
					for (List<Integer> kindOfColumn : this.borderColumns) {
						for (int columnNum : kindOfColumn) {
							ApachePOIWrapper.resetCell(this.getWorkbook(), rowNum, columnNum,
									this.borderStyle[kindOfRowIndex][kindOfColumnIndex]);
						}
						kindOfColumnIndex++;
					}
				}
				kindOfRowIndex++;
			}
		}

		// Last block
		int blockStartRow = 2 * 49;
		int kindOfRowIndex = 0;
		for (List<Integer> kindOfRow : this.borderRows) {
			for (int rawRowNum : kindOfRow) {
				int rowNum = rawRowNum + blockStartRow;

				// Before Exam week
				int kindOfColumnIndex = 0;
				for (List<Integer> kindOfColumn : this.borderColumnsLastBlock) {
					for (int columnNum : kindOfColumn) {
						ApachePOIWrapper.resetCell(this.getWorkbook(), rowNum, columnNum,
								this.borderStyle[kindOfRowIndex][kindOfColumnIndex]);
					}
					kindOfColumnIndex++;
				}

				// Exam week
				kindOfColumnIndex = 0;
				for (List<Integer> kindOfColumn : this.borderColumnsExamWeek) {
					for (int columnNum : kindOfColumn) {
						ApachePOIWrapper.resetCell(this.getWorkbook(), rowNum, columnNum,
								this.borderStyleExamWeek[kindOfRowIndex][kindOfColumnIndex]);
					}
					kindOfColumnIndex++;
				}
			}
			kindOfRowIndex++;
		}
	}

	/**
	 * Returns the included start date of the quarter. It can cause layout errors in
	 * the workbook, if the day of the date is not a Monday.
	 * 
	 * @return The included start date of the quarter
	 */
	public Calendar getQuarterStartDate() {
		return this.quarterStartDate;
	}

	/**
	 * Returns the excluded end date of the quarter. It can cause layout errors in
	 * the workbook, if the day of the date is not a Saturday.
	 * 
	 * @return The excluded end date of the quarter
	 */
	public Calendar getQuarterEndDate() {
		return this.quarterEndDate;
	}

	/**
	 * Calculates and sets the start and end date of the quarter from the given week
	 * of year and year.
	 * 
	 * @param weekOfYear The start week of the quarter
	 * @param year       The year of the quarter
	 * @param timeZone   The time zone for the start and end date
	 */
	private void setBorderDates(int weekOfYear, int year, TimeZone timeZone) {
		this.quarterStartDate = LectureWorkbook.weekOfYearToDate(weekOfYear, Calendar.MONDAY, year, timeZone);
		this.quarterEndDate = LectureWorkbook.weekOfYearToDate(weekOfYear + 11, Calendar.SATURDAY, year, timeZone);
		this.addHolidays();
	}

	/**
	 * Calculates and sets the start and end date of the quarter from the given
	 * date. The date can be any date in the first week of the quarter.
	 * 
	 * @param date Any date in the first week of the quarter
	 */
	public void setBorderDatesWithDateInFirstWeek(Calendar date) {
		this.setBorderDates(date.get(Calendar.WEEK_OF_YEAR), date.get(Calendar.YEAR), date.getTimeZone());
	}

	/**
	 * Calculates and sets the start and end date of the quarter from the given
	 * date. The date can be any date in the quarter. The quarter start weeks of the
	 * configuration file are used to get the first week of the quarter for the
	 * given date.
	 * 
	 * @param date Any date in the quarter
	 */
	public void setBorderDatesWithDateInQuarter(Calendar date) {
		int[] quarterStartWeeks = this.getConfigWorkbook().getQuarterStartWeeks();
		int week = date.get(Calendar.WEEK_OF_YEAR);
		for (int quarterStartWeek : quarterStartWeeks) {
			if (week >= quarterStartWeek && week < quarterStartWeek + 13) {
				this.setBorderDates(quarterStartWeek, date.get(Calendar.YEAR), date.getTimeZone());
				break;
			}
		}
	}

	/**
	 * Returns the lectures grouped by their name.
	 * 
	 * @return The map of grouped lectures
	 */
	public Map<String, List<Lecture>> getGroupedLectures() {
		return this.groupedLectures;
	}

	/**
	 * Sets the grouped lectures by grouping the given lecture list without. This
	 * method only sets the grouped lectures and do not change the workbook.
	 * 
	 * @param lectures A list of lectures
	 * @throws IOException
	 */
	public void setLectures(List<Lecture> lectures) throws IOException {
		this.groupedLectures = new TreeMap<String, List<Lecture>>(
				lectures.stream().collect(Collectors.groupingBy(Lecture::getName)));
		this.addHolidays();
	}

	/**
	 * Returns a list of all grouped lecture names. The list contains each lecture
	 * name only once. There are no duplicates in the list.
	 * 
	 * @return The list of all lectures without duplicates
	 */
	public List<String> getLectures() {
		Set<String> keys = this.getGroupedLectures().keySet();
		return Arrays.asList(keys.toArray(new String[keys.size()]));
	}

	/**
	 * Saves the workbook as an xlsx file with the given filename.
	 * 
	 * @param filename The name for the xlsx file
	 * @throws IOException If saving the workbook failed
	 */
	public void saveToFile(String filename) throws IOException {
		ConfigWorkbook configWorkbook = this.getConfigWorkbook();
		if (configWorkbook.isNewConfig()) {
			configWorkbook.addLectureNames(this.getLectures());
		}
		configWorkbook.close();
		this.fillWorkbook();
		File file = new File(filename);
		ApachePOIWrapper.saveWorkbookToFile(this.getWorkbook(), file);
		this.getWorkbook().close();
	}

	/**
	 * Gets the cell ranges for the lectures and map them to the lectures. The cell
	 * range will be varying for parallel lectures.
	 * 
	 * For more details how the cell ranges for parallel lectures will be handled,
	 * see {@link LectureWorkbook#adjustOverlappingCellRanges(List)}.
	 * 
	 * @see LectureWorkbook#adjustOverlappingCellRanges(List)
	 * @param groupedLectures The map of grouped lectures
	 * @return A map of grouped lectures mapped to their cell range
	 */
	private Map<String, Map<Lecture, CellRangeAddress>> mapCellRangesForParallelLectures(
			Map<String, List<Lecture>> groupedLectures) {

		Map<String, Map<Lecture, CellRangeAddress>> groupedLecturesCellRangeMap = new TreeMap<String, Map<Lecture, CellRangeAddress>>();
		Calendar quarterStartDate = this.getQuarterStartDate();

		List<CellRangeAddress> cellRanges = new ArrayList<CellRangeAddress>();

		for (Entry<String, List<Lecture>> lectureListEntry : groupedLectures.entrySet()) {
			Map<Lecture, CellRangeAddress> lectureCellRangeMap = new HashMap<Lecture, CellRangeAddress>();
			for (Lecture lecture : lectureListEntry.getValue()) {
				CellRangeAddress cellRange = LectureWorkbook.getCellRangeFromLecture(quarterStartDate, lecture);
				if (cellRange != null) {
					lectureCellRangeMap.put(lecture, cellRange);
					cellRanges.add(cellRange);
				} else {
					this.getErrorOutput().addErrorMessage("Skipped the lecture \"" + lecture.toShortString()
							+ "\", because the lecture date is not in the visible semester time table.");
				}
			}
			groupedLecturesCellRangeMap.put(lectureListEntry.getKey(), lectureCellRangeMap);
		}

		LectureWorkbook.adjustCellRanges(cellRanges);

		return groupedLecturesCellRangeMap;
	}

	/**
	 * Adjusts overlapping cell ranges in the cell range list.
	 * 
	 * If there are two lectures with overlapping start and/or end date, each
	 * lecture gets generally the half of the total cell range of both lectures. But
	 * the cell range of each lecture cannot decrease. So if one lecture cell range
	 * is smaller than the total cell range, the other lecture cell range will
	 * decrease for the remaining cell range space.
	 * 
	 * @param cellRanges The list of cell ranges to adjust
	 */
	private static void adjustCellRanges(List<CellRangeAddress> cellRanges) {
		List<CellRangeAddress> cellRangeList = new ArrayList<>(cellRanges);

		int cellRangeListSize = cellRangeList.size();
		while (cellRangeListSize > 0) {
			int lastCellRangeInList = cellRangeListSize - 1;
			CellRangeAddress cellRange = cellRangeList.get(lastCellRangeInList);
			cellRangeList.remove(lastCellRangeInList);

			List<CellRangeAddress> overlapCellRanges = new ArrayList<CellRangeAddress>();
			overlapCellRanges.addAll(LectureWorkbook.getRecursiveOverlapCellRanges(cellRange, cellRangeList));

			if (overlapCellRanges.size() > 0) {
				overlapCellRanges.add(cellRange);
				LectureWorkbook.adjustOverlappingCellRanges(overlapCellRanges);
			}

			cellRangeListSize = cellRangeList.size();
		}
	}

	/**
	 * Returns a list of all cell ranges from the given cell range list, which
	 * overlaps the given cell range and all overlapping cell ranges.
	 * 
	 * @param cellRange     The cell range for checking overlapping cell ranges
	 * @param cellRangeList The list of possible overlapping cell ranges
	 * @return A list of all overlapping cell ranges
	 */
	private static List<CellRangeAddress> getRecursiveOverlapCellRanges(CellRangeAddress cellRange,
			List<CellRangeAddress> cellRangeList) {
		return LectureWorkbook.getRecursiveOverlapCellRanges(cellRange, cellRangeList, cellRangeList.size() - 1);
	}

	/**
	 * Returns a list of all cell ranges from the given cell range list, which
	 * overlaps the given cell range and all overlapping cell ranges.
	 * 
	 * @param cellRange          The cell range for checking overlapping cell ranges
	 * @param cellRangeList      The list of possible overlapping cell ranges
	 * @param lastCellRangeIndex The index of the last cell range, which should be
	 *                           checked
	 * @return A list of all overlapping cell ranges
	 */
	private static List<CellRangeAddress> getRecursiveOverlapCellRanges(CellRangeAddress cellRange,
			List<CellRangeAddress> cellRangeList, int lastCellRangeIndex) {

		List<CellRangeAddress> overlapCellRanges = new ArrayList<CellRangeAddress>();

		lastCellRangeIndex = lastCellRangeIndex < cellRangeList.size() ? lastCellRangeIndex : cellRangeList.size() - 1;

		for (int index = lastCellRangeIndex; index >= 0; index--) {
			CellRangeAddress otherCellRange = cellRangeList.get(index);

			if (cellRange.intersects(otherCellRange)) {
				cellRangeList.remove(index);
				overlapCellRanges.addAll(
						LectureWorkbook.getRecursiveOverlapCellRanges(otherCellRange, cellRangeList, index - 1));
				overlapCellRanges.add(otherCellRange);
			}
		}

		return overlapCellRanges;
	}

	/**
	 * Adjusts overlapping cell ranges in the cell range list.
	 * 
	 * This is a helper method for {@link LectureWorkbook#adjustCellRanges(List)}.
	 * Only use this method if you are sure, that all cell ranges are overlapping,
	 * otherwise it can cause layout errors.
	 * 
	 * @param cellRanges The list of overlapping cell ranges to adjust
	 */
	private static void adjustOverlappingCellRanges(List<CellRangeAddress> cellRanges) {

		cellRanges.sort((CellRangeAddress a, CellRangeAddress b) -> a.getLastRow() - b.getLastRow());
		cellRanges.sort((CellRangeAddress a, CellRangeAddress b) -> a.getFirstRow() - b.getFirstRow());

		while (cellRanges.size() > 0) {
			CellRangeAddress cellRange = cellRanges.get(0);

			List<CellRangeAddress> overlapCellRanges = LectureWorkbook.getOverlapCellRanges(cellRange, cellRanges);
			int totalCellRanges = overlapCellRanges.size();
			if (totalCellRanges > 1) {
				CellRangeAddress nextCellRange = overlapCellRanges.get(1);

				int minRow = cellRange.getFirstRow();
				int maxRow = Collections
						.max(overlapCellRanges,
								(CellRangeAddress a, CellRangeAddress b) -> a.getLastRow() - b.getLastRow())
						.getLastRow();

				int totalHeight = maxRow - minRow;
				int possibleRangeHeight = totalHeight / totalCellRanges;
				possibleRangeHeight = totalHeight % totalCellRanges > 1 ? possibleRangeHeight + 1 : possibleRangeHeight;

				int rangeHeight = cellRange.getLastRow() - minRow;
				int newRangeHeight = Math.min(rangeHeight, possibleRangeHeight);
				int possibleLastRow = minRow + newRangeHeight;
				int lastRow = Math.max(possibleLastRow, nextCellRange.getFirstRow() - 1);
				cellRange.setLastRow(lastRow);
				nextCellRange.setFirstRow(lastRow + 1);
			}
			cellRanges.remove(0);
		}
	}

	/**
	 * Returns a list of all cell ranges from the given cell range list, which
	 * overlaps the given cell range.
	 * 
	 * @param cellRange     The cell range for checking overlapping cell ranges
	 * @param cellRangeList The list of possible overlapping cell ranges
	 * @return A list of all overlapping cell ranges
	 */
	private static List<CellRangeAddress> getOverlapCellRanges(CellRangeAddress cellRange,
			List<CellRangeAddress> cellRangeList) {
		List<CellRangeAddress> overlapCellRanges = new ArrayList<CellRangeAddress>();

		for (CellRangeAddress otherCellRange : cellRangeList) {
			if (cellRange.intersects(otherCellRange)) {
				overlapCellRanges.add(otherCellRange);
			}
		}

		return overlapCellRanges;
	}

	/**
	 * Insert the lectures from the groupedLectures variable into the lecture area
	 * of the workbook sheet.
	 */
	private void addLecturesToWorkbook() {
		XSSFWorkbook workbook = this.getWorkbook();

		ConfigWorkbook configWorkbook = this.getConfigWorkbook();
		Map<String, LectureProperties> lecturePropertiesMap = configWorkbook.getLecturePropertiesMap();
		Map<String, XSSFFont> highlightedFonts = configWorkbook.getHighlightedFonts();
		String[] ignorePrefixes = configWorkbook.getIgnorePrefixes();

		Map<String, Map<Lecture, CellRangeAddress>> groupedLecturesCellRangeMap = this
				.mapCellRangesForParallelLectures(this.getGroupedLectures());

		XSSFFont defaultFont = workbook.createFont();
		defaultFont.setFontHeight((short) 200);
		defaultFont.setFontName("Arial");

		XSSFCellStyle defaultCellStyle = workbook.createCellStyle();
		defaultCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		defaultCellStyle.setWrapText(true);
		defaultCellStyle.setVerticalAlignment(VerticalAlignment.TOP);
		defaultCellStyle.setFillForegroundColor(ApachePOIWrapper.colorToXSSFColor(Color.WHITE));
		defaultCellStyle.setFont(defaultFont);

		for (Entry<String, Map<Lecture, CellRangeAddress>> groupedLecture : groupedLecturesCellRangeMap.entrySet()) {
			Map<Lecture, CellRangeAddress> lectureCellRangeMap = groupedLecture.getValue();
			String groupedLectureName = groupedLecture.getKey();

			Map<XSSFFont, Integer[]> lectureNameFonts = LectureWorkbook.getTextHighlights(groupedLectureName,
					highlightedFonts);

			String rawLectureName = LectureWorkbook.removePrefixFromString(groupedLectureName, ignorePrefixes);
			LectureProperties lectureProperties = LectureWorkbook.getLecturePropertiesFromMap(rawLectureName,
					lecturePropertiesMap);

			XSSFFont mainFont = null;
			XSSFCellStyle cellStyle = null;

			String shortLectureName = rawLectureName == LectureWorkbook.HOLIDAY ? "" : groupedLectureName;
			if (lectureProperties != null) {
				mainFont = workbook.createFont();

				cellStyle = workbook.createCellStyle();
				cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				cellStyle.setWrapText(true);
				cellStyle.setVerticalAlignment(VerticalAlignment.TOP);

				XSSFFont lectureFont = lectureProperties.getFont();
				ApachePOIWrapper.copyFont(mainFont, lectureFont);

				cellStyle.setFillForegroundColor(lectureProperties.getFillColor());
				cellStyle.setFont(mainFont);

				String subShortLectureName = lectureProperties.getShortLectureName();
				if (subShortLectureName != "") {
					shortLectureName = shortLectureName.replace(rawLectureName,
							lectureProperties.getShortLectureName());
				}
			} else {
				mainFont = defaultFont;
				cellStyle = defaultCellStyle;
			}

			for (Entry<Lecture, CellRangeAddress> lectureCellRangeEntry : lectureCellRangeMap.entrySet()) {
				Lecture lecture = lectureCellRangeEntry.getKey();
				lecture.setShortName(shortLectureName);
				CellRangeAddress cellRange = lectureCellRangeEntry.getValue();
				this.addLectureToWorkbook(cellRange, cellStyle, mainFont, lectureNameFonts, lecture);
			}
		}
	}

	/**
	 * Inserts a given lecture into the lecture area of the workbook sheet.
	 * 
	 * The lectureNameFonts are a map which contains the font and an integer array
	 * where to add the font in the text. The array contains at index 0 the start
	 * index (inclusive) and at index 1 the end index (exclusive).
	 * 
	 * @param cellRange        The cell range for the lecture
	 * @param cellStyle        The style for the lecture
	 * @param mainFont         The main font for the lecture
	 * @param lectureNameFonts A map for highlighting areas of the lecture text
	 * @param lecture          The lecture itself
	 * @return true if inserting the lecture was successful, false otherwise
	 */
	private boolean addLectureToWorkbook(CellRangeAddress cellRange, XSSFCellStyle cellStyle, XSSFFont mainFont,
			Map<XSSFFont, Integer[]> lectureNameFonts, Lecture lecture) {
		boolean mergedSuccessful = false;
		boolean addedSuccessful = false;
		XSSFSheet sheet = ApachePOIWrapper.getSheet(this.getWorkbook());
		if (cellRange != null) {
			try {
				sheet.addMergedRegion(cellRange);
				mergedSuccessful = true;
			} catch (IllegalStateException | IllegalArgumentException e) {
				this.errorOutput.addErrorMessage("Skipped the lecture \"" + lecture.toShortString()
						+ "\", because the cell range for the lecture cannot be merged. Maybe there is a problem with overlapping lectures.");
			}
		}

		if (mergedSuccessful) {
			XSSFCell cell = sheet.getRow(cellRange.getFirstRow()).getCell(cellRange.getFirstColumn());
			boolean modifiedCellRange = !cellRange
					.equals(LectureWorkbook.getCellRangeFromLecture(this.getQuarterStartDate(), lecture));
			cell.setCellValue(
					LectureWorkbook.lectureToRichText(mainFont, lectureNameFonts, lecture, modifiedCellRange));
			cell.setCellStyle(cellStyle);
			addedSuccessful = true;
		}

		return addedSuccessful;
	}

	/**
	 * Adds all holidays between the start and end quarter date to the grouped
	 * lectures with {@value LectureWorkbook#HOLIDAY} as key for all holidays. The
	 * lectures for the holidays do not have any resources or lecturers. Each
	 * lecture for a holiday will start at 00:00:00.001 and ends at 00:00:00.000 on
	 * the next day.
	 */
	private void addHolidays() {
		Calendar quarterStartDate = this.getQuarterStartDate();
		Calendar quarterEndDate = this.getQuarterEndDate();
		Map<String, List<Lecture>> groupedLectures = this.getGroupedLectures();
		if (quarterStartDate != null && quarterEndDate != null && groupedLectures != null) {
			Calendar quarterIncludedEndDate = (Calendar) quarterEndDate.clone();
			quarterIncludedEndDate.add(Calendar.DAY_OF_MONTH, -1);
			Map<Calendar, String> holidays = Holidays.getHolidays(quarterStartDate, quarterIncludedEndDate,
					this.getConfigWorkbook().getHolidayLocale());
			List<Lecture> holidaysLecture = new ArrayList<Lecture>();
			for (Entry<Calendar, String> holiday : holidays.entrySet()) {
				Calendar startDate = holiday.getKey();
				startDate.setTimeZone(quarterStartDate.getTimeZone());
				Calendar endDate = (Calendar) startDate.clone();
				startDate.add(Calendar.MILLISECOND, 1);
				endDate.add(Calendar.DAY_OF_MONTH, 1);
				Lecture lecture = new Lecture(holiday.getValue(), startDate, endDate, "", "");
				holidaysLecture.add(lecture);
			}
			groupedLectures.put(LectureWorkbook.HOLIDAY, holidaysLecture);
		}
	}

	/**
	 * Converts a lecture to a rich text by adding the main font and highlighting
	 * fonts.
	 * 
	 * @param mainFont          The main font for the lecture
	 * @param lectureNameFonts  A map for highlighting areas of the lecture text
	 * @param lecture           The lecture itself
	 * @param modifiedCellRange True if the cell range for the lecture has changed
	 *                          (caused by parallel lectures), otherwise false
	 * @return The lecture converted to a rich text
	 */
	private static XSSFRichTextString lectureToRichText(XSSFFont mainFont, Map<XSSFFont, Integer[]> lectureNameFonts,
			Lecture lecture, boolean modifiedCellRange) {
		String shortLectureName = lecture.getShortName();
		String text = shortLectureName != null && shortLectureName != "" ? shortLectureName : lecture.getName();
		if (modifiedCellRange || !LectureWorkbook.hasLectureNormalTimeInterval(lecture)) {
			String startTime = LectureWorkbook.getTime(lecture.getStartDate());
			String endTime = LectureWorkbook.getTime(lecture.getEndDate());
			text += LectureWorkbook.LINE_BREAK + startTime + "-" + endTime;
		}
		String[] lecturers = lecture.getLecturers();
		if (lecturers != null) {
			text += LectureWorkbook.LINE_BREAK + LectureWorkbook.arrayToString(lecturers);
		}
		String[] resources = lecture.getResources();
		if (resources != null) {
			text += LectureWorkbook.LINE_BREAK + LectureWorkbook.arrayToString(lecture.getResources());
		}
		// TODO move following lines to ApachePOIWrapper?
		XSSFRichTextString richText = new XSSFRichTextString(text);
		richText.applyFont(mainFont);
		short mainFontHeight = mainFont.getFontHeight();
		String mainFontName = mainFont.getFontName();
		for (Entry<XSSFFont, Integer[]> lectureNameFont : lectureNameFonts.entrySet()) {
			Integer[] indexes = lectureNameFont.getValue();
			XSSFFont nameFont = lectureNameFont.getKey();
			nameFont.setFontHeight(mainFontHeight);
			nameFont.setFontName(mainFontName);
			richText.applyFont(indexes[0], indexes[1], nameFont);
		}
		return richText;
	}

	/**
	 * Returns the color array of the given name from the color map. If the name
	 * does not match lecture properties key, then null is returned.
	 *
	 * The name can contain a '*' as a wildcard.
	 * 
	 * @param name                 The name for matching a key
	 * @param lecturePropertiesMap A map of lecture properties
	 * @return The color array of the name from the color map, or null if name not
	 *         matches a color map key
	 */
	public static LectureProperties getLecturePropertiesFromMap(String name,
			Map<String, LectureProperties> lecturePropertiesMap) {
		LectureProperties lectureProperties = null;
		for (String key : lecturePropertiesMap.keySet()) {
			String matchKey = "\\Q" + key.replace("*", "\\E.*\\Q") + "\\E";
			if (name.matches(matchKey)) {
				lectureProperties = lecturePropertiesMap.get(key);
				break;
			}
		}
		return lectureProperties;
	}

	/**
	 * Returns a map of fonts and their index pairs in the given text. All matches
	 * of the font map key in the text will be added to the font index pair map. The
	 * index pair consists always of the start and end index.
	 * 
	 * @param text    The text, which will be scanned
	 * @param fontMap A map of pattern and their color
	 * @return A map of fonts and their index pairs (the integer array always
	 *         consists of the start and end index)
	 */
	public static Map<XSSFFont, Integer[]> getTextHighlights(String text, Map<String, XSSFFont> fontMap) {
		Map<XSSFFont, Integer[]> fontIndexMap = new HashMap<XSSFFont, Integer[]>();
		for (String key : fontMap.keySet()) {
			Pattern pattern = Pattern.compile("\\Q" + key + "\\E");
			Matcher matcher = pattern.matcher(text);
			XSSFFont font = fontMap.get(key);
			while (matcher.find()) {
				fontIndexMap.put(font, new Integer[] { matcher.start(), matcher.end() });
			}
		}
		return fontIndexMap;
	}

	/**
	 * Returns the input string without the first matching prefix of the prefixes
	 * array. If no prefix matches the string, the input string is returned without
	 * any changes.
	 * 
	 * @param string   The string for removing a prefix
	 * @param prefixes An array of prefixes
	 * @return The input string without the first matching prefix
	 */
	public static String removePrefixFromString(String string, String[] prefixes) {
		String rawString = string;
		for (String ignorePrefix : prefixes) {
			if (string.startsWith(ignorePrefix + " ")) {
				rawString = string.substring(ignorePrefix.length() + 1);
				break;
			}
		}
		return rawString;
		// TODO move method to helper class
	}

	/**
	 * Returns the array as an comma (",") separated string.
	 * 
	 * @param array The array to convert
	 * @return The comma separated array
	 */
	private static String arrayToString(String[] array) {
		String string = "";
		if (array != null) {
			for (String element : array) {
				string += "," + element;
			}
			string = string == "" ? "" : string.substring(1);
		}
		return string;
	}

	/**
	 * Returns the cell range address for a given lecture.
	 * 
	 * @param quarterStartDate The start date of the quarter
	 * @param lecture          The lecture
	 * @return The cell range address of the lecture
	 */
	public static CellRangeAddress getCellRangeFromLecture(Calendar quarterStartDate, Lecture lecture) {
		CellAddress startCellAddress = LectureWorkbook.getCellAddressFromDate(quarterStartDate, lecture.getStartDate(),
				false);
		CellAddress endCellAddress = LectureWorkbook.getCellAddressFromDate(quarterStartDate, lecture.getEndDate(),
				true);
		CellRangeAddress cellRange;

		if (startCellAddress == null || endCellAddress == null) {
			cellRange = null;
		} else {
			cellRange = new CellRangeAddress(startCellAddress.getRow(), endCellAddress.getRow(),
					startCellAddress.getColumn(), endCellAddress.getColumn());
		}

		return cellRange;
	}

	/**
	 * Returns the cell address for a given date.
	 * 
	 * If the date is the start time of the lecture, set isLectureEnd to false.
	 * 
	 * If the date is the end time of the lecture, set isLectureEnd to true.
	 * 
	 * @param quarterStartDate The start date of the quarter
	 * @param lectureDate      The start or end time of a lecture
	 * @param isLectureEnd     A boolean whether the lectureDate is the end time of
	 *                         the lecture or not
	 * @return
	 */
	public static CellAddress getCellAddressFromDate(Calendar quarterStartDate, Calendar lectureDate,
			boolean isLectureEnd) {
		CellAddress cellAddress;
		Calendar date = (Calendar) lectureDate.clone();
		if (isLectureEnd) {
			date.add(Calendar.MINUTE, -1);
		}
		int lectureDayOfWeek = date.get(Calendar.DAY_OF_WEEK);

		int daysBetween = (int) ChronoUnit.DAYS.between(quarterStartDate.toInstant(), date.toInstant());

		if (lectureDayOfWeek == Calendar.SUNDAY || lectureDayOfWeek == Calendar.SATURDAY || daysBetween > 81) {
			cellAddress = null;
		} else {
			int columnNum = daysBetween % 28;
			columnNum -= columnNum / 7 * 2;
			columnNum += columnNum / 10 + 1;

			int rowNum = daysBetween / 28 * 49 + 4;

			int minutes = date.get(Calendar.MINUTE);
			int hours = date.get(Calendar.HOUR_OF_DAY);
			int timePosition = ((hours - 8) * 4) + (minutes / 15);

			if (timePosition < 0) {
				rowNum--;
			} else if (timePosition >= 44) {
				rowNum += 44;
			} else {
				rowNum += timePosition;
			}

			cellAddress = new CellAddress(rowNum, columnNum);
		}

		return cellAddress;
	}

	/**
	 * Returns whether the start and end time of the given lecture is a "normal"
	 * start and end time or not.
	 * 
	 * The normal start times are defined in {@link LectureWorkbook#START_TIMES}
	 * 
	 * The normal end times are defined in {@link LectureWorkbook#END_TIMES}
	 * 
	 * @param lecture The lecture for checking start and end times
	 * @return true if the lecture has normal start and end time, false otherwise
	 */
	public static boolean hasLectureNormalTimeInterval(Lecture lecture) {
		String startTime = LectureWorkbook.getTime(lecture.getStartDate());
		String endTime = LectureWorkbook.getTime(lecture.getEndDate());
		return LectureWorkbook.START_TIMES.contains(startTime) && LectureWorkbook.END_TIMES.contains(endTime);
	}

	/**
	 * Returns the time of a calendar in the hh:mm format of a 24 hour clock.
	 * 
	 * @param calendar The calendar for time extraction
	 * @return The time of the calendar
	 */
	public static String getTime(Calendar calendar) {
		int hours = calendar.get(Calendar.HOUR_OF_DAY);
		int minutes = calendar.get(Calendar.MINUTE);
		String hour = hours < 10 ? "0" + hours : Integer.toString(hours);
		String minute = minutes < 10 ? "0" + minutes : Integer.toString(minutes);
		return hour + ":" + minute;
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
	public static Calendar weekOfYearToDate(int weekOfYear, int dayOfWeek, int year, TimeZone timeZone) {
		Calendar calendar = new GregorianCalendar();
		calendar.set(Calendar.DAY_OF_WEEK, dayOfWeek);
		calendar.set(Calendar.WEEK_OF_YEAR, weekOfYear);
		calendar.set(Calendar.YEAR, year);
		calendar.setTimeZone(timeZone);
		calendar.set(Calendar.HOUR_OF_DAY, 0);
		calendar.set(Calendar.MINUTE, 0);
		calendar.set(Calendar.SECOND, 0);
		calendar.set(Calendar.MILLISECOND, 0);
		return calendar;
	}

	/**
	 * Returns the input stream of a template file of a given filename. The template
	 * file is stored inside the root source folder of this class.
	 * 
	 * @param filename The filename for a template file
	 * @return The input stream of the filename
	 */
	public static InputStream getTemplateInputStream(String filename) {
		return LectureWorkbook.class.getClassLoader().getResourceAsStream(filename);
	}
}
