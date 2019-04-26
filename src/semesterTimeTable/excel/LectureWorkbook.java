package semesterTimeTable.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LectureWorkbook {

	public final static List<String> START_TIMES = new ArrayList<String>(Arrays.asList("08:00", "08:45", "09:45", "10:30",
			"11:30", "12:15", "14:00", "14:45", "15:45", "16:30", "17:30", "18:15"));
	public final static List<String> END_TIMES = new ArrayList<String>(Arrays.asList("08:45", "09:30", "10:30", "11:15",
			"12:15", "13:00", "14:45", "15:30", "16:30", "17:15", "18:15", "19:00"));

	// TODO Add more useful colors to the lectureColors array and maybe delete
	// absurd colors
	public final static short[] LECTURE_COLORS = new short[] { IndexedColors.WHITE1.getIndex(),
			IndexedColors.BLACK.getIndex(), IndexedColors.BLUE.getIndex(), IndexedColors.AQUA.getIndex(),
			IndexedColors.GREEN.getIndex(), IndexedColors.BLUE_GREY.getIndex(), IndexedColors.BRIGHT_GREEN.getIndex(),
			IndexedColors.DARK_YELLOW.getIndex(), IndexedColors.LAVENDER.getIndex(),
			IndexedColors.LIGHT_ORANGE.getIndex(), IndexedColors.OLIVE_GREEN.getIndex(),
			IndexedColors.LIGHT_GREEN.getIndex() };

	private final static String TEMPLATE_FILENAME = "template.xlsx";
	private XSSFWorkbook workbook;
	private Calendar quarterStartDate;
	private List<Lecture> lectures;

	private List<List<Integer>> borderColumnsLastRow;
	private List<List<Integer>> borderColumnsExamWeek;
	private List<List<Integer>> borderColumns;
	private List<List<Integer>> borderRowsExamWeek;
	private List<List<Integer>> borderRows;
	private XSSFCellStyle[][] borderStyle;
	private XSSFCellStyle[][] borderStyleExamWeek;

	public LectureWorkbook(Calendar quarterStartDate, List<Lecture> lectures) throws IOException {
		this.setWorkbook(LectureWorkbook.getTemplateWorkbook());
		this.createBorderStyles();
		this.lectures = lectures;
		this.setQuarterStartDate(quarterStartDate);
	}

	public LectureWorkbook(String filename, Calendar quarterStartDate, List<Lecture> lectures) throws IOException {
		File file = new File(filename);
		if (file.exists()) {
			this.setWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.setWorkbook(LectureWorkbook.getTemplateWorkbook());
		}
		this.createBorderStyles();
		this.lectures = lectures;
		this.setQuarterStartDate(quarterStartDate);
	}

	public XSSFWorkbook getWorkbook() {
		return this.workbook;
	}

	private void setWorkbook(XSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	public void resetWorkbook() {
		XSSFSheet sheet = this.workbook.getSheetAt(0);
		List<CellRangeAddress> ranges = sheet.getMergedRegions();
		for (int index = sheet.getNumMergedRegions() - 1; index > -1; index--) {
			CellRangeAddress range = ranges.get(index);
			if (this.isMergedLectureCells(range.getFirstRow(), range.getFirstColumn())) {
				sheet.removeMergedRegion(index);
			}
		}

		this.resetLectures();

		this.addLecturesToWorkbook(this.getLectures());
		this.getWorkbook().getSheetAt(0).getRow(2).getCell(1).setCellValue(this.quarterStartDate);
		XSSFFormulaEvaluator.evaluateAllFormulaCells(this.getWorkbook());
	}

	private boolean isMergedLectureCells(int row, int column) {
		if ((((row > 2 && row < 49) || (row > 51 && row < 98) || (row > 100 && row < 139))
				&& (column > 0 && column != 11 && column < 22))
				|| ((row > 138 && row < 147) && (column > 0 && column != 11 && column < 16))) {
			return true;
		} else {
			return false;
		}
	}

	private void createBorderStyles() {
		this.borderColumnsLastRow = new ArrayList<List<Integer>>(3);
		this.borderColumnsExamWeek = new ArrayList<List<Integer>>(3);
		this.borderColumns = new ArrayList<List<Integer>>(3);

		borderColumnsLastRow.add(0, Arrays.asList(1, 6, 12));
		borderColumnsExamWeek.add(0, Arrays.asList(17));
		borderColumns.add(0, new ArrayList<Integer>());
		borderColumns.get(0).addAll(borderColumnsLastRow.get(0));
		borderColumns.get(0).addAll(borderColumnsExamWeek.get(0));

		borderColumnsLastRow.add(1, Arrays.asList(2, 3, 4, 7, 8, 9, 13, 14, 15));
		borderColumnsExamWeek.add(1, Arrays.asList(18, 19, 20));
		borderColumns.add(1, new ArrayList<Integer>());
		borderColumns.get(1).addAll(borderColumnsLastRow.get(1));
		borderColumns.get(1).addAll(borderColumnsExamWeek.get(1));
		
		borderColumnsLastRow.add(2, Arrays.asList(5, 10));
		borderColumnsExamWeek.add(2, Arrays.asList(16, 21));
		borderColumns.add(2, new ArrayList<Integer>());
		borderColumns.get(2).addAll(borderColumnsLastRow.get(2));
		borderColumns.get(2).addAll(borderColumnsExamWeek.get(2));

		this.borderRowsExamWeek = new ArrayList<List<Integer>>(3);
		borderRowsExamWeek.add(0,
				Arrays.asList(5, 6, 8, 9, 12, 13, 15, 16, 19, 20, 22, 23, 25, 26, 27, 29, 30, 32, 33, 36, 37, 39, 40));
		borderRowsExamWeek.add(1, Arrays.asList(4, 7, 10, 11, 14, 17, 18, 21, 31, 34, 35, 38));
		borderRowsExamWeek.add(2, Arrays.asList(3, 24, 28));

		this.borderRows = new ArrayList<List<Integer>>(3);
		borderRows.add(0, new ArrayList<Integer>());
		borderRows.add(1, new ArrayList<Integer>());
		borderRows.add(2, new ArrayList<Integer>());

		borderRows.get(2).addAll(borderRowsExamWeek.get(2));
		borderRows.get(1).addAll(borderRowsExamWeek.get(1));
		borderRows.get(1).addAll(Arrays.asList(41, 42, 45, 48));
		borderRows.get(0).addAll(borderRowsExamWeek.get(0));
		borderRows.get(0).addAll(Arrays.asList(43, 44, 46, 47));

		// leftNoTop middleNoTop rightNoTop
		// leftGrey middleGrey rightGrey
		// leftBlack middleBlack rightBlack
		this.borderStyle = new XSSFCellStyle[3][3];

		borderStyle[0][1] = this.getWorkbook().createCellStyle();
		borderStyle[0][1].setBorderLeft(BorderStyle.THIN);
		borderStyle[0][1].setBorderRight(BorderStyle.THIN);
		borderStyle[0][1].setFillForegroundColor(IndexedColors.WHITE1.getIndex());
		borderStyle[0][1].setFillPattern(FillPatternType.SOLID_FOREGROUND);

		borderStyle[0][0] = this.getWorkbook().createCellStyle();
		borderStyle[0][0].cloneStyleFrom(borderStyle[0][1]);
		borderStyle[0][0].setBorderLeft(BorderStyle.THICK);

		borderStyle[0][2] = this.getWorkbook().createCellStyle();
		borderStyle[0][2].cloneStyleFrom(borderStyle[0][1]);
		borderStyle[0][2].setBorderRight(BorderStyle.THICK);

		borderStyle[2][1] = this.getWorkbook().createCellStyle();
		borderStyle[2][1].cloneStyleFrom(borderStyle[0][1]);
		borderStyle[2][1].setBorderTop(BorderStyle.THIN);

		borderStyle[2][0] = this.getWorkbook().createCellStyle();
		borderStyle[2][0].cloneStyleFrom(borderStyle[2][1]);
		borderStyle[2][0].setBorderLeft(BorderStyle.THICK);

		borderStyle[2][2] = this.getWorkbook().createCellStyle();
		borderStyle[2][2].cloneStyleFrom(borderStyle[2][1]);
		borderStyle[2][2].setBorderRight(BorderStyle.THICK);
		
		borderStyle[1][1] = this.getWorkbook().createCellStyle();
		borderStyle[1][1].cloneStyleFrom(borderStyle[2][1]);
		borderStyle[1][1].setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

		borderStyle[1][0] = this.getWorkbook().createCellStyle();
		borderStyle[1][0].cloneStyleFrom(borderStyle[1][1]);
		borderStyle[1][0].setBorderLeft(BorderStyle.THICK);

		borderStyle[1][2] = this.getWorkbook().createCellStyle();
		borderStyle[1][2].cloneStyleFrom(borderStyle[1][1]);
		borderStyle[1][2].setBorderRight(BorderStyle.THICK);

		this.borderStyleExamWeek = new XSSFCellStyle[3][3];
		for (int kindOfRow = 0; kindOfRow < 3; kindOfRow++) {
			for (int kindOfColumn = 0; kindOfColumn < 3; kindOfColumn++) {
				this.borderStyleExamWeek[kindOfRow][kindOfColumn] = this.getWorkbook().createCellStyle();
				this.borderStyleExamWeek[kindOfRow][kindOfColumn]
						.cloneStyleFrom(this.borderStyle[kindOfRow][kindOfColumn]);
				this.borderStyleExamWeek[kindOfRow][kindOfColumn].setFillForegroundColor(IndexedColors.PLUM.getIndex());
				this.borderStyleExamWeek[kindOfRow][kindOfColumn].setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}
		}
		for (int kindOfColumn = 0; kindOfColumn < 3; kindOfColumn++) {
			this.borderStyleExamWeek[1][kindOfColumn].setTopBorderColor(IndexedColors.BLACK1.getIndex());
		}
	}

	private void resetLectures() {

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
							this.resetCell(rowNum, columnNum, this.borderStyle[kindOfRowIndex][kindOfColumnIndex]);
						}
						kindOfColumnIndex++;
					}
				}
				kindOfRowIndex++;
			}
		}

		// Last block without exam week
		int blockStartRow = 2 * 49;
		int kindOfRowIndex = 0;
		for (List<Integer> kindOfRow : this.borderRows) {
			for (int rawRowNum : kindOfRow) {
				int rowNum = rawRowNum + blockStartRow;
				int kindOfColumnIndex = 0;
				for (List<Integer> kindOfColumn : this.borderColumnsLastRow) {
					for (int columnNum : kindOfColumn) {
						this.resetCell(rowNum, columnNum, this.borderStyle[kindOfRowIndex][kindOfColumnIndex]);
					}
					kindOfColumnIndex++;
				}
			}
			kindOfRowIndex++;
		}

		// Exam week
		kindOfRowIndex = 0;
		for (List<Integer> kindOfRow : this.borderRowsExamWeek) {
			for (int rawRowNum : kindOfRow) {
				int rowNum = rawRowNum + blockStartRow;
				int kindOfColumnIndex = 0;
				for (List<Integer> kindOfColumn : this.borderColumnsExamWeek) {
					for (int columnNum : kindOfColumn) {
						this.resetCell(rowNum, columnNum, this.borderStyleExamWeek[kindOfRowIndex][kindOfColumnIndex]);
					}
					kindOfColumnIndex++;
				}
			}
			kindOfRowIndex++;
		}
	}

	private void resetCell(int rowNum, int columnNum, XSSFCellStyle cellStyle) {
		XSSFCell cell = this.getWorkbook().getSheetAt(0).getRow(rowNum).getCell(columnNum);
		cell.setBlank();
		cell.setCellStyle(cellStyle);
	}

	public Calendar getQuarterStartDate() {
		return this.quarterStartDate;
	}

	public void setQuarterStartDate(Calendar quarterStartDate) {
		this.quarterStartDate = quarterStartDate;
		this.resetWorkbook();
	}

	public List<Lecture> getLectures() {
		return this.lectures;
	}

	public void setLectures(List<Lecture> lectures) {
		this.lectures = lectures;
		this.resetWorkbook();
	}

	public void addLectures(List<Lecture> lectures) {
		this.lectures.addAll(lectures);
		this.resetWorkbook();
	}

	public void saveToFile(String filename) throws IOException {
		File file = new File(filename);
		LectureWorkbook.saveWorkbookToFile(this.getWorkbook(), file);
		this.getWorkbook().close();
	}

	private void addLecturesToWorkbook(List<Lecture> lectures) {
		for (Lecture lecture : lectures) {
			addLectureToWorkbook(lecture);
		}
	}

	private void addLectureToWorkbook(Lecture lecture) {
		int[] cellRange = getCellRangeFromLecture(this.quarterStartDate, lecture);
		XSSFRow row = this.getWorkbook().getSheetAt(0).getRow(cellRange[0]);
		XSSFCell cell = row.getCell(cellRange[2]);
		String boltText = lecture.getName();
		String normalText = "\n";
		if (!LectureWorkbook.hasLectureNormalTimeInterval(lecture)) {
			String startTime = LectureWorkbook.getTime(lecture.getStartDate());
			String endTime = LectureWorkbook.getTime(lecture.getEndDate());
			normalText += startTime + "-" + endTime + "\n";
		}
		normalText += arrayToString(lecture.getResources()) + "\n" + arrayToString(lecture.getLecturers());

		XSSFCellStyle cellStyle = this.getWorkbook().createCellStyle();
		// TODO Wrong colors in MS Excel and no colors in LibreOffice Calc
		if (lecture.getGroupId() < LectureWorkbook.LECTURE_COLORS.length) {
			cellStyle.setFillBackgroundColor(LectureWorkbook.LECTURE_COLORS[lecture.getGroupId()]);
			cellStyle.setFillForegroundColor(LectureWorkbook.LECTURE_COLORS[lecture.getGroupId()]);
		} else {
			cellStyle.setFillBackgroundColor(LectureWorkbook.LECTURE_COLORS[0]);
			cellStyle.setFillForegroundColor(LectureWorkbook.LECTURE_COLORS[0]);
		}
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyle.setWrapText(true);
		cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellStyle(cellStyle);

		// TODO better formating

		XSSFFont font = new XSSFFont();
		// font.setBold(true);
		font.setFontHeight((short) 200);
		font.setFontName("Arial");
		XSSFRichTextString richText = new XSSFRichTextString(boltText + normalText);
		richText.applyFont(0, boltText.length(), font);

		try {
			this.workbook.getSheetAt(0)
					.addMergedRegion(new CellRangeAddress(cellRange[0], cellRange[1], cellRange[2], cellRange[3]));
		} catch (IllegalStateException e) {

			// TODO parallel lecture handling
			System.out.println("skipped lecture" + lecture.getName() + " at " + lecture.getStartDate() + " - "
					+ lecture.getEndDate());
			return;
		}

		cell.setCellValue(richText);
	}

	private static String arrayToString(String[] array) {
		String string = "";
		for (String element : array) {
			string += "," + element;
		}
		string = string == "" ? "" : string.substring(1);
		return string;
	}

	public static int[] getCellRangeFromLecture(Calendar quarterStartDate, Lecture lecture) {
		int[] startPosition = LectureWorkbook.getCellPositionFromDate(quarterStartDate, lecture.getStartDate(), false);
		int[] endPosition = LectureWorkbook.getCellPositionFromDate(quarterStartDate, lecture.getEndDate(), true);

		if (startPosition == null && endPosition == null) {
			return null;
		}

		return new int[] { startPosition[0], endPosition[0], startPosition[1], endPosition[1] };
	}

	public static int[] getCellPositionFromDate(Calendar quarterStartDate, Calendar lectureDate, boolean isLectureEnd) {
		int lectureDateOfWeek = lectureDate.get(Calendar.DAY_OF_WEEK);

		int daysBetween = (int) ChronoUnit.DAYS.between(quarterStartDate.toInstant(), lectureDate.toInstant());

		if (lectureDateOfWeek == Calendar.SUNDAY || lectureDateOfWeek == Calendar.SATURDAY || daysBetween > 81) {
			return null;
		}

		int xPosition = daysBetween % 28;
		xPosition -= xPosition / 7 * 2;
		xPosition += xPosition / 10 + 1;

		int yPosition = daysBetween / 28 * 49 + 4;

		int minutes = lectureDate.get(Calendar.MINUTE);
		int hours = lectureDate.get(Calendar.HOUR_OF_DAY);
		if (isLectureEnd) {
			if (minutes == 0) {
				minutes = 59;
				hours--;
				
			} else {
				minutes --;
			}
		}
		int timePosition = ((hours - 8) * 4) + (minutes / 15);

		if (timePosition < 0) {
			yPosition--;
		} else if (timePosition >= 44) {
			yPosition += 44;
		} else {
			yPosition += timePosition;
		}

		return new int[] { yPosition, xPosition };
	}

	public static boolean hasLectureNormalTimeInterval(Lecture lecture) {
		String startTime = LectureWorkbook.getTime(lecture.getStartDate());
		String endTime = LectureWorkbook.getTime(lecture.getEndDate());
		return LectureWorkbook.START_TIMES.contains(startTime) && LectureWorkbook.END_TIMES.contains(endTime);
	}

	public static String getTime(Calendar calendar) {
		int hours = calendar.get(Calendar.HOUR_OF_DAY);
		int minutes = calendar.get(Calendar.MINUTE);
		String hour = hours < 10 ? "0" + hours : Integer.toString(hours);
		String minute = minutes < 10 ? "0" + minutes : Integer.toString(minutes);
		return hour + ":" + minute;
	}

	public static XSSFWorkbook loadWorkbookFromFile(File file) throws IOException {
		FileInputStream excelFile = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
		return workbook;
	}

	public static XSSFSheet getSheetFromWorkbook(XSSFWorkbook workbook) {
		XSSFSheet sheet = workbook.getSheetAt(0);
		return sheet;
	}

	public static void saveWorkbookToFile(XSSFWorkbook workbook, File file) throws IOException {
		ByteArrayOutputStream stream = new ByteArrayOutputStream();
		workbook.write(stream);
		byte[] content = stream.toByteArray();
		LectureWorkbook.writeFile(file, content);
	}

	public static void writeFile(File file, byte[] content) throws IOException {
		FileOutputStream out = new FileOutputStream(file);
		out.write(content);
		out.flush();
		out.close();
	}

	public static void saveNewFile(String fileName) throws IOException {
		File outFile = new File(fileName);
		File template = new File(LectureWorkbook.class.getClassLoader().getResource("template.xlsx").getFile());
		XSSFWorkbook wbTemplate = LectureWorkbook.loadWorkbookFromFile(template);
		LectureWorkbook.saveWorkbookToFile(wbTemplate, outFile);
	}

	public static XSSFWorkbook getTemplateWorkbook() throws IOException {
		File file = new File(
				LectureWorkbook.class.getClassLoader().getResource(LectureWorkbook.TEMPLATE_FILENAME).getFile());
		return LectureWorkbook.loadWorkbookFromFile(file);
	}

}
