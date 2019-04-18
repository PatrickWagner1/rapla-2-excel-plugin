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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LectureWorkbook {

	public final static List<String> START_TIMES = new ArrayList<String>(Arrays.asList("8:00", "8:45", "9:45", "10:30",
			"11:30", "12:15", "14:00", "14:45", "15:45", "16:30", "17:30", "18:15"));
	public final static List<String> END_TIMES = new ArrayList<String>(Arrays.asList("8:45", "9:30", "10:30", "11:15",
			"12:15", "13:00", "14:45", "15:30", "16:30", "17:15", "18:15", "19:00"));

	// TODO Add more useful colors to the lectureColors array and maybe delete
	// absurd colors
	public final static short[] LECTURE_COLORS = new short[] { IndexedColors.BLUE.getIndex(),
			IndexedColors.AQUA.getIndex(), IndexedColors.GREEN.getIndex() };

	private final static String TEMPLATE_FILENAME = "template.xlsx";

	private Workbook workbookTemplate;
	private Workbook workbook;
	private Calendar quarterStartDate;
	private List<Lecture> lectures;

	public LectureWorkbook(Calendar quarterStartDate, List<Lecture> lectures) throws IOException {
		this.workbookTemplate = LectureWorkbook.getTemplateWorkbook();
		this.setLectures(lectures);
		this.setQuarterStartDate(quarterStartDate);
	}

	public LectureWorkbook(String filename, Calendar quarterStartDate, List<Lecture> lectures) throws IOException {
		File file = new File(filename);
		this.workbookTemplate = LectureWorkbook.loadWorkbookFromFile(file);
		this.setLectures(lectures);
		this.setQuarterStartDate(quarterStartDate);
	}
	
	public Workbook getWorkbook() {
		return this.workbook;
	}
	
	private void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}
	
	public void resetWorkbook() {
		this.setWorkbook(this.workbookTemplate);
		this.addLectures(this.getLectures());
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
	
	private void setLectures(List<Lecture> lectures) {
		this.lectures = lectures;
	}

	public void saveToFile(String filename) throws IOException {
		File file = new File(filename);
		LectureWorkbook.saveWorkbookToFile(this.workbook, file);
	}

	public void addLectures(List<Lecture> lectures) {
		for (Lecture lecture : lectures) {
			addLecture(lecture);
		}
	}

	public void addLecture(Lecture lecture) {
		int[] cellRange = getCellRangeFromLecture(this.quarterStartDate, lecture);
		Row row = this.getWorkbook().getSheetAt(0).getRow(cellRange[0]);
		Cell cell = row.getCell(cellRange[2]);
		String boltText = lecture.getName();
		String normalText = "\n";
		if (!LectureWorkbook.hasLectureNormalTimeInterval(lecture)) {
			String startTime = LectureWorkbook.getTime(lecture.getStartDate());
			String endTime = LectureWorkbook.getTime(lecture.getEndDate());
			normalText += startTime + "-" + endTime + "\n";
		}
		normalText += arrayToString(lecture.getResources()) + "\n" + arrayToString(lecture.getLecturers());

		CellStyle cellStyle = this.getWorkbook().createCellStyle();
		cellStyle.setFillForegroundColor(LectureWorkbook.LECTURE_COLORS[0]);
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);

		//TODO better formating
		
		Font font = new XSSFFont();
		// font.setBold(true);
		font.setFontHeight((short) 200);
		font.setFontName("Arial");
		RichTextString richText = new XSSFRichTextString(boltText + normalText);
		richText.applyFont(0, boltText.length(), font);

		try {
			this.workbook.getSheetAt(0).addMergedRegion(new CellRangeAddress(cellRange[0], cellRange[1], cellRange[2], cellRange[3]));
		} catch (IllegalStateException e) {
			
			//TODO parallel lecture handling
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
		if (isLectureEnd) {
			minutes--;
		}
		int timePosition = ((lectureDate.get(Calendar.HOUR_OF_DAY) - 8) * 4) + (minutes / 15);

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

	public static Workbook loadWorkbookFromFile(File file) throws IOException {
		FileInputStream excelFile = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(excelFile);
		return workbook;
	}

	public static Sheet getSheetFromWorkbook(Workbook workbook) {
		Sheet sheet = workbook.getSheetAt(0);
		return sheet;
	}

	public static void saveWorkbookToFile(Workbook workbook, File file) throws IOException {
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
		Workbook wbTemplate = LectureWorkbook.loadWorkbookFromFile(template);
		LectureWorkbook.saveWorkbookToFile(wbTemplate, outFile);
	}

	public static Workbook getTemplateWorkbook() throws IOException {
		File file = new File(
				LectureWorkbook.class.getClassLoader().getResource(LectureWorkbook.TEMPLATE_FILENAME).getFile());
		return LectureWorkbook.loadWorkbookFromFile(file);
	}

}
