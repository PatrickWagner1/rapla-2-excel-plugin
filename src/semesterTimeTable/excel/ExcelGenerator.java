package semesterTimeTable.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.temporal.ChronoField;
import java.time.temporal.ChronoUnit;
import java.util.Calendar;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeMap;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelGenerator {

	private List<Lecture> lectures;

	private IndexedColors[] lectureColors;

	public ExcelGenerator(List<Lecture> lectures) {
		// TODO Add more useful colors to the lectureColors array and maybe delete
		// absurd colors
		this.lectureColors = new IndexedColors[] { IndexedColors.BLUE, IndexedColors.AQUA, IndexedColors.GREEN };
		this.lectures = generateLecturesGroupId(lectures);
	}

	private List<Lecture> generateLecturesGroupId(List<Lecture> lectures) {
		Map<String, List<Lecture>> groupedLectures = new TreeMap<String, List<Lecture>>(lectures.stream().collect(Collectors.groupingBy(Lecture::getName)));
		List<Lecture> currentLectureList;
		int currentGroupId = 2;
		for (Entry<String, List<Lecture>> groupedLecture : groupedLectures.entrySet()) {
			currentLectureList = groupedLecture.getValue();
			if (groupedLecture.getKey().startsWith(Lecture.REPEAT_EXAM_START_STRING)) {
				for (Lecture currentLecture : currentLectureList) {
					currentLecture.setGroupId(Lecture.REPEAT_EXAM_ID);
				}
			} else {
				for (Lecture currentLecture : currentLectureList) {
					currentLecture.setGroupId(currentGroupId);
				}
				currentGroupId++;
			}
		}

		// for loop is only for testing reasons.
		for (Lecture lecture : lectures) {
			System.out.println(lecture.getName() + " -- " + lecture.getGroupId());
			if(lecture.getGroupId() <= this.lectureColors.length) {
				System.out.println(this.lectureColors[lecture.getGroupId()-1]);
			}
		}
		return lectures;
	}
	
	public static int[] getCellRangeFromLecture(Calendar quarterStartDate, Lecture lecture) {
		int[] startPosition = ExcelGenerator.getCellPositionFromDate(quarterStartDate, lecture.getStartDate(), false);
		int[] endPosition = ExcelGenerator.getCellPositionFromDate(quarterStartDate, lecture.getEndDate(), true);
		
		if (startPosition == null && endPosition == null) {
			return null;
		}
		
		int[] dimension = ExcelGenerator.getCellDimensionFromCellPositions(startPosition[0], startPosition[1], endPosition[0], endPosition[1]);
		
		return new int[] {startPosition[0], startPosition[1], dimension[0], dimension[1]};
	}
	
	public static int[] getCellPositionFromDate(Calendar quarterStartDate, Calendar lectureDate, boolean isLectureEnd) {		
		int lectureDateOfWeek = lectureDate.get(Calendar.DAY_OF_WEEK);
		
		int daysBetween = (int)ChronoUnit.DAYS.between(quarterStartDate.toInstant(), lectureDate.toInstant());
		
		if (lectureDateOfWeek == Calendar.SUNDAY || lectureDateOfWeek == Calendar.SATURDAY || daysBetween > 81) {
			return null;
		}
		
		int xPosition = daysBetween % 28;
		xPosition -= xPosition / 7 * 2;
		xPosition += xPosition / 14 + 1;
		
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
		
		return new int[] {xPosition, yPosition};
	}
	
	public static int[] getCellDimensionFromCellPositions(int startXPosition, int startYPosition, int endXPosition, int endYPosition) {
		int width = 1 + endXPosition - startXPosition;
		int height = 1 + endYPosition - startYPosition;
		
		return new int[] {width, height};
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
		ExcelGenerator.writeFile(file, content);
	}

	public static void writeFile(File file, byte[] content) throws IOException {
		FileOutputStream out = new FileOutputStream(file);
		out.write(content);
		out.flush();
		out.close();
	}

	public static void saveNewFile(String fileName) throws IOException {
		File outFile = new File(fileName);
		File template = new File(ExcelGenerator.class.getClassLoader().getResource("template.xlsx").getFile());
		Workbook wbTemplate = ExcelGenerator.loadWorkbookFromFile(template);
		ExcelGenerator.saveWorkbookToFile(wbTemplate, outFile);
	}

}
