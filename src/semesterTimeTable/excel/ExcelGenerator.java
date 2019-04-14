package semesterTimeTable.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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

	public Workbook loadWorkbookFromFile(File file) throws IOException {
		FileInputStream excelFile = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(excelFile);
		return workbook;
	}

	public Sheet getSheetFromWorkbook(Workbook workbook) {
		Sheet sheet = workbook.getSheetAt(0);
		return sheet;
	}

	public void saveWorkbookToFile(Workbook workbook, File file) throws IOException {
		ByteArrayOutputStream stream = new ByteArrayOutputStream();
		workbook.write(stream);
		byte[] content = stream.toByteArray();
		this.writeFile(file, content);
	}

	public void writeFile(File file, byte[] content) throws IOException {
		FileOutputStream out = new FileOutputStream(file);
		out.write(content);
		out.flush();
		out.close();
	}

	public void saveNewFile(String fileName) throws IOException {
		File outFile = new File(fileName);
		File template = new File("template.xlsx");
		Workbook wbTemplate = this.loadWorkbookFromFile(template);
		this.saveWorkbookToFile(wbTemplate, outFile);
	}
	
	public void createTemplateFromFile(File excelTemplate) throws IOException {
		Workbook workbook = this.loadWorkbookFromFile(excelTemplate);
		File template = new File("template.txt");
		this.saveWorkbookToFile(workbook, template);
	}
}
