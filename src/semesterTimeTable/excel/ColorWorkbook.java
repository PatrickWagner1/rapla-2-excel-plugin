package semesterTimeTable.excel;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ColorWorkbook {

	private final static String REMOVE_PREFIX = "WKL";
	private final static String IGNORE_PREFIX = "Klausur";

	private final static String TEMPLATE_FILENAME = "colorMap.xlsx";

	private Map<String, XSSFColor[]> colorPairs;
	private Map<String, XSSFFont> highlightedFonts;
	private String[] ignorePrefixes;

	public ColorWorkbook() throws IOException {
		this.initColorWorkbook(
				LectureWorkbook.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)));
	}

	public ColorWorkbook(String pathToTemplate) throws IOException {
		File file = new File(pathToTemplate, ColorWorkbook.TEMPLATE_FILENAME);
		if (file.exists()) {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.initColorWorkbook(LectureWorkbook
					.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)));
		}
	}

	public ColorWorkbook(String pathToTemplate, String[] lectureNames) throws IOException {
		File file = new File(pathToTemplate, ColorWorkbook.TEMPLATE_FILENAME);
		if (file.exists()) {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(
					LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)), file, lectureNames);
		}
	}

	public ColorWorkbook(String pathToTemplate, String templateFilename) throws IOException {
		File file = new File(pathToTemplate, templateFilename);
		if (file.exists()) {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.initColorWorkbook(LectureWorkbook
					.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)));
		}
	}

	public ColorWorkbook(String pathToTemplate, String templateFilename, String[] lectureNames) throws IOException {
		File file = new File(pathToTemplate, templateFilename);
		if (file.exists()) {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(
					LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)), file, lectureNames);
		}
	}

	public Map<String, XSSFColor[]> getColorPairs() {
		return this.colorPairs;
	}

	public Map<String, XSSFFont> getHighlightedFonts() {
		return this.highlightedFonts;
	}

	public String[] getIgnorePrefixes() {
		return this.ignorePrefixes;
	}

	private void initColorWorkbook(XSSFWorkbook workbook, File file, String[] lectureNames) throws IOException {
		List<String> lectureNameList = Arrays.asList(lectureNames);
		XSSFSheet sheet = workbook.getSheetAt(0);
		String removePrefix = ColorWorkbook.REMOVE_PREFIX + " ";
		String ignorePrefix = ColorWorkbook.IGNORE_PREFIX + " ";
		String rawLectureName;
		int rowNum = 3;
		for (String lectureName : lectureNameList) {
			if (!lectureName.startsWith(removePrefix)) {
				rawLectureName = lectureName.startsWith(ignorePrefix) ? lectureName.substring(ignorePrefix.length())
						: lectureName;
				
				if (rawLectureName.equals(lectureName) || !lectureNameList.contains(rawLectureName)) {
					XSSFRow row = sheet.getRow(rowNum);
					row = row == null ? sheet.createRow(rowNum) : row;
					XSSFCell cell = row.getCell(0);
					cell = cell == null ? row.createCell(0) : cell;
					cell.setCellValue(rawLectureName);
					rowNum++;
				}
			}
		}
		LectureWorkbook.saveWorkbookToFile(workbook, file);
		this.initColorWorkbook(workbook);
	}

	private void initColorWorkbook(XSSFWorkbook workbook) throws IOException {
		XSSFSheet sheet = workbook.getSheetAt(0);
		int lastRowNum = sheet.getLastRowNum();
		this.colorPairs = LectureWorkbook.getMappedColorPairs(sheet, 2, lastRowNum, 0, 0);
		this.highlightedFonts = LectureWorkbook.getMappedFontColor(sheet, 2, lastRowNum, 1, 1);
		this.ignorePrefixes = LectureWorkbook.getValuesFromWorkbook(sheet, 2, lastRowNum, 2, 2);
		workbook.close();
	}
}
