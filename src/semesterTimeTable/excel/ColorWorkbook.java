package semesterTimeTable.excel;

import java.io.File;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ColorWorkbook {
	
	private final static String TEMPLATE_FILENAME = "colorMap.xlsx";
	private Map<String, XSSFColor[]> colorPairs;
	private Map<String, XSSFFont> highlightedFonts;
	private String[] ignorePrefixes;
	
	public ColorWorkbook() throws IOException {
		this.setVariablesFromWorkbook(LectureWorkbook
				.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)));
	}

	public ColorWorkbook(String pathToTemplate) throws IOException {
		File file = new File(pathToTemplate, ColorWorkbook.TEMPLATE_FILENAME);
		if (file.exists()) {
			this.setVariablesFromWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.setVariablesFromWorkbook(LectureWorkbook.loadWorkbookFromFile(
					LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)));
		}
	}
	
	public ColorWorkbook(String pathToTemplate, String templateFilename) throws IOException {
		File file = new File(pathToTemplate, templateFilename);
		if (file.exists()) {
			this.setVariablesFromWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.setVariablesFromWorkbook(LectureWorkbook.loadWorkbookFromFile(
					LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)));
		}
	}
	
	public Map<String, XSSFColor[]> getColorPairs() {
		return this.colorPairs;
	}
	
	public Map<String, XSSFFont> getHighlightedFonts() {
		return this.highlightedFonts;
	}
	
	public String[] getIngorePrefixes() {
		return this.ignorePrefixes;
	}
	
	private void setVariablesFromWorkbook(XSSFWorkbook workbook) throws IOException {
		XSSFSheet sheet = workbook.getSheetAt(0);
		int lastRowNum = sheet.getLastRowNum();
		this.colorPairs = LectureWorkbook.getMappedColorPairs(sheet, 2, lastRowNum, 0, 0);
		this.highlightedFonts = LectureWorkbook.getMappedFontColor(sheet, 2,lastRowNum, 1, 1);
		this.ignorePrefixes = LectureWorkbook.getValuesFromWorkbook(sheet, 2, lastRowNum, 2, 2);
		workbook.close();
	}
}
