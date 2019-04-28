package semesterTimeTable.excel;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class containing style (especially color) configurations for a workbook.
 *
 */
public class ColorWorkbook {

	/** Prefix for removing all lectures starting with this prefix */
	private final static String REMOVE_PREFIX = "WKL";

	/** Prefix for ignoring in each lecture */
	private final static String IGNORE_PREFIX = "Klausur";

	/** Filename of configuration template */
	private final static String TEMPLATE_FILENAME = "colorMap.xlsx";

	/**
	 * Map of values and their color pair (color pair always consists of font color
	 * and fill color)
	 */
	private Map<String, XSSFColor[]> colorPairs;

	/** Map of values and their font to highlight the value with */
	private Map<String, XSSFFont> highlightedFonts;

	/**
	 * Array of all prefixes, which should be ignored by the grouping of lectures
	 */
	private String[] ignorePrefixes;

	/**
	 * Loads the style configuration of the configuration template.
	 * 
	 * @throws IOException If loading the template file fails
	 */
	public ColorWorkbook() throws IOException {
		this.initColorWorkbook(
				LectureWorkbook.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)));
	}

	/**
	 * Loads the custom style configuration from the given path, if there is a file
	 * in the path with the {@link ColorWorkbook#TEMPLATE_FILENAME} as filename.
	 * Otherwise the style configuration of the template loads.
	 * 
	 * @param pathToTemplate The path to the custom template
	 * @throws IOException If reading one file failed
	 */
	public ColorWorkbook(String pathToTemplate) throws IOException {
		File file = new File(pathToTemplate, ColorWorkbook.TEMPLATE_FILENAME);
		if (file.exists()) {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.initColorWorkbook(LectureWorkbook
					.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)));
		}
	}

	/**
	 * Loads the custom style configuration from the given path, if there is a file
	 * in the path with the {@link ColorWorkbook#TEMPLATE_FILENAME} as filename.
	 * Otherwise the style configuration of the template loads and creates a new
	 * configuration file containing the given lecture names in the given path with
	 * the {@link ColorWorkbook#TEMPLATE_FILENAME} as filename.
	 * 
	 * All lecture names starts with the {@link ColorWorkbook#REMOVE_PREFIX} are not
	 * added to the configuration file.
	 * 
	 * All lecture names starts with the {@link ColorWorkbook#ignorePrefixes} are
	 * added without the prefix. If there exists a lecture name with this new name,
	 * it is not added.
	 * 
	 * @param pathToTemplate The path to the custom template
	 * @param lectureNames   A list of lecture names
	 * @throws IOException If reading one file or saving the configuration file
	 *                     failed
	 */
	public ColorWorkbook(String pathToTemplate, List<String> lectureNames) throws IOException {
		File file = new File(pathToTemplate, ColorWorkbook.TEMPLATE_FILENAME);
		if (file.exists()) {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(
					LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)), file, lectureNames);
		}
	}

	/**
	 * Loads the custom style configuration from the given path, if there is a file
	 * in the path with the given template filename. Otherwise the style
	 * configuration of the template loads.
	 * 
	 * @param pathToTemplate   The path to the custom template
	 * @param templateFilename The name of the custom template
	 * @throws IOException If reading one file failed
	 */
	public ColorWorkbook(String pathToTemplate, String templateFilename) throws IOException {
		File file = new File(pathToTemplate, templateFilename);
		if (file.exists()) {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.initColorWorkbook(LectureWorkbook
					.loadWorkbookFromFile(LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)));
		}
	}

	/**
	 * Loads the custom style configuration from the given path, if there is a file
	 * in the path with the given template filename. Otherwise the style
	 * configuration of the template loads and creates a new configuration file
	 * containing the given lecture names in the given path with the given filename.
	 * 
	 * All lecture names starts with the {@link ColorWorkbook#REMOVE_PREFIX} are not
	 * added to the configuration file.
	 * 
	 * All lecture names starts with the {@link ColorWorkbook#ignorePrefixes} are
	 * added without the prefix. If there exists a lecture name with this new name,
	 * it is not added.
	 * 
	 * @param pathToTemplate   The path to the custom template
	 * @param templateFilename The name of the custom template
	 * @param lectureNames     A list of lecture names
	 * @throws IOException If reading one file or saving the configuration file
	 *                     failed
	 */
	public ColorWorkbook(String pathToTemplate, String templateFilename, List<String> lectureNames) throws IOException {
		File file = new File(pathToTemplate, templateFilename);
		if (file.exists()) {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(file));
		} else {
			this.initColorWorkbook(LectureWorkbook.loadWorkbookFromFile(
					LectureWorkbook.getTemplateFile(ColorWorkbook.TEMPLATE_FILENAME)), file, lectureNames);
		}
	}

	/**
	 * Returns a map of values and their color pair (color pair always consists of
	 * font color and fill color).
	 * 
	 * @return A map of color pairs
	 */
	public Map<String, XSSFColor[]> getColorPairs() {
		return this.colorPairs;
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
	 * Creates a new configuration file with the configuration template and the
	 * given lectures and sets all the configurable variables.
	 * 
	 * All lecture names starts with the {@link ColorWorkbook#REMOVE_PREFIX} are not
	 * added to the configuration file.
	 * 
	 * All lecture names starts with the {@link ColorWorkbook#ignorePrefixes} are
	 * added without the prefix. If there exists a lecture name with this new name,
	 * it is not added.
	 * 
	 * @param workbook     The template configuration workbook
	 * @param file         The file to save the configuration workbook into
	 * @param lectureNames A list of lecture names
	 * @throws IOException If saving the configuration workbook or closing the
	 *                     workbook failed
	 */
	private void initColorWorkbook(XSSFWorkbook workbook, File file, List<String> lectureNames) throws IOException {
		XSSFSheet sheet = workbook.getSheetAt(0);
		String removePrefix = ColorWorkbook.REMOVE_PREFIX + " ";
		String ignorePrefix = ColorWorkbook.IGNORE_PREFIX + " ";
		String rawLectureName;
		int rowNum = 3;
		for (String lectureName : lectureNames) {
			if (!lectureName.startsWith(removePrefix)) {
				rawLectureName = lectureName.startsWith(ignorePrefix) ? lectureName.substring(ignorePrefix.length())
						: lectureName;

				if (rawLectureName.equals(lectureName) || !lectureNames.contains(rawLectureName)) {
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

	/**
	 * Sets all the configurable variables.
	 * 
	 * @param workbook The template configuration workbook
	 * @throws IOException If closing the workbook failed
	 */
	private void initColorWorkbook(XSSFWorkbook workbook) throws IOException {
		XSSFSheet sheet = workbook.getSheetAt(0);
		int lastRowNum = sheet.getLastRowNum();
		this.colorPairs = LectureWorkbook.getMappedColorPairs(sheet, new CellRangeAddress(2, lastRowNum, 0, 0));
		this.highlightedFonts = LectureWorkbook.getMappedFontColor(sheet, new CellRangeAddress(2, lastRowNum, 1, 1));
		this.ignorePrefixes = LectureWorkbook.getValuesFromWorkbook(sheet, new CellRangeAddress(2, lastRowNum, 2, 2));
		workbook.close();
	}
}
