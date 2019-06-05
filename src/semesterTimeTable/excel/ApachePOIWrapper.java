package semesterTimeTable.excel;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Wrapper class for the Apache POI library. Provides a simplified interface for
 * the POI subsystem.
 */
public final class ApachePOIWrapper {

	/**
	 * Saves the given workbook to the given file.
	 * 
	 * @param workbook The workbook to save
	 * @param file     The file for saving the workbook
	 * @throws IOException If saving the workbook in the file failed
	 */
	public static void saveWorkbookToFile(XSSFWorkbook workbook, File file) throws IOException {
		ByteArrayOutputStream stream = new ByteArrayOutputStream();
		workbook.write(stream);
		byte[] content = stream.toByteArray();
		writeFile(file, content);
	}

	/**
	 * Returns the workbook of the given file.
	 * 
	 * @param file The file of the workbook
	 * @return The workbook of the given file
	 * @throws IOException If reading the workbook failed
	 */
	public static XSSFWorkbook loadWorkbookFromFile(File file) throws IOException {
		FileInputStream excelFile = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
		return workbook;
	}

	/**
	 * Returns the workbook of the given input stream.
	 * 
	 * @param inputStream The input stream of the workbook
	 * @return The workbook of the given file
	 * @throws IOException If reading the workbook failed
	 */
	public static XSSFWorkbook loadWorkbookFromInputStream(InputStream inputStream) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		return workbook;
	}

	/**
	 * Converts an array of Colors into an array of XSSFColors.
	 * 
	 * @param colors The colors to be converted
	 * @return The array of XSSFColors
	 */
	public static XSSFColor[] colorsToXSSFColors(Color[] colors) {
		XSSFColor[] xssfColors = new XSSFColor[colors.length];
		int index = 0;
		for (Color color : colors) {
			xssfColors[index] = colorToXSSFColor(color);
			index++;
		}
		return xssfColors;
	}

	/**
	 * Converts a Color into a XSSFColor.
	 * 
	 * @param color The color to be converted
	 * @return The XSSFColor
	 */
	public static XSSFColor colorToXSSFColor(Color color) {
		return new XSSFColor(color, new DefaultIndexedColorMap());
	}

	/**
	 * Returns the first (and usually only) sheet of the lecture workbook.
	 * 
	 * @return The first sheet of the workbook
	 */
	public static XSSFSheet getSheet(XSSFWorkbook workbook) {
		return workbook.getSheetAt(0);
	}

	/**
	 * Resets a cell in a given workbook by setting the cell blank and setting the
	 * given cellStyle.
	 * 
	 * @param workbook  The target workbook
	 * @param rowNum    The (0 based) row number of the cell
	 * @param columnNum The (0 based) column number of the cell
	 * @param cellStyle The style for the cell
	 */
	public static void resetCell(XSSFWorkbook workbook, int rowNum, int columnNum, XSSFCellStyle cellStyle) {
		XSSFCell cell = ApachePOIWrapper.getSheet(workbook).getRow(rowNum).getCell(columnNum);
		cell.setBlank();
		cell.setCellStyle(cellStyle);
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
						font.setColor(cell.getCellStyle().getFont().getXSSFColor());
						fontMap.put(key, font);
					}
				}
			}

		}
		return fontMap;
	}

	/**
	 * Returns an array of all string values from a cell range.
	 * 
	 * Cells with no value or an invalid string value will not be added to the
	 * array.
	 * 
	 * @param sheet     The sheet, which will be scanned
	 * @param cellRange The cell range, which will be scanned
	 * @return An array of cell string values
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

	/**
	 * Returns an array of all numeric values from a cell range.
	 * 
	 * Cells with no value or an invalid numeric value will not be added to the
	 * array. Cells floating point numbers are casted to integers. This means all
	 * digits behind the point are cut off.
	 * 
	 * @param sheet     The sheet, which will be scanned
	 * @param cellRange The cell range, which will be scanned
	 * @return An array of cell integer values
	 */
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

	/**
	 * Returns an array of all date values from a cell range.
	 * 
	 * Cells with no value or an invalid date value will not be added to the array.
	 * 
	 * @param sheet     The sheet, which will be scanned
	 * @param cellRange The cell range, which will be scanned
	 * @return An array of cell date values
	 */
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

	/**
	 * Saves the given content to the given file.
	 * 
	 * @param file    The file for saving the content
	 * @param content The content to save
	 * @throws IOException If saving the content in the file failed
	 */
	public static void writeFile(File file, byte[] content) throws IOException {
		FileOutputStream out = new FileOutputStream(file);
		out.write(content);
		out.flush();
		out.close();
	}

	/**
	 * Copy all relevant font properties from the given font to copy to the font.
	 * 
	 * @param font       The font to copy the properties to
	 * @param fontToCopy The font to copy the properties from
	 */
	public static void copyFont(XSSFFont font, XSSFFont fontToCopy) {
		font.setBold(fontToCopy.getBold());
		font.setItalic(fontToCopy.getItalic());
		font.setColor(fontToCopy.getXSSFColor());
		font.setFontHeight(fontToCopy.getFontHeight());
		font.setFontName(fontToCopy.getFontName());
		font.setStrikeout(fontToCopy.getStrikeout());
		font.setUnderline(fontToCopy.getUnderline());
	}

}