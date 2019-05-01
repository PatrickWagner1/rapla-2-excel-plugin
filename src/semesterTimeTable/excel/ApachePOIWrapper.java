package semesterTimeTable.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Wrapper class for the Apache POI library.
 */
public final class ApachePOIWrapper implements IApachePOIWrapper {

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

}