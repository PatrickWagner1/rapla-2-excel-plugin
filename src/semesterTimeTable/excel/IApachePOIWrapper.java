package semesterTimeTable.excel;

import java.io.File;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Interface for the Apache POI wrapper class.
 */
public interface IApachePOIWrapper {

	public static void saveWorkbookToFile(XSSFWorkbook workbook, File file) throws IOException {
	}

	public static void writeFile(File file, byte[] content) throws IOException {
	}

}
