package semesterTimeTable.excel;

import org.apache.poi.xssf.usermodel.XSSFColor;

public class LectureProperties {
	
	private String lectureName;
	
	private String shortLectureName;

	private XSSFColor fontColor;
	
	private XSSFColor fillColor;
	
	public LectureProperties(String lectureName, String shortLectureName, XSSFColor fontColor, XSSFColor fillColor) {
		this.setLectureName(lectureName);
		this.setShortLectureName(shortLectureName);
		this.setFontColor(fontColor);
		this.setFillColor(fillColor);
	}

	public String getLectureName() {
		return this.lectureName;
	}

	public void setLectureName(String lectureName) {
		this.lectureName = lectureName;
	}

	public String getShortLectureName() {
		return this.shortLectureName;
	}

	public void setShortLectureName(String shortLectureName) {
		this.shortLectureName = shortLectureName;
	}

	public XSSFColor getFontColor() {
		return this.fontColor;
	}

	public void setFontColor(XSSFColor fontColor) {
		this.fontColor = fontColor;
	}

	public XSSFColor getFillColor() {
		return this.fillColor;
	}

	public void setFillColor(XSSFColor fillColor) {
		this.fillColor = fillColor;
	}
}
