package semesterTimeTable.excel;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * Objects of this class can be used to save excel cell style and font
 * informations for a specific lecture.
 */
public class LectureProperties {

	/** The lecture name */
	private String lectureName;

	/** The short lecture name */
	private String shortLectureName;

	/** The font for the lecture cell */
	private XSSFFont font;

	/** The fill color for the lecture cell */
	private XSSFColor fillColor;

	/**
	 * Creates an object with excel cell style and font for a lecture.
	 * 
	 * @param lectureName      The name of the lecture
	 * @param shortLectureName The short name of the lecture
	 * @param font             The font of the lecture cell
	 * @param fillColor        The fill color of the lecture cell
	 */
	public LectureProperties(String lectureName, String shortLectureName, XSSFFont font, XSSFColor fillColor) {
		this.setLectureName(lectureName);
		this.setShortLectureName(shortLectureName);
		this.setFont(font);
		this.setFillColor(fillColor);
	}

	/**
	 * Returns the name of the lecture.
	 * 
	 * @return The lecture name
	 */
	public String getLectureName() {
		return this.lectureName;
	}

	/**
	 * Sets the name of the lecture.
	 * 
	 * @param lectureName The name for the lecture
	 */
	public void setLectureName(String lectureName) {
		this.lectureName = lectureName;
	}

	/**
	 * Returns the short name of the lecture.
	 * 
	 * @return The short lecture name
	 */
	public String getShortLectureName() {
		return this.shortLectureName;
	}

	/**
	 * Sets the short name of the lecture.
	 * 
	 * @param shortLectureName The short name for the lecture
	 */
	public void setShortLectureName(String shortLectureName) {
		this.shortLectureName = shortLectureName;
	}

	/**
	 * Returns the font of the lecture excel cell.
	 * 
	 * @return The lecture cell font
	 */
	public XSSFFont getFont() {
		return this.font;
	}

	/**
	 * Sets the font for the lecture excel cell.
	 * 
	 * @param font The font for the lecture excel cell
	 */
	public void setFont(XSSFFont font) {
		this.font = font;
	}

	/**
	 * Returns the fill color of the lecture excel cell.
	 * 
	 * @return The lecture excel cell fill color
	 */
	public XSSFColor getFillColor() {
		return this.fillColor;
	}

	/**
	 * Sets the fill color for the lecture excel cell.
	 * 
	 * @param fillColor The fill color for the lecture excel cell
	 */
	public void setFillColor(XSSFColor fillColor) {
		this.fillColor = fillColor;
	}
}
