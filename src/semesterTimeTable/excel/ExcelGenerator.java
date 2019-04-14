package semesterTimeTable.excel;

import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeMap;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.IndexedColors;

public class ExcelGenerator {

	private List<Lecture> lectures;
	
	private IndexedColors[] lectureColors;
	
	public ExcelGenerator(List<Lecture> lectures) {
		// TODO Add more useful colors to the lectureColors array and maybe delete absurd colors
		this.lectureColors = new IndexedColors[] {IndexedColors.BLUE, IndexedColors.AQUA, IndexedColors.GREEN};
		this.lectures = generateLecturesGroupId(lectures);
	}
	
	private List<Lecture> generateLecturesGroupId(List<Lecture> lectures) {
		Map<String, List<Lecture>> groupedLectures = new TreeMap<String, List<Lecture>>(lectures.stream().collect(Collectors.groupingBy(Lecture::getName)));
		List<Lecture> currentLectureList;
		int currentGroupId = 2;
		for (Entry<String, List<Lecture>> groupedLecture : groupedLectures.entrySet()) {
			currentLectureList = groupedLecture.getValue();
			if (groupedLecture.getKey().startsWith(Lecture.REPEAT_EXAM_START_STRING)) {
				for (Lecture currentLecture: currentLectureList) {
					currentLecture.setGroupId(Lecture.REPEAT_EXAM_ID);
				}
			} else {
				for (Lecture currentLecture: currentLectureList) {
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
	
	
}
