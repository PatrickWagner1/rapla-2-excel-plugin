package semesterTimeTable.excel;

import java.time.LocalDate;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import de.jollyday.Holiday;
import de.jollyday.HolidayManager;
import de.jollyday.ManagerParameters;

public class Holidays {

	/**
	 * Returns a map of the calendar object and the name of a holiday between the
	 * startDate (inclusive) and the endDate (inclusive) in the given location of
	 * the locale variable.
	 * 
	 * The language of the locale doesn't matter and do not have any influence on
	 * the return statement. For example a new locale for the holidays in
	 * Baden-Wuerttemberg in Germany can be <code>new Locale("de","de","bw")</code>
	 * or <code>new Locale("en","de","bw)</code>
	 * 
	 * @param startDate The inclusive first relevant date for the holiday scanning
	 * @param endDate   The inclusive last relevant date for the holiday scanning
	 * @param locale    The locale for the location of the holiday scanning
	 * @return A map of the date and name of all holidays between the start and end
	 *         date
	 */
	public static Map<Calendar, String> getHolidays(Calendar startDate, Calendar endDate, Locale locale) {
		HolidayManager manager = HolidayManager.getInstance(ManagerParameters.create(locale.getCountry()));
		LocalDate start = Holidays.calendarToLocalDate(startDate);
		LocalDate end = Holidays.calendarToLocalDate(endDate);
		String[] variants = locale.getVariant().split("(_|-)");
		Set<Holiday> holidays = manager.getHolidays(start, end, variants);

		Map<Calendar, String> holidayMap = new TreeMap<Calendar, String>();
		for (Holiday holiday : holidays) {
			Calendar date = Holidays.localDateToCalendar(holiday.getDate());
			holidayMap.put(date, holiday.getDescription());
		}
		return holidayMap;
	}

	/**
	 * Converts a Calendar object to a Local Date object.
	 * 
	 * @param calendar The calendar to convert
	 * @return The converted localDate
	 */
	public static LocalDate calendarToLocalDate(Calendar calendar) {
		return LocalDate.of(calendar.get(Calendar.YEAR), calendar.get(Calendar.MONTH) + 1,
				calendar.get(Calendar.DAY_OF_MONTH));
	}

	/**
	 * Converts a Locale Date object to a Calendar object.
	 * 
	 * @param localDate The localDate to convert
	 * @return The converted calendar
	 */
	public static Calendar localDateToCalendar(LocalDate localDate) {
		return new GregorianCalendar(localDate.getYear(), localDate.getMonthValue() - 1, localDate.getDayOfMonth());
	}
}
