package semesterTimeTable.excel;

/**
 * The interface is to make sure, that a class implements a print method for
 * messages.
 */
public interface Output {

	/**
	 * Should print the given message somewhere.
	 * 
	 * Usually it should print the message in a new line.
	 * 
	 * @param message The message to print
	 */
	public void println(String message);
}
