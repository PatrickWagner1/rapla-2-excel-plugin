package semesterTimeTable.excel;

/**
 * The class can be used like a logging class for error messages.
 *
 */
public class ErrorOutput {

	/** The object where new error messages should be printed to */
	private Output output;

	/** The string of all error messages */
	private String errorOutput;

	/**
	 * Creates an error output object with no error messages. New error messages
	 * will be printed to the error output stream.
	 */
	public ErrorOutput() {
		this.errorOutput = "";
	}

	/**
	 * Creates an error output object with no error messages. New error messages
	 * will be printed to the given output object.
	 * 
	 * @param output The object for printing the error messages
	 */
	public ErrorOutput(Output output) {
		this.output = output;
		this.errorOutput = "";
	}

	/**
	 * Creates an error output object with an error message. New error messages will
	 * be printed to the error output stream.
	 * 
	 * @param errorMessage The first error message for the error output
	 */
	public ErrorOutput(String errorMessage) {
		this.errorOutput = errorMessage;
		this.printErrorMessage(errorMessage);
	}

	/**
	 * Creates an error output object with an error message. New error messages will
	 * be printed to the given output object.
	 * 
	 * @param output       The object for printing the error messages
	 * @param errorMessage The first error message for the error output
	 */
	public ErrorOutput(Output output, String errorMessage) {
		this.output = output;
		this.errorOutput = errorMessage;
		this.printErrorMessage(errorMessage);
	}

	/**
	 * Returns all error messages separated by a new line.
	 * 
	 * @return All error messages
	 */
	public String getErrorOutput() {
		return this.errorOutput;
	}

	/**
	 * Adds a single error message to the error output.
	 * 
	 * @param errorMessage The error message
	 */
	public void addErrorMessage(String errorMessage) {
		this.errorOutput += errorMessage + "\n";
		this.printErrorMessage(errorMessage);
	}

	/**
	 * Adds multiple error messages to the error output. Each message will be
	 * separated by a new line.
	 * 
	 * @param errorMessages An array of error messages
	 */
	public void addErrorMessages(String[] errorMessages) {
		String errorOutput = "";
		for (String errorMessage : errorMessages) {
			errorOutput += errorMessage + "\n";
		}
		this.errorOutput += errorOutput;
		this.printErrorMessage(errorOutput);
	}

	/**
	 * Removes all saved error messages.
	 */
	public void resetErrorOutput() {
		this.errorOutput = "";
	}

	/**
	 * Prints the given error message to the output object or to the error output
	 * stream, if the output object is null.
	 * 
	 * @param newErrorMessage The error message
	 */
	private void printErrorMessage(String newErrorMessage) {
		if (this.output != null) {
			this.output.println(newErrorMessage);
		} else {
			System.err.println(newErrorMessage);
		}
	}
}
