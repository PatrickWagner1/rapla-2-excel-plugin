package semesterTimeTable.excel;

import java.awt.Frame;
import java.awt.TextArea;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

public class StandaloneFrame extends Frame {

	private static final long serialVersionUID = 1L;

	private static final String LINE_BREAK = "\n";

	private TextArea textArea;

	public StandaloneFrame() {
		this.setTitle("Rapla 2 Excel");
		this.setSize(500, 600);
		this.addWindowListener(new WindowListener());

		this.textArea = new TextArea("Waiting for CSV file selection...");
		this.textArea.setEditable(false);
		this.add(textArea);

		this.setVisible(true);
	}

	public void addTextLine(String text) {
		String newText = this.textArea.getText() + StandaloneFrame.LINE_BREAK + text;
		this.textArea.setText(newText);
	}

	public void close() {
		this.dispose();
		System.exit(0);
	}

	class WindowListener extends WindowAdapter {
		public void windowClosing(WindowEvent windowEvent) {
			windowEvent.getWindow().dispose();
			System.exit(0);
		}
	}
}
