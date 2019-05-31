package semesterTimeTable.excel.standalone;

import java.awt.Frame;
import java.awt.Image;
import java.awt.TextArea;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import semesterTimeTable.excel.Output;

public class StandaloneFrame extends Frame implements Output {

	/** Default serial version UID */
	private static final long serialVersionUID = 1L;

	/** Line break in files */
	private static final String LINE_BREAK = "\n";

	/** The text area in the frame for printing status messages */
	private TextArea textArea;

	/**
	 * Creates the main Frame for the GUI.
	 */
	public StandaloneFrame() {
		this.setTitle("Rapla 2 Excel");
		this.setSize(500, 600);
		this.addWindowListener(new WindowListener());

		this.textArea = new TextArea("Waiting for CSV file selection...");
		this.textArea.setEditable(false);
		this.add(textArea);

		List<Image> iconImages = StandaloneFrame
				.getImageListFromResources(new String[] { "rapla_32x32.png", "rapla_64x64.png", "rapla_128x128.png" });
		this.setIconImages(iconImages);
		this.setVisible(true);
	}

	/**
	 * Adds the given text in a new line to the text area.
	 * 
	 * @param text The text to add to the text area
	 */
	public void println(String text) {
		String newText = this.textArea.getText() + StandaloneFrame.LINE_BREAK + text;
		this.textArea.setText(newText);
	}

	/**
	 * Closes the Frame end exits the program.
	 */
	public void close() {
		this.dispose();
		System.exit(0);
	}

	/**
	 * The window listener is a simple window adapter that closes the window and
	 * exits the program, if clicking the close button.
	 */
	class WindowListener extends WindowAdapter {

		/**
		 * Closes the window and exits the program, if clicking the close button.
		 */
		public void windowClosing(WindowEvent windowEvent) {
			windowEvent.getWindow().dispose();
			System.exit(0);
		}
	}

	/**
	 * Returns a list of images from the given image names. The images has to be
	 * located in the root source folder of the class.
	 * 
	 * Only images which will be find by the image name will be added to the list.
	 * Otherwise the image will be skipped.
	 * 
	 * @param imageNames The names of the images in the resources
	 * @return A list of images
	 */
	private static List<Image> getImageListFromResources(String[] imageNames) {
		List<Image> imageList = new ArrayList<Image>();
		for (String imageName : imageNames) {
			try {
				Image image = ImageIO.read(StandaloneFrame.class.getClassLoader().getResourceAsStream(imageName));
				imageList.add(image);
			} catch (IOException | IllegalArgumentException e) {
				System.err.println("Cannot add image \"" + imageName + "\" to the image list");
			}
		}
		return imageList;
	}
}
