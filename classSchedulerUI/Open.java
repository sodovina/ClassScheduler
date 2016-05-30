package classSchedulerUI;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.JMenuItem;
import javax.swing.filechooser.FileNameExtensionFilter;

public class Open extends JMenuItem implements ActionListener {
	private static final long serialVersionUID = 2685806301646457391L;
	private MainFrame frame;

	public Open(String name, MainFrame frame) {
		super(name);
		this.frame = frame;
		addActionListener(this);
	}

	public void actionPerformed(ActionEvent arg0) {
		File file;
		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setCurrentDirectory(new File(System.getProperty("user.dir")));
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Config files(*.cfg)", "cfg", "text");
		fileChooser.setFileFilter(filter);
		int result = fileChooser.showOpenDialog(this);
		if (result == JFileChooser.APPROVE_OPTION) {
			file = fileChooser.getSelectedFile();
			frame.getPath(file.getAbsolutePath());
		}
	}
}
