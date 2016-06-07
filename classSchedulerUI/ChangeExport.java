package classSchedulerUI;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JFrame;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;

public class ChangeExport extends JMenuItem implements ActionListener {

	private static final long serialVersionUID = -5929635078457025410L;
	private MainFrame frame;

	public ChangeExport(String name, MainFrame frame) {
		super(name);
		this.frame = frame;
		addActionListener(this);
	}

	@Override
	public void actionPerformed(ActionEvent arg0) {
		// TODO Auto-generated method stub
		String[] choices = { "Export by departments", "Export by years(Default)", "Export all in one" };
		JFrame choice_frame = new JFrame("Input Dialog Example 3");
		String choice = (String) JOptionPane.showInputDialog(choice_frame, "Choose desired export format?",
				"Change export format", JOptionPane.QUESTION_MESSAGE, null, choices, choices[1]);
		if (choice != null) {
			if (choice.contains("Export by departments"))
				frame.setExportFormat(0);
			else if (choice.contains("Export by years(Default)"))
				frame.setExportFormat(1);
			else if (choice.contains("Export all in one"))
				frame.setExportFormat(2);
		}
	}
}
