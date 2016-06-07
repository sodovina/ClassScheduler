package classSchedulerUI;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JMenuItem;

public class Run extends JMenuItem implements ActionListener {
	private static final long serialVersionUID = -6298225701850144295L;
	private MainFrame frame;

	public Run(String name, MainFrame mf) {
		super(name);
		this.frame = mf;
		addActionListener(this);
	}

	public void actionPerformed(ActionEvent arg0) {
		Thread t = new Thread() {
			public void run() {
				frame.runApp();
			}
		};
		t.start();
	}

	
}
