package classSchedulerUI;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JMenuItem;

public class Run extends JMenuItem implements ActionListener {
	private static final long serialVersionUID = -6298225701850144295L;
	private MainFrame mf;

	public Run(String name, MainFrame mf) {
		super(name);
		this.mf = mf;
		addActionListener(this);
	}

	public void actionPerformed(ActionEvent arg0) {
		Thread t = new Thread() {
			public void run() {
				double start = System.currentTimeMillis();
				mf.runApp();
				double finish = System.currentTimeMillis();
				double time = finish - start;
				mf.updateText("Koha e nevojitur per gjenerim: " + getTime(time), 1);
			}
		};
		t.start();
	}

	private String getTime(double time) {
		int s;
		s = (int) time / 1000;
		return s + " sekonda";
	}
}
