package classSchedulerUI;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;

public class About extends JMenuItem implements ActionListener {

	private static final long serialVersionUID = 5148288717787921514L;

	public About(String name) {
		super(name);
		addActionListener(this);
	}

	public void actionPerformed(ActionEvent arg0) {
		JOptionPane.showMessageDialog(null, "Detyra 1 nga lenda Intelegjenc Artificiale"
				+ "\nFajlli config duhet te permbaj:\n" + "\n#Lendet - Hashtagu hapes\ndrejtimi = 3\nviti = 3"
				+ "\nlenda = emertimi_lendes\nprof = profesori"
				+ "\nnumri_studentave = 30\nlab = false\n#end - Hashtagu mbylles"
				+ "\n\nSi dhe sallat(klaset)\n\n#Klaset\nemertimi = Salla 112\nkapaciteti = 30\nlab = false\n#end");
	}
}
