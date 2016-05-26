package classSchedulerUI;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JMenuItem;

public class ExitMenu extends JMenuItem implements ActionListener{
	/**
	 * 
	 */
	private static final long serialVersionUID = -6783544302949917244L;

	public ExitMenu(String name){
		super(name);
		addActionListener(this);
	}

	@Override
	public void actionPerformed(ActionEvent arg0) {
		System.exit(0);
	}
}
