package classSchedulerUI;

import java.awt.Color;
import java.awt.Container;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.SwingConstants;

import classSchedulerExporter.ExportToXLS;
import classScheduler_src.Algorithm;
import classScheduler_src.ConfigParser;
import classScheduler_src.Week;

public class MainFrame extends JFrame {
	/**
	 * 
	 */
	private static final long serialVersionUID = -2047309250954286106L;
	private ConfigParser cp;
	private JLabel status = new JLabel("Status:");
	private JLabel status_text = new JLabel("Idle");
	private JLabel h_text = new JLabel("");

	public MainFrame() {
		Container cp = getContentPane();
		cp.setLayout(null);

		JMenuBar menuBar = new JMenuBar();
		JMenu menu = new JMenu("File");
		JMenuItem file = new Run("Run", this);
		menu.add(file);
		file = new Open("Open", this);
		menu.add(file);
		file = new ExitMenu("Exit");
		menu.add(file);
		menuBar.add(menu);
		menu = new JMenu("About");
		file = new About("Usage");
		menu.add(file);
		menuBar.add(menu);
		this.setJMenuBar(menuBar);

		status.setBounds(10, 200, 40, 25);
		cp.add(status);
		status_text.setBounds(55, 200, 150, 25);
		status_text.setForeground(Color.red);
		cp.add(status_text);
		status = new JLabel("h(x):");
		status.setBounds(210, 200, 25, 25);
		cp.add(status);
		h_text.setBounds(240, 200, 150, 25);
		h_text.setForeground(Color.red);
		cp.add(h_text);

		JLabel disclaimer = new JLabel(
				"<html>Info!<br>Gjeneruesi i orarit mesimor per " + "drejtimet Shkenca Kompjuterike, Matematike dhe "
						+ "Matematike Financiare. Bazohet ne algoritmin Random "
						+ "Restart Hill Climbing. Per ta ekzekutuar programin duhet te importoni nje *config file" + ""
						+ "(i cili duhet te permbaj infot e sallave, profesoreve, lendeve etj.) Pas gjenerimit shkruan rezultatet "
								+ "duke perdorur JXL API.</html>",
				SwingConstants.CENTER);
		disclaimer.setBounds(10, 0, 350, 150);
		cp.add(disclaimer);

		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setSize(400, 300);
		setLocationRelativeTo(null);
		setTitle("Class Scheduler - FSHMN");
		setVisible(true);
	}

	public void getPath(String path) {
		cp = new ConfigParser(path);
		status_text.setText("Config file Loaded!");
		revalidate();
		repaint();
	}

	public void updateText(String input) {
		h_text.setText(input);
		revalidate();
		repaint();
	}

	public void runApp() {
		if (cp != null) {
			status_text.setText("Running!");
			revalidate();
			repaint();
			Algorithm a = new Algorithm(cp, this);
			Week[][] k = new Week[5][7];
			k = a.runAlgorithm();
			status_text.setText("Finished (Exporting)!");
			new ExportToXLS(k);
			status_text.setText("Finished (Exported)!");
		} else
			JOptionPane.showMessageDialog(null, "Select config file first!");
	}
}