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
import classSchedulerExporter.ExportToXLS2;
import classSchedulerExporter.ExportToXLS3;
import classScheduler_src.Algorithm;
import classScheduler_src.ConfigParser;
import classScheduler_src.Week;

public class MainFrame extends JFrame {
	private static final long serialVersionUID = -2047309250954286106L;
	private ConfigParser cp;
	private int export_format = 1;
	private JLabel status = new JLabel("Status:"), status_text = new JLabel("Idle"), h_text = new JLabel(""),
			time_label = new JLabel("");

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
		menu = new JMenu("Tools");
		file = new ChangeExport("Change export format", this);
		menu.add(file);
		menuBar.add(menu);

		menu = new JMenu("About");
		file = new About("Usage");
		menu.add(file);
		menuBar.add(menu);
		this.setJMenuBar(menuBar);

		status.setBounds(10, 200, 40, 25);
		cp.add(status);
		time_label.setBounds(10, 180, 300, 25);
		time_label.setForeground(Color.BLUE);
		cp.add(time_label);
		status_text.setBounds(55, 200, 200, 25);
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
	}

	public void updateText(String input, int i) {
		if (i == 0)
			h_text.setText(input);
		else
			time_label.setText(input);
	}

	public void setExportFormat(int i) {
		export_format = i;
	}

	public void runApp() {
		if (cp != null) {
			status_text.setText("Running!");
			updateText("", 1);
			double start = System.currentTimeMillis();
			Week[][] t2 = new Algorithm(cp, this).runAlgorithm(0, 20000);
			if (export_format == 0)
				new ExportToXLS(t2);
			else if (export_format == 1)
				new ExportToXLS3(t2);
			else if (export_format == 2)
				new ExportToXLS2(t2);
			status_text.setText("Finished (Exported)!");
			double finish = System.currentTimeMillis();
			double time = finish - start;
			updateText("Koha e nevojitur per gjenerim: " + getTime(time), 1);
		} else
			JOptionPane.showMessageDialog(null, "Select config file first!");
	}

	private String getTime(double time) {
		int s;
		s = (int) time / 1000;
		return s + " sekonda";
	}
}