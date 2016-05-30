package classSchedulerExporter;

import java.util.ArrayList;

public class Strings {
	private ArrayList<String> clocks = new ArrayList<String>();
	private ArrayList<String> days = new ArrayList<String>();
	private ArrayList<String> years = new ArrayList<String>();
	private ArrayList<String> departments = new ArrayList<String>();

	public Strings() {
		addClocks();
		addDays();
		addYears();
		addDepartments();
	}

	private void addClocks() {
		clocks.add("8:00-9:30");
		clocks.add("9:45-11:15");
		clocks.add("11:30-13:00");
		clocks.add("13:15-14:45");
		clocks.add("15:00-16:30");
		clocks.add("16:45-18:15");
		clocks.add("18:30-20:00");
	}

	private void addDays() {
		days.add("E Hene");
		days.add("E Marte");
		days.add("E Merkure");
		days.add("E Enjte");
		days.add("E Premte");
	}

	private void addYears() {
		years.add("Viti i pare");
		years.add("Viti i dyte");
		years.add("Viti i trete");
	}

	private void addDepartments() {
		departments.add("Shkenca Kompjuterike");
		departments.add("Matematike");
		departments.add("Matematike Financiare");
	}

	public ArrayList<String> getClocks() {
		return clocks;
	}

	public ArrayList<String> getDays() {
		return days;
	}

	public ArrayList<String> getYears() {
		return days;
	}

	public ArrayList<String> getDepartments() {
		return departments;
	}

	public String getFileName() {
		return "Orari Mesimor";
	}
}
