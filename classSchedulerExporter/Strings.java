package classSchedulerExporter;

import java.util.ArrayList;

public class Strings {
	private ArrayList<String> oret = new ArrayList<String>();
	private ArrayList<String> ditet = new ArrayList<String>();

	public Strings() {
		addOret();
		addDitet();
	}

	private void addOret() {
		oret.add("8:00-9:30");
		oret.add("9:45-11:15");
		oret.add("11:30-13:00");
		oret.add("13:15-14:45");
		oret.add("15:00-16:30");
		oret.add("16:45-18:15");
		oret.add("18:30-20:00");
	}
	private void addDitet() {
		ditet.add("E Hene");
		ditet.add("E Marte");
		ditet.add("E Merkure");
		ditet.add("E Enjte");
		ditet.add("E Premte");
	}
	public ArrayList<String> getOret(){
		return oret;
	}
	public ArrayList<String> getDitet(){
		return ditet;
	}
}
