package classScheduler_src;

import java.util.ArrayList;

public class Week {

	ArrayList<Classes> classes;

	public Week(ArrayList<Classes> k) {
		classes = k;
	}

	public void addSubjectToScheduler(int i, Subjects s) {
		classes.get(i).addSubject(s);
	}

	public int getDepartment(int i) {
		return classes.get(i).getDepartment();
	}

	public int getYear(int i) {
		return classes.get(i).getYear();
	}

	public boolean getFreeClass(int i) {
		return classes.get(i).freeClass();
	}

	public boolean isLab(int i) {
		return classes.get(i).isLab();
	}

	public String getClassName(int i) {
		return classes.get(i).getClassName();
	}

	public String getSubjectTitle(int i) {
		return classes.get(i).getSubjectTitle();
	}

	public String getProfessor(int i) {
		return classes.get(i).getProfessor();
	}

	public int getClassCapacity(int i) {
		return classes.get(i).getClassCapacity();
	}

	public int haveChoosenSubject(int i) {
		return classes.get(i).getStudentsNumber();
	}
}