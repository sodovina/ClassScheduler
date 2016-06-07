package classScheduler_src;

import java.util.ArrayList;

public class Week {

	private ArrayList<Classes> classes;

	public Week(ArrayList<Classes> k) {
		classes = k;
	}

	public ArrayList<Classes> getClasses() {
		return classes;
	}

	public void addSubjectToScheduler(int i, Subjects s) {
		classes.get(i).addSubject(s);
	}

	public <E> E get(String name, Class<E> type, int i) {
		if (name.contains("studentsNumber"))
			return type.cast(classes.get(i).getStudentsNumber());
		else if (name.contains("lab"))
			return type.cast(classes.get(i).isLab());
		else if (name.contains("year"))
			return type.cast(classes.get(i).getYear());
		else if (name.contains("subjectTitle"))
			return type.cast(classes.get(i).getSubjectTitle());
		else if (name.contains("professor"))
			return type.cast(classes.get(i).getProfessor());
		else if (name.contains("department"))
			return type.cast(classes.get(i).getDepartment());
		else if (name.contains("freeClass"))
			return type.cast(classes.get(i).freeClass());
		else if (name.contains("className"))
			return type.cast(classes.get(i).getClassName());
		else if (name.contains("classCapacity"))
			return type.cast(classes.get(i).getClassCapacity());
		return null;
	}
}