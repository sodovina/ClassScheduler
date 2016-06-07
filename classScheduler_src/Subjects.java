package classScheduler_src;

public class Subjects {
	private int year = 0, department = 0, haveChoosenSubject;
	private String subjectTitle, professor = "";
	private boolean lab;

	public Subjects(int department, int year, String subjectTitle, String professor, int haveChoosenSubject,
			boolean lab) {
		this.department = department;
		this.year = year;
		this.subjectTitle = subjectTitle;
		this.professor = professor;
		this.haveChoosenSubject = haveChoosenSubject;
		this.lab = lab;
	}

	public <E> E getC(String name, Class<E> type) {
		if (name.contains("studentsNumber"))
			return type.cast(haveChoosenSubject);
		else if (name.contains("lab"))
			return type.cast(lab);
		else if (name.contains("year"))
			return type.cast(year);
		else if (name.contains("subjectTitle"))
			return type.cast(subjectTitle);
		else if (name.contains("professor"))
			return type.cast(professor);
		else if (name.contains("department"))
			return type.cast(department);
		return null;
	}
}