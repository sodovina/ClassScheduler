package classScheduler_src;

public class Classes {
	// Department = 1 Shkenca Kompjuterike
	// Department = 2 Matematike
	// Department = 3 Matematike Financiare
	private boolean lab, freeClass = true;
	private String className, professor = "", subjectTitle;
	private int number_of_students = 0, department = 0, classCapacity = 0, year = 0;

	public Classes(String className, int classCapacity, boolean lab) {
		this.className = className;
		this.lab = lab;
		this.classCapacity = classCapacity;
	}

	public void addSubject(Subjects sub) {
		professor = sub.getC("professor", String.class);
		subjectTitle = sub.getC("subjectTitle", String.class);
		department = sub.getC("department", Integer.class);
		year = sub.getC("year", Integer.class);
		number_of_students = sub.getC("studentsNumber", Integer.class);
		freeClass = false;
	}

	public int getStudentsNumber() {
		return number_of_students;
	}

	public int getClassCapacity() {
		return classCapacity;
	}

	public String getClassName() {
		return className;
	}

	public boolean freeClass() {
		return freeClass;
	}

	public boolean isLab() {
		return lab;
	}

	public int getYear() {
		return year;
	}

	public String getSubjectTitle() {
		return subjectTitle;
	}

	public String getProfessor() {
		return professor;
	}

	public int getDepartment() {
		return department;
	}
}