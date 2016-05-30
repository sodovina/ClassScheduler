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

	public int getStudentsNumber() {
		return haveChoosenSubject;
	}

	public int getYear() {
		return year;
	}

	public int getDepartment() {
		return department;
	}

	public String getSubjectTitle() {
		return subjectTitle;
	}

	public boolean isLab() {
		return lab;
	}

	public String getProfessor() {
		return professor;
	}
}