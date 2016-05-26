package classScheduler_src;

public class Conditions {
	public boolean test(Week[][] interval, ConfigParser cp, int i, int j, int k, int l) {
		boolean free = interval[i][j].getFreeClass(k);
		boolean lab = cp.getSubjects().get(l).isLab() == interval[i][j].isLab(k);
		boolean capacity = interval[i][j].getClassCapacity(k) >= cp.getSubjects().get(l).getStudentsNumber();
		boolean prof = takenProfessor(interval, cp.getSubjects().get(l), i, j);
		boolean dept_year = checpDepartment_Year(interval, cp.getSubjects().get(l), i, j);
		return (free & lab & capacity & prof & dept_year);
	}

	public boolean takenProfessor(Week[][] interval, Subjects s, int i, int j) {
		for (int k = 0; k < (interval[i][j].classes).size(); k++) {
			if ((interval[i][j].getProfessor(k) == null) && (interval[i][j].getProfessor(k)).equals(s.getProfessor())) {
				return false;
			}
		}
		return true;
	}

	public boolean checpDepartment_Year(Week[][] interval, Subjects s, int i, int j) {
		for (int k = 0; k < 7; k++) {
			if (s.getDepartment() == interval[i][j].getDepartment(k)
					&& s.getYear() == new Integer(interval[i][j].getYear(k)).intValue()) {
				return false;
			}
		}
		return true;
	}
}
