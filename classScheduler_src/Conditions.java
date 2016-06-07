package classScheduler_src;

public class Conditions {
	public boolean test(Week[][] matrix, ConfigParser cp, int day, int interval, int c_class, int start) {
		boolean free = matrix[day][interval].get("freeClass",Boolean.class,c_class);
		boolean lab = cp.getSubjects().get(start).getC("lab", Boolean.class) == matrix[day][interval].get("lab",Boolean.class,c_class);
		boolean capacity = matrix[day][interval].get("classCapacity",Integer.class,c_class) >= cp.getSubjects().get(start)
				.getC("studentsNumber", Integer.class);
		boolean prof = takenProfessor(matrix, cp.getSubjects().get(start), day, interval);
		boolean dept_year = checkDepartment_Year(matrix, cp.getSubjects().get(start), day, interval);
		return (free & lab & capacity & prof & dept_year);
	}

	public boolean takenProfessor(Week[][] matrx, Subjects sub, int day, int interval) {
		for (int k = 0; k < matrx[day][interval].getClasses().size(); k++) {
			if ((matrx[day][interval].get("professor",String.class,k)).equals(sub.getC("professor", String.class))) {
				return false;
			}
		}
		return true;
	}

	public boolean checkDepartment_Year(Week[][] matrx, Subjects sub, int day, int interval) {
		for (int k = 0; k < matrx[day][interval].getClasses().size(); k++) {
			if (sub.getC("department", Integer.class) == matrx[day][interval].get("department",Integer.class,k)
					&& sub.getC("year", Integer.class) == new Integer(matrx[day][interval].get("year",Integer.class,k)).intValue()) {
				return false;
			}
		}
		return true;
	}
}
