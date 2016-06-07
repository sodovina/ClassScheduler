package classScheduler_src;

import java.util.Random;
import classSchedulerUI.MainFrame;

public class Algorithm {
	private Random rand = new Random();
	private Conditions cond = new Conditions();
	private ConfigParser cp;
	private MainFrame frame;
	private Week[][] matrix;

	public Algorithm(ConfigParser cp, MainFrame frame) {
		this.cp = cp;
		this.frame = frame;
	}

	public Week[][] runAlgorithm(int start, int finish) {
		Week[][] current = fill(matrix = newMatrix(), 0);
		while (++start < finish) {
			Week[][] next = fill(matrix = newMatrix(), 0);
			current = function(next, 0) < function(current, 0) ? current : next;
			frame.updateText("" + function(current, 0), 0);
		}
		return current;
	}

	private Week[][] fill(Week[][] fillMatrix, int start) {
		int day = rand.nextInt(5), interval = rand.nextInt(7), c_class = rand.nextInt(cp.getRooms().size());
		if (cond.test(matrix, cp, day, interval, c_class, start)) {
			fillMatrix[day][interval].addSubjectToScheduler(c_class, cp.getSubjects().get(start));
			start++;
		}
		return (start < cp.getSubjectsCount() ? fill(fillMatrix, start) : fillMatrix);
	}

	private Week[][] newMatrix() {
		Week[][] matrx = new Week[5][7];
		for (int i = 0; i < matrx.length; i++) {
			for (int j = 0; j < matrx[0].length; j++) {
				matrx[i][j] = new Week(cp.getRooms());
			}
		}
		return matrx;
	}

	private double function(Week[][] checkMatrix, double maxValue) {
		for (int i = 0; i < checkMatrix.length; i++) {
			for (int j = 0; j < checkMatrix[0].length; j++) {
				for (int k = 0; k < checkMatrix[i][j].getClasses().size(); k++) {
					if (checkMatrix[i][j].get("freeClass",Boolean.class,k))
						continue;
					maxValue += (checkMatrix[i][j].get("studentsNumber",Integer.class,k) * 1.0
							/ checkMatrix[i][j].get("classCapacity",Integer.class,k)*1.0);
				}
			}
		}
		return maxValue;
	}
}
