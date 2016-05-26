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

	public Week[][] runAlgorithm() {
		matrix = newMatrix();
		int iteration = 0;
		matrix = fill(matrix);
		Week[][] current = matrix;
		matrix = newMatrix();
		Week[][] next;
		while (iteration < 100000) {
			int l = 0;
			while (l < cp.getSubjectsCount()) {
				int dita = rand.nextInt(5), kolona = rand.nextInt(cp.getRooms().size()), b = rand.nextInt(7);
				if (cond.test(matrix, cp, dita, b, kolona, l)) {
					matrix[dita][b].addSubjectToScheduler(kolona, cp.getSubjects().get(l));
					l++;
				}
			}
			next = matrix;
			if (h(next) > h(current)) {
				current = next;
				frame.updateText("" + h(current));
			}
			matrix = newMatrix();
			iteration++;
		}
		return current;
	}

	private Week[][] fill(Week[][] fillMatrix) {
		for (int l = 0; l < cp.getSubjectsCount();) {
			int i = rand.nextInt(5), j = rand.nextInt(7), k = rand.nextInt(cp.getRooms().size());
			if (cond.test(matrix, cp, i, j, k, l)) {
				fillMatrix[i][j].addSubjectToScheduler(k, cp.getSubjects().get(l));
				l++;
			}
		}
		return fillMatrix;
	}

	private double h(Week[][] checkMatrix) {
		double max = 0;
		for (int i = 0; i < checkMatrix.length; i++) {
			for (int j = 0; j < checkMatrix[0].length; j++) {
				for (int k = 0; k < 7; k++) {
					if (!checkMatrix[i][j].getFreeClass(k)) {
						max += (new Double(checkMatrix[i][j].haveChoosenSubject(k)).doubleValue()
								/ new Double(checkMatrix[i][j].getClassCapacity(k)).doubleValue());
					}
				}
			}
		}
		return max;
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
}
