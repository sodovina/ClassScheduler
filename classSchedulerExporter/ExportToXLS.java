package classSchedulerExporter;

import classScheduler_src.Week;
import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.format.Alignment;
import jxl.format.VerticalAlignment;

public class ExportToXLS {
	private Week[][] output;
	private Strings strings = new Strings();

	public ExportToXLS(Week[][] output) {
		this.output = output;
		writeToXLS();
	}

	private void writeToXLS() {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File("Orari Mesimor.xls"));
			WritableSheet wsheet = workbook.createSheet("Shkenca Kompjuterike", 0);
			WritableSheet wsheet2 = workbook.createSheet("Matematike", 1);
			WritableSheet wsheet3 = workbook.createSheet("Matematike Financiare", 2);
			WritableFont cellFont = new WritableFont(WritableFont.ARIAL, 10);
			cellFont.setBoldStyle(WritableFont.BOLD);
			// Shkenca Kompjuterike
			for (int i = 0; i < 8; i++) {
				for (int j = 0; j < 50; j++)
					addCell(wsheet, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
			}
			for (int i = 0; i < 8; i++) {
				for (int j = 52; j < 102; j++)
					addCell(wsheet, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
			}
			for (int i = 0; i < 8; i++) {
				for (int j = 104; j < 154; j++)
					addCell(wsheet, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
			}
			addCell(wsheet, Border.ALL, BorderLineStyle.THIN, 0, 0, "Viti i pare", 1);
			for (int i = 0; i < strings.getOret().size(); i++)
				addCell(wsheet, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 0, strings.getOret().get(i), 1);
			for (int i = 0; i < strings.getDitet().size(); i++)
				addCell(wsheet, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 1 : i * 10,
						strings.getDitet().get(i), 1);

			addCell(wsheet, Border.ALL, BorderLineStyle.THIN, 0, 51, "Viti i dyte", 1);
			for (int i = 0; i < strings.getOret().size(); i++)
				addCell(wsheet, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 51, strings.getOret().get(i), 1);
			for (int i = 0; i < strings.getDitet().size(); i++)
				addCell(wsheet, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 52 : i * 10 + 52,
						strings.getDitet().get(i), 1);

			addCell(wsheet, Border.ALL, BorderLineStyle.THIN, 0, 103, "Viti i trete", 1);
			for (int i = 0; i < strings.getOret().size(); i++)
				addCell(wsheet, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 103, strings.getOret().get(i), 1);
			for (int i = 0; i < strings.getDitet().size(); i++)
				addCell(wsheet, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 104 : i * 10 + 104,
						strings.getDitet().get(i), 1);
			int row1 = 1;
			int row2 = 1;
			int row3 = 1;
			for (int d = 0; d < output.length; d++) {
				if (d == 0)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 1;
						row2 = 52;
						row3 = 104;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 1)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 10;
						row2 = 62;
						row3 = 114;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 2)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 20;
						row2 = 72;
						row3 = 124;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 3)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 30;
						row2 = 82;
						row3 = 134;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 4)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 40;
						row2 = 92;
						row3 = 144;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 1
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}

			}
			wsheet.mergeCells(0, 1, 0, 9);
			wsheet.mergeCells(0, 10, 0, 19);
			wsheet.mergeCells(0, 20, 0, 29);
			wsheet.mergeCells(0, 30, 0, 39);
			wsheet.mergeCells(0, 40, 0, 49);
			wsheet.mergeCells(0, 52, 0, 61);
			wsheet.mergeCells(0, 62, 0, 71);
			wsheet.mergeCells(0, 72, 0, 81);
			wsheet.mergeCells(0, 82, 0, 91);
			wsheet.mergeCells(0, 92, 0, 101);
			wsheet.mergeCells(0, 104, 0, 113);
			wsheet.mergeCells(0, 114, 0, 123);
			wsheet.mergeCells(0, 124, 0, 133);
			wsheet.mergeCells(0, 134, 0, 143);
			wsheet.mergeCells(0, 144, 0, 153);

			// Matematike
			for (int i = 0; i < 8; i++) {
				for (int j = 0; j < 50; j++)
					addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
			}
			for (int i = 0; i < 8; i++) {
				for (int j = 52; j < 102; j++)
					addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
			}
			for (int i = 0; i < 8; i++) {
				for (int j = 104; j < 154; j++)
					addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
			}
			addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, 0, 0, "Viti i pare", 1);
			for (int i = 0; i < strings.getOret().size(); i++)
				addCell(wsheet2, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 0, strings.getOret().get(i), 1);
			for (int i = 0; i < strings.getDitet().size(); i++)
				addCell(wsheet2, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 1 : i * 10,
						strings.getDitet().get(i), 1);

			addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, 0, 51, "Viti i dyte", 1);
			for (int i = 0; i < strings.getOret().size(); i++)
				addCell(wsheet2, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 51, strings.getOret().get(i), 1);
			for (int i = 0; i < strings.getDitet().size(); i++)
				addCell(wsheet2, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 52 : i * 10 + 52,
						strings.getDitet().get(i), 1);

			addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, 0, 103, "Viti i trete", 1);
			for (int i = 0; i < strings.getOret().size(); i++)
				addCell(wsheet2, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 103, strings.getOret().get(i), 1);
			for (int i = 0; i < strings.getDitet().size(); i++)
				addCell(wsheet2, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 104 : i * 10 + 104,
						strings.getDitet().get(i), 1);
			row1 = 1;
			row2 = 1;
			row3 = 1;
			for (int d = 0; d < output.length; d++) {
				if (d == 0)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 1;
						row2 = 52;
						row3 = 104;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 1)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 10;
						row2 = 62;
						row3 = 114;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 2)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 20;
						row2 = 72;
						row3 = 124;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 3)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 30;
						row2 = 82;
						row3 = 134;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 4)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 40;
						row2 = 92;
						row3 = 144;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 2
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet2, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}

			}
			wsheet2.mergeCells(0, 1, 0, 9);
			wsheet2.mergeCells(0, 10, 0, 19);
			wsheet2.mergeCells(0, 20, 0, 29);
			wsheet2.mergeCells(0, 30, 0, 39);
			wsheet2.mergeCells(0, 40, 0, 49);
			wsheet2.mergeCells(0, 52, 0, 61);
			wsheet2.mergeCells(0, 62, 0, 71);
			wsheet2.mergeCells(0, 72, 0, 81);
			wsheet2.mergeCells(0, 82, 0, 91);
			wsheet2.mergeCells(0, 92, 0, 101);
			wsheet2.mergeCells(0, 104, 0, 113);
			wsheet2.mergeCells(0, 114, 0, 123);
			wsheet2.mergeCells(0, 124, 0, 133);
			wsheet2.mergeCells(0, 134, 0, 143);
			wsheet2.mergeCells(0, 144, 0, 153);

			// Matematike Financiare
			for (int i = 0; i < 8; i++) {
				for (int j = 0; j < 50; j++)
					addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
			}
			for (int i = 0; i < 8; i++) {
				for (int j = 52; j < 102; j++)
					addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
			}
			for (int i = 0; i < 8; i++) {
				for (int j = 104; j < 154; j++)
					addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
			}
			addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, 0, 0, "Viti i pare", 1);
			for (int i = 0; i < strings.getOret().size(); i++)
				addCell(wsheet3, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 0, strings.getOret().get(i), 1);
			for (int i = 0; i < strings.getDitet().size(); i++)
				addCell(wsheet3, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 1 : i * 10,
						strings.getDitet().get(i), 1);

			addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, 0, 51, "Viti i dyte", 1);
			for (int i = 0; i < strings.getOret().size(); i++)
				addCell(wsheet3, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 51, strings.getOret().get(i), 1);
			for (int i = 0; i < strings.getDitet().size(); i++)
				addCell(wsheet3, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 52 : i * 10 + 52,
						strings.getDitet().get(i), 1);

			addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, 0, 103, "Viti i trete", 1);
			for (int i = 0; i < strings.getOret().size(); i++)
				addCell(wsheet3, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 103, strings.getOret().get(i), 1);
			for (int i = 0; i < strings.getDitet().size(); i++)
				addCell(wsheet3, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 104 : i * 10 + 104,
						strings.getDitet().get(i), 1);
			row1 = 1;
			row2 = 1;
			row3 = 1;
			for (int d = 0; d < output.length; d++) {
				if (d == 0)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 1;
						row2 = 52;
						row3 = 104;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 1)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 10;
						row2 = 62;
						row3 = 114;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 2)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 20;
						row2 = 72;
						row3 = 124;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 3)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 30;
						row2 = 82;
						row3 = 134;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}
				else if (d == 4)
					for (int i = 0; i < output[0].length; i++) {
						row1 = 40;
						row2 = 92;
						row3 = 144;
						int col = i + 1;
						for (int j = 0; j < 7; j++) {
							if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row1++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row2++,
										output[d][i].getClassName(j), 0);
							}
							if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == 3
									& !output[d][i].getSubjectTitle(j).equals("")) {
								// System.out.printf("%d\t%d\t%s %s %s
								// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getSubjectTitle(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getProfessor(j), 0);
								addCell(wsheet3, Border.ALL, BorderLineStyle.THIN, col, row3++,
										output[d][i].getClassName(j), 0);
							}
						}
					}

			}
			wsheet3.mergeCells(0, 1, 0, 9);
			wsheet3.mergeCells(0, 10, 0, 19);
			wsheet3.mergeCells(0, 20, 0, 29);
			wsheet3.mergeCells(0, 30, 0, 39);
			wsheet3.mergeCells(0, 40, 0, 49);
			wsheet3.mergeCells(0, 52, 0, 61);
			wsheet3.mergeCells(0, 62, 0, 71);
			wsheet3.mergeCells(0, 72, 0, 81);
			wsheet3.mergeCells(0, 82, 0, 91);
			wsheet3.mergeCells(0, 92, 0, 101);
			wsheet3.mergeCells(0, 104, 0, 113);
			wsheet3.mergeCells(0, 114, 0, 123);
			wsheet3.mergeCells(0, 124, 0, 133);
			wsheet3.mergeCells(0, 134, 0, 143);
			wsheet3.mergeCells(0, 144, 0, 153);

			workbook.write();
			workbook.close();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void addCell(WritableSheet sheet, Border border, BorderLineStyle borderLineStyle, int col, int row,
			String desc, int i) throws WriteException {
		if (i == 1) {
			WritableFont cellFont = new WritableFont(WritableFont.ARIAL, 10);
			cellFont.setBoldStyle(WritableFont.BOLD);
			WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
			cellFormat.setBorder(border, borderLineStyle);
			cellFormat.setAlignment(Alignment.CENTRE);
			cellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
			Label label = new Label(col, row, desc, cellFormat);
			sheet.addCell(label);
		} else {
			WritableCellFormat cellFormat = new WritableCellFormat();
			cellFormat.setBorder(border, borderLineStyle);
			Label label = new Label(col, row, desc, cellFormat);
			sheet.addCell(label);
		}
	}
}
