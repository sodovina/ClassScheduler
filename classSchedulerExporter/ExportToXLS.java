package classSchedulerExporter;

import classScheduler_src.Week;
import java.io.File;
import java.io.IOException;

import jxl.CellView;
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
			WritableWorkbook workbook = Workbook.createWorkbook(new File(strings.getFileName() + ".xls"));
			WritableSheet shk_sheet = workbook.createSheet(strings.getDepartments().get(0), 0);
			WritableSheet m_sheet = workbook.createSheet(strings.getDepartments().get(1), 1);
			WritableSheet mf_sheet = workbook.createSheet(strings.getDepartments().get(2), 2);
			WritableFont cellFont = new WritableFont(WritableFont.ARIAL, 10);
			cellFont.setBoldStyle(WritableFont.BOLD);
			writeDept(shk_sheet, 1);
			writeDept(m_sheet, 2);
			writeDept(mf_sheet, 3);
			workbook.write();
			workbook.close();

		} catch (IOException e) {
			System.out.println(e.getMessage());
		} catch (WriteException e) {
			System.out.println(e.getMessage());
		}
	}

	private void writeDept(WritableSheet sheet, int dept) throws WriteException {
		for (int i = 0; i < 8; i++) {
			for (int j = 0; j < 50; j++)
				addCell(sheet, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
		}
		for (int i = 0; i < 8; i++) {
			for (int j = 52; j < 102; j++)
				addCell(sheet, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
		}
		for (int i = 0; i < 8; i++) {
			for (int j = 104; j < 154; j++)
				addCell(sheet, Border.ALL, BorderLineStyle.THIN, i, j, "", 0);
		}
		addCell(sheet, Border.ALL, BorderLineStyle.THIN, 0, 0, strings.getYears().get(0), 1);
		for (int i = 0; i < strings.getClocks().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 0, strings.getClocks().get(i), 1);
		for (int i = 0; i < strings.getDays().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 1 : i * 10, strings.getDays().get(i), 1);
		addCell(sheet, Border.ALL, BorderLineStyle.THIN, 0, 51, strings.getYears().get(1), 1);
		for (int i = 0; i < strings.getClocks().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 51, strings.getClocks().get(i), 1);
		for (int i = 0; i < strings.getDays().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 52 : i * 10 + 52,
					strings.getDays().get(i), 1);
		addCell(sheet, Border.ALL, BorderLineStyle.THIN, 0, 103, strings.getYears().get(2), 1);
		for (int i = 0; i < strings.getClocks().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 103, strings.getClocks().get(i), 1);
		for (int i = 0; i < strings.getDays().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 104 : i * 10 + 104,
					strings.getDays().get(i), 1);
		int row1 = 1, row2 = 1, row3 = 1;
		for (int d = 0; d < output.length; d++) {
			if (d == 0)
				for (int i = 0; i < output[0].length; i++) {
					row1 = 1;
					row2 = 52;
					row3 = 104;
					int col = i + 1;
					for (int j = 0; j < 7; j++) {
						if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].getClassName(j),
									0);
						}
						if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].getClassName(j),
									0);
						}
						if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].getClassName(j),
									0);
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
						if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].getClassName(j),
									0);
						}
						if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].getClassName(j),
									0);
						}
						if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].getClassName(j),
									0);
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
						if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].getClassName(j),
									0);
						}
						if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].getClassName(j),
									0);
						}
						if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].getClassName(j),
									0);
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
						if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].getClassName(j),
									0);
						}
						if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].getClassName(j),
									0);
						}
						if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].getClassName(j),
									0);
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
						if (output[d][i].getYear(j) == 1 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].getClassName(j),
									0);
						}
						if (output[d][i].getYear(j) == 2 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].getClassName(j),
									0);
						}
						if (output[d][i].getYear(j) == 3 & output[d][i].getDepartment(j) == dept
								& output[d][i].getSubjectTitle(j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].getSubjectTitle(j),output[d][i].getProfessor(j),output[d][i].getClassName(j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
									output[d][i].getSubjectTitle(j), 0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].getProfessor(j),
									0);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].getClassName(j),
									0);
						}
					}
				}
		}
		sheet.mergeCells(0, 1, 0, 9);
		sheet.mergeCells(0, 10, 0, 19);
		sheet.mergeCells(0, 20, 0, 29);
		sheet.mergeCells(0, 30, 0, 39);
		sheet.mergeCells(0, 40, 0, 49);
		sheet.mergeCells(0, 52, 0, 61);
		sheet.mergeCells(0, 62, 0, 71);
		sheet.mergeCells(0, 72, 0, 81);
		sheet.mergeCells(0, 82, 0, 91);
		sheet.mergeCells(0, 92, 0, 101);
		sheet.mergeCells(0, 104, 0, 113);
		sheet.mergeCells(0, 114, 0, 123);
		sheet.mergeCells(0, 124, 0, 133);
		sheet.mergeCells(0, 134, 0, 143);
		sheet.mergeCells(0, 144, 0, 153);
		CellView c = new CellView();
		c.setAutosize(true);
		for (int x = 0; x < 8; x++) {
			sheet.setColumnView(x, c);
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
