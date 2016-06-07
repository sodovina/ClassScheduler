package classSchedulerExporter;

import classScheduler_src.Week;
import java.io.File;
import java.io.IOException;

import jxl.CellView;
import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.format.Alignment;
import jxl.format.VerticalAlignment;

public class ExportToXLS3 {
	private Week[][] output;
	private Strings strings = new Strings();
	private Colour c = Colour.BLACK;
	public ExportToXLS3(Week[][] output) {
		this.output = output;
		writeToXLS();
	}

	private void writeToXLS() {
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(strings.getFileName() + " te gjitha sipas viteve.xls"));
			WritableSheet sheet = workbook.createSheet(strings.getDepartments().get(0), 0);
			WritableFont cellFont = new WritableFont(WritableFont.ARIAL, 10);
			cellFont.setBoldStyle(WritableFont.BOLD);
			writeDept(sheet);
			workbook.write();
			workbook.close();

		} catch (IOException e) {
			System.out.println(e.getMessage());
		} catch (WriteException e) {
			System.out.println(e.getMessage());
		}
	}

	private void writeDept(WritableSheet sheet) throws WriteException {
		for (int i = 0; i < 8; i++) {
			for (int j = 0; j < 50; j++)
				addCell(sheet, Border.ALL, BorderLineStyle.THIN, i, j, "", 0,c);
		}
		for (int i = 0; i < 8; i++) {
			for (int j = 52; j < 102; j++)
				addCell(sheet, Border.ALL, BorderLineStyle.THIN, i, j, "", 0,c);
		}
		for (int i = 0; i < 8; i++) {
			for (int j = 104; j < 154; j++)
				addCell(sheet, Border.ALL, BorderLineStyle.THIN, i, j, "", 0,c);
		}
		addCell(sheet, Border.ALL, BorderLineStyle.THIN, 0, 0, strings.getYears().get(0), 1,c);
		for (int i = 0; i < strings.getClocks().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 0, strings.getClocks().get(i), 1,c);
		for (int i = 0; i < strings.getDays().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 1 : i * 10, strings.getDays().get(i), 1,c);
		addCell(sheet, Border.ALL, BorderLineStyle.THIN, 0, 51, strings.getYears().get(1), 1,c);
		for (int i = 0; i < strings.getClocks().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 51, strings.getClocks().get(i), 1,c);
		for (int i = 0; i < strings.getDays().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 52 : i * 10 + 52,
					strings.getDays().get(i), 1,c);
		addCell(sheet, Border.ALL, BorderLineStyle.THIN, 0, 103, strings.getYears().get(2), 1,c);
		for (int i = 0; i < strings.getClocks().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 103, strings.getClocks().get(i), 1,c);
		for (int i = 0; i < strings.getDays().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 104 : i * 10 + 104,
					strings.getDays().get(i), 1,c);
		int row1 = 1, row2 = 1, row3 = 1;
		for (int d = 0; d < output.length; d++) {
			if (d == 0)
				for (int i = 0; i < output[0].length; i++) {
					row1 = 1;
					row2 = 52;
					row3 = 104;
					int col = i + 1;
					for (int j = 0; j < 7; j++) {
						if (output[d][i].get("year",Integer.class,j) == 1 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("className",String.class,j),
									0,c);
						}
						if (output[d][i].get("year",Integer.class,j) == 2 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].get("className",String.class,j),
									0,c);
						}
						if (output[d][i].get("year",Integer.class,j) == 3 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].get("className",String.class,j),
									0,c);
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
						if (output[d][i].get("year",Integer.class,j) == 1 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("className",String.class,j),
									0,c);
						}
						if (output[d][i].get("year",Integer.class,j) == 2 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].get("className",String.class,j),
									0,c);
						}
						if (output[d][i].get("year",Integer.class,j) == 3 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].get("className",String.class,j),
									0,c);
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
						if (output[d][i].get("year",Integer.class,j) == 1 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("className",String.class,j),
									0,c);
						}
						if (output[d][i].get("year",Integer.class,j) == 2 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].get("className",String.class,j),
									0,c);
						}
						if (output[d][i].get("year",Integer.class,j) == 3 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].get("className",String.class,j),
									0,c);
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
						if (output[d][i].get("year",Integer.class,j) == 1 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("className",String.class,j),
									0,c);
						}
						if (output[d][i].get("year",Integer.class,j) == 2 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].get("className",String.class,j),
									0,c);
						}
						if (output[d][i].get("year",Integer.class,j) == 3 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].get("className",String.class,j),
									0,c);
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
						if (output[d][i].get("year",Integer.class,j) == 1 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("className",String.class,j),
									0,c);
						}
						if (output[d][i].get("year",Integer.class,j) == 2 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row2++, output[d][i].get("className",String.class,j),
									0,c);
						}
						if (output[d][i].get("year",Integer.class,j) == 3 & output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].get("year",Integer.class,j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++,
									output[d][i].get("subjectTitle",String.class,j), 0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].get("professor",String.class,j),
									0,c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row3++, output[d][i].get("className",String.class,j),
									0,c);
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
	private Colour getColor(int i) {
		if (i == 1)
			return Colour.BLUE;
		else if (i == 2)
			return Colour.GREEN;
		else
			return Colour.BROWN;
	}

	private void addCell(WritableSheet sheet, Border border, BorderLineStyle borderLineStyle, int col, int row,
			String desc, int i, Colour co) throws WriteException {
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
			WritableFont cellFont = new WritableFont(WritableFont.createFont("Arial"), WritableFont.DEFAULT_POINT_SIZE,
					WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE, c);
			WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
			cellFormat.setBorder(border, borderLineStyle);
			Label label = new Label(col, row, desc, cellFormat);
			sheet.addCell(label);
		}
	}
}
