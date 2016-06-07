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

public class ExportToXLS2 {
	private Week[][] output;
	private Strings strings = new Strings();

	public ExportToXLS2(Week[][] output) {
		this.output = output;
		writeToXLS();
	}

	private void writeToXLS() {
		try {
			WritableWorkbook workbook = Workbook
					.createWorkbook(new File(strings.getFileName() + " te gjitha bashke.xls"));
			WritableSheet sheet = workbook.createSheet(strings.getFileName(), 0);
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
		Colour c = Colour.BLACK;
		for (int i = 0; i < 8; i++) {
			for (int j = 0; j < 105; j++)
				addCell(sheet, Border.ALL, BorderLineStyle.THIN, i, j, "", 0, c);
		}
		addCell(sheet, Border.ALL, BorderLineStyle.THIN, 0, 0, "Te githa vitet", 1, c);
		for (int i = 0; i < strings.getClocks().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, i + 1, 0, strings.getClocks().get(i), 1, c);
		for (int i = 0; i < strings.getDays().size(); i++)
			addCell(sheet, Border.ALL, BorderLineStyle.MEDIUM, 0, i == 0 ? i + 1 : i * 21, strings.getDays().get(i), 1,
					c);
		int row1 = 1;
		for (int d = 0; d < output.length; d++) {
			if (d == 0)
				for (int i = 0; i < output[0].length; i++) {
					row1 = 1;
					int col = i + 1;
					for (int j = 0; j < 7; j++) {
						if (output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].get("subjectTitle",String.class,j), 0, c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("professor",String.class,j),
									0, c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("className",String.class,j),
									0, c);
						}
					}
				}
			else if (d == 1)
				for (int i = 0; i < output[0].length; i++) {
					row1 = 21;
					int col = i + 1;
					for (int j = 0; j < 7; j++) {
						if (output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].get("subjectTitle",String.class,j), 0, c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("professor",String.class,j),
									0, c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("className",String.class,j),
									0, c);
						}
					}
				}
			else if (d == 2)
				for (int i = 0; i < output[0].length; i++) {
					row1 = 42;
					int col = i + 1;
					for (int j = 0; j < 7; j++) {
						if (output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].get("subjectTitle",String.class,j), 0, c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("professor",String.class,j),
									0, c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("className",String.class,j),
									0, c);
						}
					}
				}
			else if (d == 3)
				for (int i = 0; i < output[0].length; i++) {
					row1 = 63;
					int col = i + 1;
					for (int j = 0; j < 7; j++) {
						if (output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].get("subjectTitle",String.class,j), 0, c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("professor",String.class,j),
									0, c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("className",String.class,j),
									0, c);
						}
					}
				}
			else if (d == 4)
				for (int i = 0; i < output[0].length; i++) {
					row1 = 84;
					int col = i + 1;
					for (int j = 0; j < 7; j++) {
						if (output[d][i].get("subjectTitle",String.class,j) != null) {
							// System.out.printf("%d\t%d\t%s %s %s
							// %n",output[d][i].getYear(j),j,output[d][i].get("subjectTitle",String.class,j),output[d][i].get("professor",String.class,j),output[d][i].get("className",String.class,j));
							c = getColor(output[d][i].get("department",Integer.class,j));
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++,
									output[d][i].get("subjectTitle",String.class,j), 0, c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("professor",String.class,j),
									0, c);
							addCell(sheet, Border.ALL, BorderLineStyle.THIN, col, row1++, output[d][i].get("className",String.class,j),
									0, c);
						}
					}
				}
		}
		sheet.mergeCells(0, 1, 0, 20);
		sheet.mergeCells(0, 21, 0, 41);
		sheet.mergeCells(0, 42, 0, 62);
		sheet.mergeCells(0, 63, 0, 83);
		sheet.mergeCells(0, 84, 0, 104);
		CellView cv = new CellView();
		cv.setAutosize(true);
		for (int x = 0; x < 8; x++) {
			sheet.setColumnView(x, cv);
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
			String desc, int i, Colour c) throws WriteException {
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
