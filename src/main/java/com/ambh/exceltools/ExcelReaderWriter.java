package com.ambh.exceltools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.function.Function;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReaderWriter {

	protected File file;

	public ExcelReaderWriter(File file) {
		this.file = file;
	}

	private void process(Workbook wb, Sheet sheet, Function<String[], RowInfoToWrite> processor)
			throws IOException {

		Row row;
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {

			row = sheet.getRow(i);
			if (row == null) {
				break;
			}
			String[] rowAsStr = new String[row.getLastCellNum()];

			Cell cell;
			for (int j = 0; j < rowAsStr.length; j++) {
				cell = row.getCell(j);
				cell.setCellType(CellType.STRING);
				rowAsStr[j] = cell.getStringCellValue();
			}

			RowInfoToWrite toWrite = processor.apply(rowAsStr);

			for (int j = 0; j < toWrite.data.length; j++) {
				cell = row.createCell(j + toWrite.firstColumn);
				cell.setCellType(CellType.STRING);
				cell.setCellValue(toWrite.data[j]);
			}
		}

		FileOutputStream fileOut = new FileOutputStream(getFile());
		wb.write(fileOut);
		fileOut.close();

		wb.close();

	}

	public void process(String sheetName, Function<String[], RowInfoToWrite> processor)
			throws EncryptedDocumentException, IOException {
		
		Workbook wb = WorkbookFactory.create(new FileInputStream(file));
		Sheet sheet = wb.getSheet(sheetName);

		process(wb, sheet, processor);
	}

	public void process(Function<String[], RowInfoToWrite> processor) throws EncryptedDocumentException, IOException {
		
		Workbook wb = WorkbookFactory.create(new FileInputStream(file));
		Sheet sheet = wb.getSheetAt(0);

		process(wb, sheet, processor);
	}

	public File getFile() {
		return file;
	}

	public void setFile(File file) {
		this.file = file;
	}

	public static class RowInfoToWrite {
		int firstColumn;
		String[] data;

		public RowInfoToWrite(int firstColumn, String[] data) {
			this.firstColumn = firstColumn;
			this.data = data;
		}

		public String[] getData() {
			return data;
		}

		public void setData(String[] data) {
			this.data = data;
		}

		public int getFirstColumn() {
			return firstColumn;
		}

		public void setFirstColumn(int firstColumn) {
			this.firstColumn = firstColumn;
		}

	}

}
