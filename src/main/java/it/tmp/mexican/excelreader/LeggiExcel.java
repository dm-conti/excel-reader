/*
 * @(#)LeggiExcel.java        created on 14/ott/2014
 *
 * Copyright (c) 2007-2014 QuiGroup,
 *
 * This software is the confidential and proprietary information of QuiGroup 
 * Networks srl, Inc. ("Confidential Information").  You shall not
 * disclose such Confidential Information and shall use it only in
 * accordance with the terms of the license agreement you entered into
 * with QuiGroup.
 */
package it.tmp.mexican.excelreader;

import java.io.File;
import java.io.IOException;
import java.util.Date;

import jxl.Cell;
import jxl.DateCell;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * Class <code>LeggiExcel.java</code>
 *
 * @author Domenico Conti [domenico.conti@quigroup.it]
 *
 */
public class LeggiExcel {

	/**
	 * @param args
	 * @throws WriteException
	 */
	public static void main(String[] args) throws WriteException {
		Cell cell;

		String txtValue = "";
		Date dateValue = null;
		int intValue = 0;
		try {
			// Apro il file di excel da leggere
			Workbook workbookToRead = Workbook
					.getWorkbook(new File("data.xls"));

			// Apro il file su cui scrivere
			WritableWorkbook workbookToWrite = Workbook
					.createWorkbook(new File("new_data.xls"));
			workbookToWrite.createSheet("Report", 0);

			// Seleziono il foglio sul quale voglio operare (il primo foglio ha
			// indice 0)
			Sheet sheet = workbookToRead.getSheet(0);

			// Creo il foglio sul quale voglio operare (il primo foglio ha
			// indice 0)
			WritableSheet sheetToWrite = workbookToWrite.getSheet("Report");

			// Leggo tutte le righe
			int numeroRighe = sheet.getRows();// calcolo quante righe ci sono
												// nel foglio
			int numeroColonne = sheet.getColumns();

			// indice riga, parto da 1 per saltare l'intestazione dei campi

			// Attento che c'è una riga in più per l'intestazione!
			for (int row = 1; row < numeroRighe; row++) {
				System.out.print("Riga: " + row + " :: ");
				
				for (int z = 0; z < 4; z++) {
					
					int actualRows = sheetToWrite.getRows();
					System.out.println("Actual Rows = " + actualRows);
					
					for (int column = 0; column < numeroColonne; column++) {
						cell = sheet.getCell(column, row);

						if (cell instanceof LabelCell) {
							txtValue = ((LabelCell) cell).getString();
							System.out.print(" | " + txtValue + " | ");

							Label tmpCell = new Label(column, actualRows,txtValue);
							sheetToWrite.addCell(tmpCell);
						}

						if (cell instanceof DateCell) {
							dateValue = ((DateCell) cell).getDate();
							System.out.print(" | " + dateValue + " | ");

							DateTime tmpCell = new DateTime(column, actualRows,dateValue);
							sheetToWrite.addCell(tmpCell);
						}

						if (cell instanceof NumberCell) {
							intValue = (int) ((NumberCell) cell).getValue();
							System.out.print(" | " + intValue + " | ");

							Number tmpCell = new Number(column, actualRows,intValue);
							sheetToWrite.addCell(tmpCell);
						}
					}
				}
				System.out.println("");
			}

			// Chiudo excel e libero la memoria
			workbookToRead.close();

			// All sheets and cells added. Now write out the workbook
			workbookToWrite.write();
			workbookToWrite.close();

		} catch (BiffException | IOException e) {
			e.printStackTrace();
		} finally {

		}
	}
}