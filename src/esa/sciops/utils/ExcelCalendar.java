package esa.sciops.utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDate;

import org.apache.poi.ss.usermodel.PaperSize;
import org.apache.poi.ss.usermodel.PrintOrientation;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCalendar {

	private XSSFWorkbook wb;

    private XSSFSheet sheet;
    
	public void openTemplate(String templateFilename) throws Exception {

        try (FileInputStream fis = new FileInputStream(templateFilename)) {
            wb = new XSSFWorkbook(fis);
        }
	    
	    sheet = wb.getSheetAt(0);
	}
	
	public void generateCalendar(CalendarGrid grid, int year, String outputFilename) throws Exception {
	
		int[] xlsRows = {5,12,19,27,34,41,49,56,63,71,78,85,93,100,107};
		int[] xlsRowHeight = {400,300,460,460,460,300,320};
		int rootColumn = 5;

	    // dump row heights
        for (int xlsRow : xlsRows) {
            for (int j = 0; j < xlsRowHeight.length; j++) {
                sheet.getRow(xlsRow + j - 1).setHeight((short) xlsRowHeight[j]);
            }
        }
				
		// Configure margins and print setup
		sheet.setMargin(Sheet.LeftMargin, 0.0);
		sheet.setMargin(Sheet.RightMargin, 0.0);
		sheet.setMargin(Sheet.TopMargin, 0.0);
		sheet.setMargin(Sheet.BottomMargin, 0.0);
		sheet.setMargin(Sheet.HeaderMargin, 0.0);
		sheet.setMargin(Sheet.FooterMargin, 0.0);
		sheet.setHorizontallyCenter(true);
		
		wb.setPrintArea(
	            0, //sheet index
	            0, //start column
	            118, //end column
	            0, //start row
	            120);  //end row
		
		sheet.getPrintSetup().setOrientation(PrintOrientation.LANDSCAPE);
		sheet.getPrintSetup().setPaperSize(PaperSize.A3_PAPER);
		sheet.getPrintSetup().setFitHeight((short)1);
		sheet.getPrintSetup().setFitWidth((short)1);
		
		XSSFCell cell;
		
		// Various pretties
		cell = getCell(1 - 1, 10 - 1);
		cell.setCellValue(year + "-" + (year+1));

		cell = getCell(97 - 1, 3 - 1);
		cell.setCellValue(year + 1);
		
		cell = getCell(104 - 1, 3 - 1);
		cell.setCellValue(year + 1);

		cell = getCell(111 - 1, 3 - 1);
		cell.setCellValue(year + 1);

		cell = getCell(97 - 1, 117 - 1);
		cell.setCellValue(year + 1);
		
		cell = getCell(104 - 1, 117 - 1);
		cell.setCellValue(year + 1);

		cell = getCell(111 - 1, 117 - 1);
		cell.setCellValue(year + 1);

		cell = getCell(115 - 1, 118 - 1);
		cell.setCellValue("esacCalGen/JSH " + LocalDate.now());
		
		Styler styler = new Styler(this);
		
		for (int row = 0; row < grid.getNumRows(); row++) {
			for (int col = 0; col < grid.getNumCols(); col++) {
				grid.getCalendarCell(row,col).renderCell(this, styler, xlsRows[row] - 1, (rootColumn - 1) + col * 3); // ZERO based convention!!
			}
		}
		
		FileOutputStream stream = new FileOutputStream(outputFilename);
		wb.write(stream);
		stream.close();
		wb.close();

	}

	public XSSFCell getCell(int r, int c) {
		XSSFCell cell = sheet.getRow(r).getCell(c);
		if (cell == null) {
			cell = sheet.getRow(r).createCell(c);
			cell.setCellValue("");
		}
		return cell;
	}

	public XSSFWorkbook getWorkbook() {
		return wb;
	}

	
}
