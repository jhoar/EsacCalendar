package esa.sciops.utils;


/**
 * A calendar is a grid of cells; each cell will know how to
 * draw itself in the XLS sheet, given a row and column to start from
 * 
 * The cell itself is a 3x7 grid of cells in the spreadsheet
 * 
 * @author johnhoar
 *
 */
interface CalendarCell {

	/**
	 * 
	 * Paint a 
	 * 
	 * @param wb XLSX workbook
	 * @param sheet Sheet number
	 * @param r Row number (0-based)
	 * @param c Column number (0-based)
	 */
	void renderCell(ExcelCalendar cal, Styler styler, int r, int c);

}
