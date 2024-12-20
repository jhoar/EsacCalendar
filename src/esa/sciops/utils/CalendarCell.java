package esa.sciops.utils;


/**
 * A calendar is a grid of cells; with a style to apply to a cell given a row and column to start from
 * The cell itself is a 3x7 grid of cells in the spreadsheet
 * 
 * @author johnhoar
 *
 */
interface CalendarCell {

	void renderCell(ExcelCalendar cal, Styler styler, int r, int c);

}
