package esa.sciops.utils;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellStyle;

@Getter
@Setter
public class EmptyCalendarCell implements CalendarCell {
	
	private boolean firstCell = false;
	private boolean lastCell = false;

	public String toString() {
		return "Nothing";
	}

	/**
	 * Paint a calendar cell with the information and formatting corresponding to an cell which does not
	 * represent a calenday day.
	 * row and column in the method signature refers to the row and column of the calendar grid; the transformation
	 * to excel cells occurs within this method.
	 */
	@Override
	public void renderCell(ExcelCalendar cal, Styler styler, int r, int c) {

		// Borders, painting cell by cell...
		// 0,0
		styler.setStyle(r, c, ExcelStyle.builder().
				topBorderStyle(CellStyle.BORDER_THIN).
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 0,1
		styler.setStyle(r, c + 1, ExcelStyle.builder().
				topBorderStyle(CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 0,2
		styler.setStyle(r, c + 2, ExcelStyle.builder().
				topBorderStyle(CellStyle.BORDER_THIN).
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 1,0
		styler.setStyle(r + 1, c, ExcelStyle.builder().
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 1,1
		styler.setStyle(r + 1, c + 1, ExcelStyle.builder().
				fillColour(Styler.BLANK).
				build());

		// 1,2
		styler.setStyle(r + 1, c + 2, ExcelStyle.builder().
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 2,0
		styler.setStyle(r + 2, c, ExcelStyle.builder().
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 2,1
		styler.setStyle(r + 2, c + 1, ExcelStyle.builder().
				fillColour(Styler.BLANK).
				build());

		// 2,2
		styler.setStyle(r + 2, c + 2, ExcelStyle.builder().
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 3,0
		styler.setStyle(r + 3, c, ExcelStyle.builder().
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 3,1
		styler.setStyle(r + 3, c + 1, ExcelStyle.builder().
				fillColour(Styler.BLANK).
				build());

		// 3,2
		styler.setStyle(r + 3, c + 2, ExcelStyle.builder().
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 4,0
		styler.setStyle(r + 4, c, ExcelStyle.builder().
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 4,1
		styler.setStyle(r + 4, c + 1, ExcelStyle.builder().
				fillColour(Styler.BLANK).
				build());

		// 4,2
		styler.setStyle(r + 4, c + 2, ExcelStyle.builder().
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 5,0
		styler.setStyle(r + 5, c, ExcelStyle.builder().
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 5,1
		styler.setStyle( r + 5, c + 1, ExcelStyle.builder().
				fillColour(Styler.BLANK).
				build());

		// 5,2
		styler.setStyle(r + 5, c + 2, ExcelStyle.builder().
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 6,0
		styler.setStyle(r + 6, c, ExcelStyle.builder().
				bottomBorderStyle(CellStyle.BORDER_THIN).
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 6,1
		styler.setStyle(r + 6, c + 1, ExcelStyle.builder().
				bottomBorderStyle(CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

		// 6,2
		styler.setStyle(r + 6, c + 2, ExcelStyle.builder().
				bottomBorderStyle(CellStyle.BORDER_THIN).
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(Styler.BLANK).
				build());

	}

}
