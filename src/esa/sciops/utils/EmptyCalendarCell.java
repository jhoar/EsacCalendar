package esa.sciops.utils;
import org.apache.poi.ss.usermodel.CellStyle;

public class EmptyCalendarCell implements CalendarCell {
	
	private boolean firstCell = false;
	private boolean lastCell = false;


	public boolean isFirstCell() {
		return firstCell;
	}


	public void setFirstCell(boolean firstCell) {
		this.firstCell = firstCell;
	}


	public boolean isLastCell() {
		return lastCell;
	}


	public void setLastCell(boolean lastCell) {
		this.lastCell = lastCell;
	}


	/**
	 * Paint a calendar cell with the information and formatting.
	 */
	@Override
	public void renderCell(ExcelCalendar cal, Styler styler, int r, int c) {
		
		// Borders, painting cell by cell...
		// 0,0
		if ( isFirstCell() ) {
			styler.setStyle(r, c, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r, c, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );			
		}

		// 0,1
		styler.setStyle(r, c + 1, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );

		// 0,2
		if ( isLastCell() ) {
			styler.setStyle(r, c + 2, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r, c + 2, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );		
		}
		
		// 1,0
		if ( isFirstCell() ) {
			styler.setStyle(r + 1, c, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 1, c, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );			
		}

		// 1,1
		styler.setStyle(r + 1, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );

		// 1,2
		if ( isLastCell() ) {
			styler.setStyle(r + 1, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 1, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );		
		}
		
		// 2,0
		if ( isFirstCell() ) {
			styler.setStyle(r + 2, c, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 2, c, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );			
		}

		// 2,1
		styler.setStyle(r + 2, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );

		// 2,2
		if ( isLastCell() ) {
			styler.setStyle(r + 2, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 2, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );		
		}

		// 3,0
		if ( isFirstCell() ) {
			styler.setStyle(r + 3, c, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 3, c, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );			
		}

		// 3,1
		styler.setStyle(r + 3, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );

		// 3,2
		if ( isLastCell() ) {
			styler.setStyle(r + 3, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 3, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );		
		}
		
		// 4,0
		if ( isFirstCell() ) {
			styler.setStyle(r + 4, c, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 4, c, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );			
		}

		// 4,1
		styler.setStyle(r + 4, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );

		// 4,2
		if ( isLastCell() ) {
			styler.setStyle(r + 4, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 4, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );		
		}

		// 5,0
		if ( isFirstCell() ) {
			styler.setStyle(r + 5, c, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 5, c, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );			
		}

		// 5,1
		styler.setStyle(r + 5, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );

		// 5,2
		if ( isLastCell() ) {
			styler.setStyle(r + 5, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 5, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );		
		}

		// 6,0
		if ( isFirstCell() ) {
			styler.setStyle(r + 6, c, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 6, c, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );			
		}

		// 6,1
		styler.setStyle(r + 6, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );

		// 6,2
		if ( isLastCell() ) {
			styler.setStyle(r + 6, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		} else {
			styler.setStyle(r + 6, c + 2, CellStyle.BORDER_NONE, CellStyle.BORDER_THIN, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, Styler.BLANK, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );		
		}
	
	}

}
