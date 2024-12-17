package esa.sciops.utils;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.temporal.TemporalField;
import java.time.temporal.WeekFields;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * A calendar cell containing a day
 * @author johnhoar
 *
 */
public class DateCalendarCell implements CalendarCell {
	
	private LocalDate date;
	private boolean printWeekNum = false;
	private boolean homeHoliday = false;
	private boolean slot1Holiday = false;
	private boolean slot2Holiday = false;
	private boolean slot3Holiday = false;
	private boolean weekend = false;
	private boolean firstCell = false;
	private boolean lastCell = false;
	private boolean firstMonthOfQuarter = false;
	private boolean lastMonthOfQuarter = false;

	public DateCalendarCell(LocalDate ldate) {

		date = ldate;
		
		DayOfWeek dayOfWeek = date.getDayOfWeek();
		
		if (dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY) {
			weekend = true;
		}
	}

	public boolean isPrintWeekNum() {
		return printWeekNum;
	}

	public void setPrintWeekNum(boolean printWeekNum) {
		this.printWeekNum = printWeekNum;
	}

	public boolean isHomeHoliday() {
		return homeHoliday;
	}

	public void setHomeHoliday(boolean homeHoliday) {
		this.homeHoliday = homeHoliday;
	}

	public boolean isSlot1Holiday() {
		return slot1Holiday;
	}

	public void setSlot1Holiday(boolean slot1Holiday) {
		this.slot1Holiday = slot1Holiday;
	}

	public boolean isSlot2Holiday() {
		return slot2Holiday;
	}

	public void setSlot2Holiday(boolean slot2Holiday) {
		this.slot2Holiday = slot2Holiday;
	}

	public boolean isSlot3Holiday() {
		return slot3Holiday;
	}

	public void setSlot3Holiday(boolean slot3Holiday) {
		this.slot3Holiday = slot3Holiday;
	}

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

	public boolean isWeekend() {
		return weekend;
	}
	
	public boolean isFirstMonthOfQuarter() {
		return firstMonthOfQuarter;
	}
	
	public void setFirstMonthOfQuarter(boolean b) {
		this.firstMonthOfQuarter = b;
	}

	public boolean isLastMonthOfQuarter() {
		return lastMonthOfQuarter;
	}
	
	public void setLastMonthOfQuarter(boolean b) {
		this.lastMonthOfQuarter = b;		
	}


	int DoMRoffset = 0;
	int DoMCoffset = 1;
	int DoYRoffset = 1;
	int DoYCoffset = 1;
	int WoYRoffset = 5;
	int WoYCoffset = 0;
	
	/**
	 * Paint a calendar cell with the information and formatting.
	 */
	@Override
	public void renderCell(ExcelCalendar cal, Styler styler, int r, int c) {
		
		// Content
		XSSFCell DoMcell = cal.getCell(r + DoMRoffset, c + DoMCoffset);
		DoMcell.setCellValue(date.getDayOfMonth());
		
		XSSFCell DoYcell = cal.getCell(r + DoYRoffset, c + DoYCoffset);
		String formatted = String.format("%03d", date.getDayOfYear());
		DoYcell.setCellValue(formatted);

		if (isPrintWeekNum()) {
			XSSFCell WoYcell = cal.getCell(r + WoYRoffset, c + WoYCoffset);
			TemporalField woy = WeekFields.ISO.weekOfWeekBasedYear(); 
			WoYcell.setCellValue("W" + date.get(woy));
		}
				
		short leftB;
		short rightB;
		short topB;
		short bottomB;

		// Borders, painting cell by cell...
		// 0,0
		if ( isFirstMonthOfQuarter() ) {
			topB = CellStyle.BORDER_THIN;
		} else {
			topB = CellStyle.BORDER_THICK;
		}
		
		bottomB = CellStyle.BORDER_NONE;
		
		if ( isFirstCell() ) {
			leftB = CellStyle.BORDER_THICK;
		} else {
			leftB = CellStyle.BORDER_THIN;
		}
		
		rightB = CellStyle.BORDER_NONE;
		
		styler.setStyle(r, c, topB, bottomB, leftB, rightB, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );			

		
		
		// 0,1 DoM Text Style
		if ( isFirstMonthOfQuarter() ) {
			topB = CellStyle.BORDER_THIN;
		} else {
			topB = CellStyle.BORDER_THICK;
		}
		
		bottomB = CellStyle.BORDER_NONE;
		leftB = CellStyle.BORDER_NONE;
		rightB = CellStyle.BORDER_NONE;

		styler.setStyle(r, c + 1, topB, bottomB, leftB, rightB, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER, styler.getDoMFont());


		
		// 0,2
		if ( isFirstMonthOfQuarter() ) {
			topB = CellStyle.BORDER_THIN;
		} else {
			topB = CellStyle.BORDER_THICK;
		}
		
		bottomB = CellStyle.BORDER_NONE;
		leftB = CellStyle.BORDER_NONE;
		
		if ( isLastCell() ) {
			rightB = CellStyle.BORDER_THICK;
		} else {
			rightB = CellStyle.BORDER_THIN;
		}

		styler.setStyle(r, c + 2, topB, bottomB, leftB, rightB, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
	

		
		// 1,0
		topB = CellStyle.BORDER_NONE;
		bottomB = CellStyle.BORDER_NONE;
		
		if ( isFirstCell() ) {
			leftB = CellStyle.BORDER_THICK;
		} else {
			leftB = CellStyle.BORDER_THIN;
		}
		
		rightB = CellStyle.BORDER_NONE;

		styler.setStyle(r + 1, c, topB, bottomB, leftB, rightB, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
	

		
		// 1,1 DoY Text Style
		styler.setStyle(r + 1, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER, styler.getDoYFont());
		

		
		// 1,2
		topB = CellStyle.BORDER_NONE;
		bottomB = CellStyle.BORDER_NONE;
		leftB = CellStyle.BORDER_NONE;
		
		if ( isLastCell() ) {
			rightB = CellStyle.BORDER_THICK;
		} else {
			rightB = CellStyle.BORDER_THIN;
		}

		styler.setStyle(r + 1, c + 2, topB, bottomB, leftB, rightB, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		

		
		// 2,0
		topB = CellStyle.BORDER_NONE;
		bottomB = CellStyle.BORDER_NONE;
		
		if ( isFirstCell() ) {
			leftB = CellStyle.BORDER_THICK;
		} else {
			leftB = CellStyle.BORDER_THIN;
		}
		
		rightB = CellStyle.BORDER_NONE;

		styler.setStyle(r + 2, c, topB, bottomB, leftB, rightB, slot1Holiday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );


		
		// 2,1
		styler.setStyle(r + 2, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, slot1Holiday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );


		
		// 2,2
		topB = CellStyle.BORDER_NONE;
		bottomB = CellStyle.BORDER_NONE;
		leftB = CellStyle.BORDER_NONE;
		
		if ( isLastCell() ) {
			rightB = CellStyle.BORDER_THICK;
		} else {
			rightB = CellStyle.BORDER_THIN;
		}

		styler.setStyle(r + 2, c + 2, topB, bottomB, leftB, rightB, slot1Holiday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );

	
		
		// 3,0
		topB = CellStyle.BORDER_NONE;
		bottomB = CellStyle.BORDER_NONE;
		
		if ( isFirstCell() ) {
			leftB = CellStyle.BORDER_THICK;
		} else {
			leftB = CellStyle.BORDER_THIN;
		}
		
		rightB = CellStyle.BORDER_NONE;

		styler.setStyle(r + 3, c, topB, bottomB, leftB, rightB, slot2Holiday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );

	
		
		// 3,1
		styler.setStyle(r + 3, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, slot2Holiday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );


		
		// 3,2
		topB = CellStyle.BORDER_NONE;
		bottomB = CellStyle.BORDER_NONE;
		leftB = CellStyle.BORDER_NONE;
		
		if ( isLastCell() ) {
			rightB = CellStyle.BORDER_THICK;
		} else {
			rightB = CellStyle.BORDER_THIN;
		}

		styler.setStyle(r + 3, c + 2, topB, bottomB, leftB, rightB, slot2Holiday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
		

		
		// 4,0
		topB = CellStyle.BORDER_NONE;
		bottomB = CellStyle.BORDER_NONE;
		
		if ( isFirstCell() ) {
			leftB = CellStyle.BORDER_THICK;
		} else {
			leftB = CellStyle.BORDER_THIN;
		}
		
		rightB = CellStyle.BORDER_NONE;

		styler.setStyle(r + 4, c, topB, bottomB, leftB, rightB, slot3Holiday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );


		
		// 4,1
		styler.setStyle(r + 4, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, slot3Holiday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );


		
		// 4,2
		topB = CellStyle.BORDER_NONE;
		bottomB = CellStyle.BORDER_NONE;
		leftB = CellStyle.BORDER_NONE;
		
		if ( isLastCell() ) {
			rightB = CellStyle.BORDER_THICK;
		} else {
			rightB = CellStyle.BORDER_THIN;
		}

		styler.setStyle(r + 4, c + 2, topB, bottomB, leftB, rightB, slot3Holiday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );


		
		// 5,0 WoY Text Style

		topB = CellStyle.BORDER_NONE;
		bottomB = CellStyle.BORDER_NONE;
		
		if ( isFirstCell() ) {
			leftB = CellStyle.BORDER_THICK;
		} else {
			leftB = CellStyle.BORDER_THIN;
		}
		
		rightB = CellStyle.BORDER_NONE;

		styler.setStyle(r + 5, c, topB, bottomB, leftB, rightB, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER, styler.getWoYFont() );			

	
		
		// 5,1
		styler.setStyle(r + 5, c + 1, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, CellStyle.BORDER_NONE, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );


		
		// 5,2
		topB = CellStyle.BORDER_NONE;
		bottomB = CellStyle.BORDER_NONE;
		leftB = CellStyle.BORDER_NONE;
		
		if ( isLastCell() ) {
			rightB = CellStyle.BORDER_THICK;
		} else {
			rightB = CellStyle.BORDER_THIN;
		}

		styler.setStyle(r + 5, c + 2, topB, bottomB, leftB, rightB, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );

		// 6,0
		topB = CellStyle.BORDER_NONE;

		if ( isLastMonthOfQuarter() ) {
			bottomB = CellStyle.BORDER_THIN;
		} else {
			bottomB = CellStyle.BORDER_THICK;
		}
		
		if ( isFirstCell() ) {
			leftB = CellStyle.BORDER_THICK;
		} else {
			leftB = CellStyle.BORDER_THIN;
		}
		
		rightB = CellStyle.BORDER_NONE;
		
		styler.setStyle(r + 6, c, topB, bottomB, leftB, rightB, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );


		
		// 6,1
		topB = CellStyle.BORDER_NONE;

		if ( isLastMonthOfQuarter() ) {
			bottomB = CellStyle.BORDER_THIN;
		} else {
			bottomB = CellStyle.BORDER_THICK;
		}

		leftB = CellStyle.BORDER_NONE;
		rightB = CellStyle.BORDER_NONE;

		styler.setStyle(r + 6, c + 1, topB, bottomB, leftB, rightB, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );


		
		// 6,2
		topB = CellStyle.BORDER_NONE;

		if ( isLastMonthOfQuarter() ) {
			bottomB = CellStyle.BORDER_THIN;
		} else {
			bottomB = CellStyle.BORDER_THICK;
		}

		leftB = CellStyle.BORDER_NONE;
		
		if ( isLastCell() ) {
			rightB = CellStyle.BORDER_THICK;
		} else {
			rightB = CellStyle.BORDER_THIN;
		}

		styler.setStyle(r + 6, c + 2, topB, bottomB, leftB, rightB, homeHoliday(), CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER );
	
	}
	
	/**
	 * Decide which colour to paint a cell if the cell is to be used for Slot1 (ESOC) holidays
	 * Since this is decided on a row-by-row basis, it should also manage Home (ESAC) holidays and weekends
	 * @return A colour
	 */
	private XSSFColor slot1Holiday() {
		
		XSSFColor c = Styler.WHITE;
		
		if (isSlot1Holiday()) {
			c = Styler.SLOT1;
		} else if (isHomeHoliday()) {
			c = Styler.HOME;
		} else if (isWeekend()) {
			c = Styler.WEEKEND;
		}

		return c;
		
	}

	/**
	 * Decide which colour to paint a cell if the cell is to be used for Slot1 (ESOC) holidays
	 * Since this is decided on a row-by-row basis, it should also manage Home (ESAC) holidays and weekends
	 * @return A colour
	 */
	private XSSFColor slot2Holiday() {
		
		XSSFColor c = Styler.WHITE;
		
		if (isSlot2Holiday()) { 
			c = Styler.SLOT2;
		} else if (isHomeHoliday()) {
			c = Styler.HOME;
		} else if (isWeekend()) {
			c = Styler.WEEKEND;
		}

		return c;
		
	}
	
	
	/**
	 * Decide which colour to paint a cell if the cell is to be used for Slot3 (ESTEC)  holidays
	 * Since this is decided on a row-by-row basis, it should also manage Home (ESAC)  holidays and weekends
	 * @return A colour
	 */
	private XSSFColor slot3Holiday() {

		XSSFColor c = Styler.WHITE;
		
		if (isSlot3Holiday()) {
			c = Styler.SLOT3;
		} else if (isHomeHoliday()) {
			c = Styler.HOME;
		} else if (isWeekend()) {
			c = Styler.WEEKEND;
		}

		return c;
		
	}

	/**
	 * Decide which colour to paint a cell if the cell is to be used for ESAC holidays
	 * @return A colour
	 */
	private XSSFColor homeHoliday() {
		
		XSSFColor c = Styler.WHITE;
		
		if (isHomeHoliday()) {
			c = Styler.HOME;
		} else if (isWeekend()) {
			c = Styler.WEEKEND;
		}

		return c;
		
	}
	
}