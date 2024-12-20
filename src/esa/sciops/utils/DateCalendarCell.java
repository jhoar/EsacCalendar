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
	
	private final LocalDate date;
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

	public String toString() {
		return date.toString();
	}

	public DateCalendarCell(LocalDate ldate) {

		date = LocalDate.from(ldate);
		
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

	// Day of Month offst in 3x7 grid of spreadsheet cells
	int DoMRoffset = 0;
	int DoMCoffset = 1;

	// Day of Year
	int DoYRoffset = 1;
	int DoYCoffset = 1;

	// Week of Year
	int WoYRoffset = 5;
	int WoYCoffset = 0;
	
	/**
	 * Paint a calendar cell with the information and formatting.
	 */
	@Override
	public void renderCell(ExcelCalendar cal, Styler styler, int r, int c) {
		
		// Set the day of this month
		XSSFCell DoMcell = cal.getCell(r + DoMRoffset, c + DoMCoffset);
		DoMcell.setCellValue(date.getDayOfMonth());

		// Set the day of this year, properly formatted
		XSSFCell DoYcell = cal.getCell(r + DoYRoffset, c + DoYCoffset);
		String formatted = String.format("%03d", date.getDayOfYear());
		DoYcell.setCellValue(formatted);

		// Set the week of this year, if needed, and properly formatted
		if (isPrintWeekNum()) {
			XSSFCell WoYcell = cal.getCell(r + WoYRoffset, c + WoYCoffset);
			TemporalField woy = WeekFields.ISO.weekOfWeekBasedYear(); 
			WoYcell.setCellValue("W" + date.get(woy));
		}
				
		// Set style, painting cell by cell...

		// 0,0
		styler.setStyle(r, c, ExcelStyle.builder().
				topBorderStyle(isFirstMonthOfQuarter() ? CellStyle.BORDER_THIN : CellStyle.BORDER_THICK).
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(homeHoliday()).
				build());

		// 0,1 DoM Text Style
        styler.setStyle(r, c + 1, ExcelStyle.builder().
				topBorderStyle(isFirstMonthOfQuarter() ? CellStyle.BORDER_THIN : CellStyle.BORDER_THICK).
				fillColour(homeHoliday()).
				font(styler.getDoMFont()).
				build());

		// 0,2
		styler.setStyle(r, c + 2, ExcelStyle.builder().
				topBorderStyle(isFirstMonthOfQuarter() ? CellStyle.BORDER_THIN : CellStyle.BORDER_THICK).
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(homeHoliday()).
				build());

		// 1,0
		styler.setStyle(r + 1, c, ExcelStyle.builder().
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(homeHoliday()).
				build());

		// 1,1 DoY Text Style
		styler.setStyle(r + 1, c + 1, ExcelStyle.builder().
				fillColour(homeHoliday()).
				font(styler.getDoYFont()).
				build());

		// 1,2
		styler.setStyle(r + 1, c + 2, ExcelStyle.builder().
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(homeHoliday()).
				build());

		// 2,0
		styler.setStyle(r + 2, c, ExcelStyle.builder().
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(slot1Holiday()).
				build());

		// 2,1
		styler.setStyle(r + 2, c + 1, ExcelStyle.builder().
				fillColour(slot1Holiday()).
				build());

		// 2,2
		styler.setStyle(r + 2, c + 2, ExcelStyle.builder().
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(slot1Holiday()).
				build());

		// 3,0
		styler.setStyle(r + 3, c, ExcelStyle.builder().
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(slot2Holiday()).
				build());

		// 3,1
		styler.setStyle(r + 3, c + 1, ExcelStyle.builder().
				fillColour(slot2Holiday()).
				build());

		// 3,2
		styler.setStyle(r + 3, c + 2, ExcelStyle.builder().
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(slot2Holiday()).
				build());

		// 4,0
		styler.setStyle(r + 4, c, ExcelStyle.builder().
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(slot3Holiday()).
				build());

		// 4,1
		styler.setStyle(r + 4, c + 1, ExcelStyle.builder().
				fillColour(slot3Holiday()).
				build());

		// 4,2
		styler.setStyle(r + 4, c + 2, ExcelStyle.builder().
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(slot3Holiday()).
				build());

		// 5,0 WoY Text Style
		styler.setStyle(r + 5, c, ExcelStyle.builder().
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(homeHoliday()).
				font(styler.getWoYFont()).
				build());

		// 5,1
		styler.setStyle(r + 5, c + 1, ExcelStyle.builder().
				fillColour(homeHoliday()).
				build());

		// 5,2
		styler.setStyle(r + 5, c + 2, ExcelStyle.builder().
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(homeHoliday()).
				build());

		// 6,0
		styler.setStyle(r + 6, c, ExcelStyle.builder().
				bottomBorderStyle(isLastMonthOfQuarter() ? CellStyle.BORDER_THIN : CellStyle.BORDER_THICK).
				leftBorderStyle(isFirstCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(homeHoliday()).
				build());

		// 6,1
		styler.setStyle(r + 6, c + 1, ExcelStyle.builder().
				bottomBorderStyle(isLastMonthOfQuarter() ? CellStyle.BORDER_THIN : CellStyle.BORDER_THICK).
				fillColour(homeHoliday()).
				font(styler.getDoMFont()).
				build());

		// 6,2
		styler.setStyle(r + 6, c + 2, ExcelStyle.builder().
				bottomBorderStyle(isLastMonthOfQuarter() ? CellStyle.BORDER_THIN : CellStyle.BORDER_THICK).
				rightBorderStyle(isLastCell() ? CellStyle.BORDER_THICK : CellStyle.BORDER_THIN).
				fillColour(homeHoliday()).
				build());

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
	 * Decide which colour to paint a cell if the cell is to be used for Slot2 (STSci) holidays
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