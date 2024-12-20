package esa.sciops.utils;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.Month;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class CalendarGrid {

	int nGridRows;
	final int nGridCols = 37; // the number of columns needed to house a calendar that always has the first column corresponding to a Monday

	private CalendarCell[][] calendarGrid; // Grid of calendar cells; this does not map directly to Excel cells since rows can be used to separate years
	private Map<LocalDate, Set<String>> hols = new HashMap<>();
	
	public CalendarGrid(int startYear, Month startMonth, int nMonths, String holidayInput) {

		nGridRows = nMonths;
		calendarGrid = new CalendarCell[nGridRows][nGridCols]; // Grid of calendar cells; this does not map directly to Excel cells since rows can be used to separate years

		if (holidayInput != null) {
			readHolidays(holidayInput);
		}

		int monthValue = startMonth.getValue();
		int yearValue = startYear;

		// Process every month = row
		for (int currentRow = 0; currentRow < nMonths; currentRow++ ) {

			Month month = Month.of(monthValue);

			LocalDate firstDateOfMonth = LocalDate.of(yearValue, monthValue, 1);
			DayOfWeek dayOfWeek = firstDateOfMonth.getDayOfWeek();
			boolean leapYear = firstDateOfMonth.isLeapYear();

			// Number of days of the month, accounting for leap years
			int nDays = month.length(leapYear);

			// The calendar always has Monday as the first column, so calculate the column index from getValue()
			int firstDayOfMonthInGrid = dayOfWeek.getValue() - 1;

			// Calculate the index of the last column in that month (nDays - 1)
			int lastDayOfMonthInGrid = firstDayOfMonthInGrid + nDays - 1;

			// Set up rolling date
			LocalDate ldate = LocalDate.from(firstDateOfMonth);

			// Process every column, which may or may not be a real day in the calendar
			for (int currentCol = 0; currentCol < nGridCols; currentCol++) {

				// Check if the cell is a real calendar day
				if (currentCol >= firstDayOfMonthInGrid && currentCol <= lastDayOfMonthInGrid) {

					// Generate a Calendar cell
					DateCalendarCell cell = new DateCalendarCell(ldate);

					// If the cell is the first of the month or a Monday, we should enable printing the week number
					if (currentCol == firstDayOfMonthInGrid || ldate.getDayOfWeek() == DayOfWeek.MONDAY) {
						cell.setPrintWeekNum(true);
					}

					// If this is the first day of the month but NOT the first column in the grid set the field
					// The formatting of the first calendar column is managed in the XLS template
					// TODO Maybe the formatting of the cell should be done entirely here and not split
					if ((currentCol == firstDayOfMonthInGrid && currentCol != 0)) {
						cell.setFirstCell(true);
					}

					// If this is the last day of the month but NOT the last column in the grid set the field
					// The formatting of the last calendar column is managed in the XLS template
					// TODO Maybe the formatting of the cell should be done entirely here and not split
					if (currentCol == lastDayOfMonthInGrid && currentCol != nGridCols - 1) {
						cell.setLastCell(true);
					}

					// Determine if this is the first or last month of a (year) quarter
					int m = (month.getValue() - 1) % 3;
					if (m == 0) {
						cell.setFirstMonthOfQuarter(true);
					}
					if (m == 2) {
						cell.setLastMonthOfQuarter(true);
					}

					// Check against the holidays and set the slot aettings
					Set<String> set = hols.get(ldate);
					if (set != null) {
						for(String slot: set) {
							if (slot.equalsIgnoreCase("HOME")) {
								cell.setHomeHoliday(true);
							}
							if (slot.equalsIgnoreCase("SLOT1")) {
								cell.setSlot1Holiday(true);
							}
							if (slot.equalsIgnoreCase("SLOT2")) {
								cell.setSlot2Holiday(true);
							}
							if (slot.equalsIgnoreCase("SLOT3")) {
								cell.setSlot3Holiday(true);
							}
						}
					}

					// Assign the cell to the grid
					calendarGrid[currentRow][currentCol] = cell;

					// Increment date
					ldate = ldate.plusDays(1);

				} else {

					EmptyCalendarCell cell = new EmptyCalendarCell();

					// For non-calendar cells we just want to know if this is the first of last cells in the grid
					if (currentCol == 0) {
						cell.setFirstCell(true);
					}

					if (currentCol == nGridCols - 1) {
						cell.setLastCell(true);
					}

					calendarGrid[currentRow][currentCol] = cell;
				}

			}

			// Check if we roll into next year
			if (monthValue == Month.DECEMBER.getValue()) {
				monthValue = Month.JANUARY.getValue();
				yearValue++;
			} else {
				monthValue++;
			}
		}
	}

	private void readHolidays(String filename) {

		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");

        try (Scanner sc = new Scanner(new FileReader(filename))) {

            while (sc.hasNextLine()) {
                String text = sc.nextLine();
                String[] strs = text.split(" ", 2);

                LocalDate date = LocalDate.parse(strs[1], formatter);

                Set<String> set = hols.computeIfAbsent(date, k -> new HashSet<>());
                set.add(strs[0]);

            }

        } catch (FileNotFoundException e) {
            System.out.println("Input file not found");
            System.exit(1);
        }
 
	}

	public CalendarCell getCalendarCell(int row, int col) {
		return calendarGrid[row][col];
	}

	public int getNumRows() {
		return nGridRows;
	}

	public int getNumCols() {
		return nGridCols;
	}

}
