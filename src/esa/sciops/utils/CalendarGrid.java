package esa.sciops.utils;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.Month;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;

public class CalendarGrid {

	int nGridRows = 15; // 1.25 years
	int nGridCols = 37; // the number of columns needed to house a calendar that always has the first column corresponding to a DayOfWeek

	CalendarCell calendarGrid[][] = new CalendarCell[nGridRows][nGridCols];
	
	Map<LocalDate, Set<String>> hols = new HashMap<LocalDate, Set<String>>();
	
	public void generateGrid(int startYear, String holidayInput) {
		
		int currentRow = 0;
		boolean leapYear;
		Month [] quarter = new Month[]{Month.JANUARY, Month.FEBRUARY, Month.MARCH};
		
		if (holidayInput != null) {
			readHolidays(holidayInput);
		}
		
		leapYear = LocalDate.of(startYear, Month.JANUARY, 1).isLeapYear();
		currentRow = generateRows(startYear, Month.values(), currentRow, leapYear);
		
		startYear++;
		
		leapYear = LocalDate.of(startYear, Month.JANUARY, 1).isLeapYear();
		currentRow = generateRows(startYear, quarter, currentRow, leapYear);
		
	}

	private int generateRows(int year, Month[] months, int currentRow, boolean leapYear) {

		for(Month month: months) {

			LocalDate date = LocalDate.of(year, month, 1);
			
			DayOfWeek dayOfWeek = date.getDayOfWeek();
			int nDays = month.length(leapYear);
			
			int firstDayOfMonthInGrid = dayOfWeek.ordinal();
			int lastDayOfMonthInGrid = firstDayOfMonthInGrid + nDays - 1;
						
			for (int i = 0; i < nGridCols; i++) {
				
				// Is a calendar day
				if (i >= firstDayOfMonthInGrid && i <= lastDayOfMonthInGrid) {
					
					LocalDate ldate = LocalDate.of(year, month, i - firstDayOfMonthInGrid + 1);
					
					DateCalendarCell cell = new DateCalendarCell(ldate);
					calendarGrid[currentRow][i] = cell;

					if (i == firstDayOfMonthInGrid || ldate.getDayOfWeek() == DayOfWeek.MONDAY) {
						cell.setPrintWeekNum(true);
					}

					if ((i == firstDayOfMonthInGrid && i != 0)) {
						cell.setFirstCell(true);
					}

					if (i == lastDayOfMonthInGrid && i != nGridCols - 1) {
						cell.setLastCell(true);
					}
					
					int m = (month.getValue() - 1) % 3;
					if (m == 0) {
						cell.setFirstMonthOfQuarter(true);
					}
					if (m == 2) {
						cell.setLastMonthOfQuarter(true);
					}
					
					Set<String> set = hols.get(ldate);
					if (set != null) {
						for(String site: set) {
							if (site.equalsIgnoreCase("HOME")) {
								cell.setHomeHoliday(true);
							}
							if (site.equalsIgnoreCase("SLOT1")) {
								cell.setSlot1Holiday(true);
							}
							if (site.equalsIgnoreCase("SLOT2")) {
								cell.setSlot2Holiday(true);
							}
							if (site.equalsIgnoreCase("SLOT3")) {
								cell.setSlot3Holiday(true);
							}
						}
					}

					// Set Euclid specific items
					// SOPS

					//

				} else {
					
					EmptyCalendarCell cell = new EmptyCalendarCell();
					
					if (i == 0) {
						cell.setFirstCell(true);
					}

					if (i == nGridCols - 1) {
						cell.setLastCell(true);
					}

					calendarGrid[currentRow][i] = cell;
				}

			}
				
			// Finally move to next row;
			currentRow++;
		}
		
		return currentRow;
		
	}

	private void readHolidays(String filename) {

		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy");
		
		Scanner sc = null;

        try {
            sc = new Scanner(new FileReader(filename));

            while (sc.hasNextLine()) {
                String text = sc.nextLine();
                String[] strs = text.split(" ", 2);

                LocalDate date = LocalDate.parse(strs[1], formatter);
                
                Set<String> set = hols.get(date);
                if (set == null) {
                	set = new HashSet<String>();
                	hols.put(date, set);
                }
                set.add(strs[0]);
              
            }

        } catch (FileNotFoundException e) {
            System.out.println("Input file not found");
            System.exit(1);
        } finally {
            sc.close();
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
