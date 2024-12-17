package esa.sciops.utils;

import java.io.File;
import java.time.LocalDate;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;

public class EsacCal {
		
	public static void main(String args[]) throws Exception {

		EsacCal c = new EsacCal();
		c.run(args);
		
		
	}

	private void run(String[] args) throws Exception {

		int year;
		
		Option helpOpt = Option.builder("help").argName("help").hasArg(false).desc("You're reading this. Congratulations.").build();
		Option holOpt = Option.builder("holidays").argName("holidays").hasArg(true).desc("Name of Holiday file, optional, if omitted will look for holidays.txt. In the worse case it will assume no holidays. Bad luck!").build();
		Option yearOpt = Option.builder("year").argName("year").hasArg(true).desc("Year of calendar, if omitted will assume next year").build();
		Option templateOpt = Option.builder("template").argName("template").hasArg(true).desc("Name of calendar template file (XLSX). If omitted it will look for EsacCalendarTemplate.xlsx. If that is not found, computer says no.").build();
		Option outputOpt = Option.builder("output").argName("output").hasArg(true).desc("Name of output file, if omitted we'll do something useful").build();

		Options opts = new Options();
		opts.addOption(helpOpt);
		opts.addOption(holOpt);
		opts.addOption(yearOpt);
		opts.addOption(templateOpt);
		opts.addOption(outputOpt);
		
		
		CommandLine line = null;
				
		CommandLineParser parser = new DefaultParser();
		try {
			// parse the command line arguments
			line = parser.parse( opts, args );
		}
		catch( ParseException exp ) {
			// oops, something went wrong
			System.err.println( "Parsing failed.  Reason: " + exp.getMessage() );
		}

		if(line.getOptions().length == 0) {
			HelpFormatter formatter = new HelpFormatter();
			formatter.printHelp( "esacCalGen.jar", opts );
		}
		
		if(line.hasOption("help")) {
			HelpFormatter formatter = new HelpFormatter();
			formatter.printHelp( "esacCalGen.jar", opts );
			System.exit(2);
		}
		
		if(line.hasOption("year")) {
			
			String yearStr = line.getOptionValue("year");
			
			try {
				year = Integer.parseInt(yearStr);
				System.out.println("Generating calendar for " + year);
			} catch (NumberFormatException e) {
				LocalDate now = LocalDate.now();
				year = now.getYear() + 1;
				System.out.println("Did not understand the number you entered (" + yearStr + "), let's assume next year!..." + year);
			}
			
		} else { 
			LocalDate now = LocalDate.now();
			year = now.getYear() + 1;
			System.out.println("Fine, let's assume next year..." + year);
		}
				
		String holsStr = null;
		
		if(line.hasOption("holidays")) {
			
			holsStr = line.getOptionValue("holidays");
			if (!checkFile(holsStr)) {
				System.out.println("OK, let's look for holidays.txt");
				if (checkFile("holidays.txt")) {
					holsStr = "holidays.txt";
					System.out.println("Using holiday file:" + holsStr);
				} else {
					System.out.println("No holidays for you!");
				}
			} else {
				System.out.println("Using holiday file:" + holsStr);
			}
			
		} else { 
			System.out.println("OK, let's look for holidays.txt");
			if (checkFile("holidays.txt")) {
				holsStr = "holidays.txt";
				System.out.println("Using holiday file:" + holsStr);
			} else {
				System.out.println("No holidays for you!");
			}
		}
		
		String templateStr = null;
		
		if(line.hasOption("template")) {
			
			templateStr = line.getOptionValue("template");
			if (!checkFile(templateStr)) {
				System.out.println("OK, let's look for EsacCalendarTemplate.xlsx");
				if (checkFile("EsacCalendarTemplate.xlsx")) {
					templateStr = "EsacCalendarTemplate.xlsx";
					System.out.println("Using template file:" + templateStr);
				} else {
					System.out.println("Look, I can't work with nothing!");
					System.exit(1);
				}
			} else {
				System.out.println("Using template file:" + templateStr);
			}
		} else { 
			System.out.println("OK, let's look for EsacCalendarTemplate.xlsx");
			if (checkFile("EsacCalendarTemplate.xlsx")) {
				templateStr = "EsacCalendarTemplate.xlsx";
				System.out.println("Using template file:" + templateStr);
			} else {
				System.out.println("Look, I can't work with nothing!");
				System.exit(1);
			}
		}

		String outputStr = null;
		
		if(line.hasOption("output")) {
			
			outputStr = line.getOptionValue("output");
			System.out.println("Writing to file:" + outputStr);
		
		} else { 
	
			outputStr = "ESAC-Calendar-" + year + "-" + (year + 1) + ".xlsx";
			System.out.println("Silent type, eh? Well I'll write the calendar to " + outputStr);
			
		}

		CalendarGrid grid = new CalendarGrid();
		grid.generateGrid(year, holsStr);
		
		ExcelCalendar excel = new ExcelCalendar();
		excel.openTemplate(templateStr);
		excel.generateCalendar(grid, year, outputStr);

		System.out.println("\"ALL DONE\"");

		
	}

	private boolean checkFile(String fileName) {
		File c = new File(fileName);
		if(c.exists() && c.isFile()) {
			return true;
		} else {
			return false;
		}
	}

}
