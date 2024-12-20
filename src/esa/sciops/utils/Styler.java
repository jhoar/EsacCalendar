package esa.sciops.utils;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Encapsulates all style information to apply to a cell in an Excel spreadsheet
 */
public class Styler {

	private static final Map<ExcelStyle, XSSFCellStyle> styleMap = new HashMap<>();
	
	private static boolean init = false;
	
	ExcelCalendar calendar;
	
	public Styler(ExcelCalendar cal) {
		calendar = cal;
		initialise(calendar.getWorkbook());
	}

	public void setStyle(int r, int c, ExcelStyle style) {
		XSSFCell cell = calendar.getCell(r, c);
		XSSFCellStyle estyle = getOrCreateStyle(calendar.getWorkbook(), style);
		cell.setCellStyle(estyle);
	}

	public Font getDoMFont() {
		return DOMFONT;
	}

	public Font getDoYFont() {
		return DOYFONT;
	}
	
	public Font getWoYFont() {
		return WOYFONT;
	}

	private void initialise(XSSFWorkbook wb) {
		
		if(init) {
			return;
		}
		
		// One-off font setup
		DOMFONT = wb.createFont();
		DOMFONT.setFontHeightInPoints((short)18);
		DOMFONT.setFontName("ReykjavikOne OT CGauge");
		
		DOYFONT = wb.createFont();
		DOYFONT.setFontHeightInPoints((short)13);
		DOYFONT.setFontName("ReykjavikOne OT AGauge");

		WOYFONT = wb.createFont();
		WOYFONT.setFontHeightInPoints((short)10);
		WOYFONT.setFontName("ReykjavikOne OT AGauge");		
		
		init = true;

	}
	
	private static XSSFCellStyle getOrCreateStyle(XSSFWorkbook wb, ExcelStyle style) {
		
		XSSFCellStyle s = styleMap.get(style);
		if (s == null) {
			s = wb.createCellStyle();
			
			s.setBorderTop(style.getTopBorderStyle());
			s.setTopBorderColor(style.getTopBorderColour().getIndex());
			
			s.setBorderBottom(style.getBottomBorderStyle());
			s.setBottomBorderColor(style.getBottomBorderColour().getIndex());

			s.setBorderLeft(style.getLeftBorderStyle());
			s.setLeftBorderColor(style.getLeftBorderColour().getIndex());
			
			s.setBorderRight(style.getRightBorderStyle());
			s.setRightBorderColor(style.getRightBorderColour().getIndex());
			
			s.setFillBackgroundColor(style.getFillColour());
			s.setFillForegroundColor(style.getFillColour());
			s.setFillPattern(CellStyle.SOLID_FOREGROUND);
			
			s.setAlignment(style.getHAlign());
			s.setVerticalAlignment(style.getVAlign());
			s.setFont(style.getFont());
			
			styleMap.put(style, s);

		}
		
		return s;
	}

	public static final XSSFColor WHITE    = new XSSFColor(new java.awt.Color(255, 255, 255));
	public static final XSSFColor WEEKEND  = new XSSFColor(new java.awt.Color(208, 208, 208));
	public static final XSSFColor HOME     = new XSSFColor(new java.awt.Color(173, 227, 255));
	public static final XSSFColor SLOT1     = new XSSFColor(new java.awt.Color(247, 175, 191));
	public static final XSSFColor SLOT2     = new XSSFColor(new java.awt.Color(184, 208, 138));
	public static final XSSFColor SLOT3    = new XSSFColor(new java.awt.Color(253, 220, 126));
	public static final XSSFColor BLANK    = new XSSFColor(new java.awt.Color(109, 109, 109));
	
	private Font DOMFONT;
	private Font DOYFONT;
	private Font WOYFONT;

}
