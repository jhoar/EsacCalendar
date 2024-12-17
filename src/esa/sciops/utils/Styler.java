package esa.sciops.utils;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Styler {

	private static Map<ExcelStyle, XSSFCellStyle> styleMap = new HashMap<ExcelStyle, XSSFCellStyle>();
	
	private static boolean init = false;
	
	ExcelCalendar calendar;
	
	public Styler(ExcelCalendar cal) {
		calendar = cal;
		initialise(calendar.getWorkbook());
	}
	
	/**
	 * Set the cell style for a given cell
	 * @param wb Workbook
	 * @param sheetId Sheet number
	 * @param r row 
	 * @param c column
	 * @param top Top border line type
	 * @param bottom Bottom border line type
	 * @param left Left border line type
	 * @param right Right border line type
	 * @param colour Fill colour of the cell
	 * @param halign Horizontal text alignment
	 * @param valign Vertical text alignment
	 */
	public void setStyle(int r, int c, short top, short bottom, short left, short right, XSSFColor colour, short halign, short valign) {

		XSSFCell cell = calendar.getCell(r, c);
		
		ExcelStyle style = buildStyle(top, bottom, left, right, colour, halign,	valign, null);
		
		XSSFCellStyle estyle = getOrCreateStyle(calendar.getWorkbook(), style);
	    
	    cell.setCellStyle(estyle);

	}

	/**
	 * Set the cell style for a given cell, including a font definition
	 * @param wb Workbook
	 * @param sheetId Sheet number
	 * @param r row 
	 * @param c column
	 * @param top Top border line type
	 * @param bottom Bottom border line type
	 * @param left Left border line type
	 * @param right Right border line type
	 * @param colour Fill colour of the cell
	 * @param halign Horizontal text alignment
	 * @param valign Vertical text alignment
	 */
	public void setStyle(int r, int c, short top, short bottom, short left, short right, XSSFColor colour, short halign, short valign, Font font) {
		
		XSSFCell cell = calendar.getCell(r, c);
		
		ExcelStyle style = buildStyle(top, bottom, left, right, colour, halign,	valign, font);
		
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
			
	private static ExcelStyle buildStyle(short top, short bottom, short left,
			short right, XSSFColor colour, short halign, short valign, Font font) {
		// Work around crappy limitation in Excel in number of styles in a sheet (4000)
		ExcelStyle style = new ExcelStyle();
		style.TopBorderStyle = top;
		style.TopBorderColour = IndexedColors.BLACK;
		style.BottomBorderStyle = bottom;
		style.BottomBorderColour = IndexedColors.BLACK;
		style.LeftBorderStyle = left;
		style.LeftBorderColour = IndexedColors.BLACK;
		style.RightBorderStyle = right;
		style.RightBorderColour = IndexedColors.BLACK;

		style.FillColour = colour;
		style.hAlign = halign;
		style.vAlign = valign;
		style.font = font;
		return style;
	}
	
	private static XSSFCellStyle getOrCreateStyle(XSSFWorkbook wb, ExcelStyle style) {
		
		XSSFCellStyle s = styleMap.get(style);
		if (s == null) {
			s = wb.createCellStyle();
			
			s.setBorderTop(style.TopBorderStyle);
			s.setTopBorderColor(style.TopBorderColour.getIndex());
			
			s.setBorderBottom(style.BottomBorderStyle);
			s.setBottomBorderColor(style.BottomBorderColour.getIndex());

			s.setBorderLeft(style.LeftBorderStyle);
			s.setLeftBorderColor(style.LeftBorderColour.getIndex());
			
			s.setBorderRight(style.RightBorderStyle);
			s.setRightBorderColor(style.RightBorderColour.getIndex());
			
			s.setFillBackgroundColor(style.FillColour);
			s.setFillForegroundColor(style.FillColour);
			s.setFillPattern(CellStyle.SOLID_FOREGROUND);
			
			s.setAlignment(style.hAlign);
			s.setVerticalAlignment(style.vAlign);
			s.setFont(style.font);
			
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
