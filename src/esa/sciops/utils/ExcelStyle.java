package esa.sciops.utils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFColor;


/**
 * Class to capture cell styling information for use in a HashMap-based cache
 * @author johnhoar
 *
 */
public class ExcelStyle {

	public short TopBorderStyle = CellStyle.BORDER_NONE;
	public short BottomBorderStyle = CellStyle.BORDER_NONE;
	public short LeftBorderStyle = CellStyle.BORDER_NONE;
	public short RightBorderStyle = CellStyle.BORDER_NONE;

	public IndexedColors TopBorderColour = IndexedColors.BLACK;
	public IndexedColors BottomBorderColour = IndexedColors.BLACK;
	public IndexedColors LeftBorderColour = IndexedColors.BLACK;
	public IndexedColors RightBorderColour = IndexedColors.BLACK;

	public XSSFColor FillColour = Styler.WHITE;

	public short hAlign = CellStyle.ALIGN_CENTER;
	public short vAlign = CellStyle.VERTICAL_CENTER;
	public Font font;

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime
				* result
				+ ((BottomBorderColour == null) ? 0 : BottomBorderColour
						.hashCode());
		result = prime * result + BottomBorderStyle;
		result = prime * result
				+ ((FillColour == null) ? 0 : FillColour.hashCode());
		result = prime
				* result
				+ ((LeftBorderColour == null) ? 0 : LeftBorderColour.hashCode());
		result = prime * result + LeftBorderStyle;
		result = prime
				* result
				+ ((RightBorderColour == null) ? 0 : RightBorderColour
						.hashCode());
		result = prime * result + RightBorderStyle;
		result = prime * result
				+ ((TopBorderColour == null) ? 0 : TopBorderColour.hashCode());
		result = prime * result + TopBorderStyle;
		result = prime * result + ((font == null) ? 0 : font.hashCode());
		result = prime * result + hAlign;
		result = prime * result + vAlign;
		return result;
	}
	
	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		ExcelStyle other = (ExcelStyle) obj;
		if (BottomBorderColour != other.BottomBorderColour)
			return false;
		if (BottomBorderStyle != other.BottomBorderStyle)
			return false;
		if (FillColour == null) {
			if (other.FillColour != null)
				return false;
		} else if (!FillColour.equals(other.FillColour))
			return false;
		if (LeftBorderColour != other.LeftBorderColour)
			return false;
		if (LeftBorderStyle != other.LeftBorderStyle)
			return false;
		if (RightBorderColour != other.RightBorderColour)
			return false;
		if (RightBorderStyle != other.RightBorderStyle)
			return false;
		if (TopBorderColour != other.TopBorderColour)
			return false;
		if (TopBorderStyle != other.TopBorderStyle)
			return false;
		if (font == null) {
			if (other.font != null)
				return false;
		} else if (!font.equals(other.font))
			return false;
		if (hAlign != other.hAlign)
			return false;
		if (vAlign != other.vAlign)
			return false;
		return true;
	}
	
	
}
