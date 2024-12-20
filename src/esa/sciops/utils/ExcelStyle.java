package esa.sciops.utils;
import lombok.Builder;
import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFColor;



/**
 * Class to capture cell styling information for use in a HashMap-based cache
 * @author johnhoar
 *
 */
@Builder
@Getter
public class ExcelStyle {

	@Builder.Default private final short topBorderStyle = CellStyle.BORDER_NONE;
	@Builder.Default private final short bottomBorderStyle = CellStyle.BORDER_NONE;
	@Builder.Default private final short leftBorderStyle = CellStyle.BORDER_NONE;
	@Builder.Default private final short rightBorderStyle = CellStyle.BORDER_NONE;

	@Builder.Default private final IndexedColors topBorderColour = IndexedColors.BLACK;
	@Builder.Default private final IndexedColors bottomBorderColour = IndexedColors.BLACK;
	@Builder.Default private final IndexedColors leftBorderColour = IndexedColors.BLACK;
	@Builder.Default private final IndexedColors rightBorderColour = IndexedColors.BLACK;

	@Builder.Default private final XSSFColor fillColour = Styler.WHITE;

	@Builder.Default private final short hAlign = CellStyle.ALIGN_CENTER;
	@Builder.Default private final short vAlign = CellStyle.VERTICAL_CENTER;

	private final Font font;

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime
				* result
				+ ((bottomBorderColour == null) ? 0 : bottomBorderColour
						.hashCode());
		result = prime * result + bottomBorderStyle;
		result = prime * result
				+ ((fillColour == null) ? 0 : fillColour.hashCode());
		result = prime
				* result
				+ ((leftBorderColour == null) ? 0 : leftBorderColour.hashCode());
		result = prime * result + leftBorderStyle;
		result = prime
				* result
				+ ((rightBorderColour == null) ? 0 : rightBorderColour
						.hashCode());
		result = prime * result + rightBorderStyle;
		result = prime * result
				+ ((topBorderColour == null) ? 0 : topBorderColour.hashCode());
		result = prime * result + topBorderStyle;
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
		if (bottomBorderColour != other.bottomBorderColour)
			return false;
		if (bottomBorderStyle != other.bottomBorderStyle)
			return false;

		if (fillColour == null) {
			if (other.fillColour != null)
				return false;
		} else if (!fillColour.equals(other.fillColour))
			return false;

		if (leftBorderColour != other.leftBorderColour)
			return false;

		if (leftBorderStyle != other.leftBorderStyle)
			return false;

		if (rightBorderColour != other.rightBorderColour)
			return false;

		if (rightBorderStyle != other.rightBorderStyle)
			return false;

		if (topBorderColour != other.topBorderColour)
			return false;

		if (topBorderStyle != other.topBorderStyle)
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
