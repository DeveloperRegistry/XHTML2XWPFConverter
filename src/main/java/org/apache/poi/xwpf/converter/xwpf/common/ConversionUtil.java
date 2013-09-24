/*
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 *   distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific language governing permissions and
 *   limitations under the License.
 */
package org.apache.poi.xwpf.converter.xwpf.common;

/**
 * @author Anton
 * 
 */
public class ConversionUtil {

	private static final short EXCEL_COLUMN_WIDTH_FACTOR = 256;
	private static final int UNIT_OFFSET_LENGTH = 7;
	private static final int[] UNIT_OFFSET_MAP = new int[] { 0, 36, 73, 109,
			146, 182, 219 };
	private static final double PIXEL_MULTIPLIER_WITH_OFFSET = 15.57;

	/**
	 * This method converts pixels to width units. 
	 * @param width width in pixels
	 * @return width units
	 */
	public static int convertTableCellPixelsToWidthUnits(double width) {
		int widthUnits = (int) (EXCEL_COLUMN_WIDTH_FACTOR * (width / UNIT_OFFSET_LENGTH));

		widthUnits = widthUnits + UNIT_OFFSET_MAP[((int)width % UNIT_OFFSET_LENGTH)];

		return widthUnits;
	}
	
	/**
	 * This method converts pixels to 20-th points.
	 * @param width width to converted
	 * @return conversion result
	 */
	public static double convertPixelsTo20thPoints( double width )
	{
		return (width * PIXEL_MULTIPLIER_WITH_OFFSET);
	}
}
