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
package org.apache.poi.xwpf.converter.xwpf.bo;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.converter.xwpf.common.ElementType;
import org.apache.poi.xwpf.converter.xwpf.common.HTMLConstants;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;

/**
 * This class encapsulates a table parsing element.
 * 
 * @author Anton
 * 
 */
public class TableParsingElement extends AbstractParsingElement {

	private XWPFTable docxTable;

	/**
	 * Constructor
	 * 
	 * @param topLevelElement
	 *            indicates if top level element or sub-element
	 * @param document
	 *            document
	 */
	public TableParsingElement(boolean topLevelElement, XWPFDocument document) {
		super(ElementType.TABLE, topLevelElement, document);
		docxTable = document.createTable();

		//Remove default rows
		for (int i = 0; i <= docxTable.getNumberOfRows(); i++) {
			docxTable.removeRow(i);
		}

	}

	/**
	 * @return the docxTable
	 */
	public XWPFTable getDocxTable() {
		return docxTable;
	}

	/**
	 * @param docxTable
	 *            the docxTable to set
	 */
	public void setDocxTable(XWPFTable docxTable) {
		this.docxTable = docxTable;
	}

	/**
	 * This method sets border on the table
	 * 
	 * @param visible
	 *            visible if true
	 * @param thick
	 *            if true, thick
	 * @param size
	 *            size
	 * @param space
	 *            space
	 */
	public void setBorder(boolean visible, boolean thick, int size, int space) {
		XWPFBorderType borderToUse = XWPFBorderType.SINGLE;
		String color = HTMLConstants.COLOR_WHITE;
		if (thick) {
			borderToUse = XWPFBorderType.THICK;
		}
		if (visible) {
			color = HTMLConstants.COLOR_BLACK;
		}
	//	System.out.println("setBorder. visible=" + visible + "; thick=" + thick
	//			+ "; size=" + size + "; color=" + color);		
		docxTable.setInsideHBorder(borderToUse, Units.toEMU(size),
				Units.toEMU(space), color);
		docxTable.setInsideVBorder(borderToUse, Units.toEMU(size),
				Units.toEMU(space), color);		
	}

	/**
	 * Sets cell margins
	 * 
	 * @param top
	 *            top margin
	 * @param left
	 *            left margin
	 * @param bottom
	 *            bottom margin
	 * @param right
	 *            right margin
	 */
	public void setCellMargins(int top, int left, int bottom, int right) {
		docxTable.setCellMargins(Units.toEMU(top), Units.toEMU(left),
				Units.toEMU(bottom), Units.toEMU(right));
	//	System.out.println("setCellMargins. top=" + top + "; left=" + left
	//			+ "; bottom" + bottom + "; right=" + right);
	}

	/**
	 * This method sets width.
	 * 
	 * @param width
	 *            table width
	 */
	public void setWidth(int width) {
		docxTable.setWidth(Units.toEMU(width));
		//System.out.println("setWidth:" + width);
	}
}
