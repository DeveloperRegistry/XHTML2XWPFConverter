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

import org.apache.poi.xwpf.converter.xwpf.common.ElementType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * This class encapsulates a table cell parsing element.
 * 
 * @author Anton
 * 
 */
public class TableCellParsingElement extends AbstractParsingElement {

	private XWPFTableRow docxTableRow;
	private XWPFTableCell docxTableCell;

	/**
	 * Constructor
	 * 
	 * @param docxTable
	 *            table
	 * @param document
	 *            document
	 */
	public TableCellParsingElement(XWPFTableRow docxTableRow,
			XWPFDocument document) {
		super(ElementType.TABLE_CELL, false, document);
		super.setMayContainParagraph(true);
		this.docxTableRow = docxTableRow;
		this.docxTableCell = this.docxTableRow.createCell();
		//System.out.println("Created TABLE_CELL in docxTableRow="+docxTableRow);
	}

	/**
	 * @return the docxTableRow
	 */
	public XWPFTableRow getDocxTableRow() {
		return docxTableRow;
	}

	/**
	 * @param docxTableRow
	 *            the docxTableRow to set
	 */
	public void setDocxTableRow(XWPFTableRow docxTableRow) {
		this.docxTableRow = docxTableRow;
	}

	/**
	 * @return the docxTableCell
	 */
	public XWPFTableCell getDocxTableCell() {
		return docxTableCell;
	}

	/**
	 * @param docxTableCell
	 *            the docxTableCell to set
	 */
	public void setDocxTableCell(XWPFTableCell docxTableCell) {
		this.docxTableCell = docxTableCell;
	}

}
