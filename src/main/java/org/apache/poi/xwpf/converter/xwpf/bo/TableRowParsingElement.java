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
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * This class encapsulates a table row parsing element.
 * 
 * @author Anton
 * 
 */
public class TableRowParsingElement extends AbstractParsingElement {

	private XWPFTable docxTable;
	private XWPFTableRow docxTableRow;
	private TableParsingElement tableParsingElement;

	/**
	 * Constructor
	 * 
	 * @param tableParsingElement
	 *            table parsing element
	 * @param document
	 *            document
	 */
	public TableRowParsingElement(TableParsingElement tableParsingElement,
			XWPFDocument document) {
		super(ElementType.TABLE_ROW, false, document);
		this.tableParsingElement = tableParsingElement;
		this.docxTable = tableParsingElement.getDocxTable();
		System.out.println("Creating new row at position: "
				+ tableParsingElement.getDocxTable().getNumberOfRows());
		this.docxTableRow = this.docxTable
				.insertNewTableRow(tableParsingElement.getDocxTable()
						.getNumberOfRows());
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
	 * @return the tableParsingElement
	 */
	public TableParsingElement getTableParsingElement() {
		return tableParsingElement;
	}

	/**
	 * @param tableParsingElement
	 *            the tableParsingElement to set
	 */
	public void setTableParsingElement(TableParsingElement tableParsingElement) {
		this.tableParsingElement = tableParsingElement;
	}

}
