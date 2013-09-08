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

import java.math.BigInteger;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.converter.xwpf.common.ElementType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

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

	/**
	 * This method sets height
	 * 
	 * @param height
	 *            height
	 * @param usePercentage
	 *            if true, use percentage instead of pixels
	 */
	public void setHeight(int height, boolean usePercentage) {
		if (usePercentage) {
			CTPageSz pageSize = this.getDocument().getDocument().getBody()
					.getSectPr().getPgSz();
			BigInteger documentHeight = pageSize.getH();
			int tableRowHeight = (documentHeight.intValue() / 100) * height;
			this.docxTableRow.setHeight(tableRowHeight);

		} else {
			this.docxTableRow.setHeight(height);
		}
	}

	/**
	 * This method sets width
	 * 
	 * @param width
	 *            width
	 * @param usePercentage
	 *            if true, use percentage instead of pixels
	 */
	public void setWidth(int width, boolean usePercentage) {

		CTTc cttc = this.docxTableCell.getCTTc();
		CTTcPr pr = cttc.getTcPr();
		if (pr == null) {
			pr = cttc.addNewTcPr();
		}

		CTTblWidth ctbl = pr.getTcW();
		if (ctbl == null) {
			ctbl = pr.addNewTcW();
		}

		if (usePercentage) {
			CTPageSz pageSize = this.getDocument().getDocument().getBody()
					.getSectPr().getPgSz();
			BigInteger documentWidth = pageSize.getW();
			int tableWidth = (documentWidth.intValue() / 100) * width;
			ctbl.setW(BigInteger.valueOf(tableWidth));

		} else {
			ctbl.setW(BigInteger.valueOf(Units.toEMU(width)));
		}
	}

}
