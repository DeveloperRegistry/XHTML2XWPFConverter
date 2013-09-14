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

import org.apache.poi.xwpf.converter.xwpf.common.ConversionUtil;
import org.apache.poi.xwpf.converter.xwpf.common.ElementType;
import org.apache.poi.xwpf.converter.xwpf.common.HTMLConstants;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

/**
 * This class encapsulates a table cell parsing element.
 * 
 * @author Anton
 * 
 */
public class TableCellParsingElement extends AbstractParsingElement {

	private XWPFTableRow docxTableRow;
	private XWPFTableCell docxTableCell;
	private TableRowParsingElement tableRowParsingElement;
	private int rowSpanCellNumber;
	private CTDecimalNumber gridSpan;
	private CTTblWidth cTTblWidth;
	private CTTcPr cTTcPr;

	/**
	 * @return the rowSpanCellNumber
	 */
	public int getRowSpanCellNumber() {
		return rowSpanCellNumber;
	}

	/**
	 * @param rowSpanCellNumber
	 *            the rowSpanCellNumber to set
	 */
	public void setRowSpanCellNumber(int rowSpanCellNumber) {
		this.rowSpanCellNumber = rowSpanCellNumber;
	}

	/**
	 * Constructor
	 * 
	 * @param tableRowParsingElement
	 *            table row parsing element
	 * @param document
	 *            document
	 */
	public TableCellParsingElement(
			TableRowParsingElement tableRowParsingElement, XWPFDocument document) {
		super(ElementType.TABLE_CELL, false, document);
		super.setMayContainParagraph(true);
		this.tableRowParsingElement = tableRowParsingElement;
		this.docxTableRow = tableRowParsingElement.getDocxTableRow();

		this.addRowSpanCellIfNeeded();

		this.docxTableCell = this.docxTableRow.createCell();
		CTTc cttc = docxTableCell.getCTTc();
		this.cTTcPr = cttc.getTcPr();
		if (this.cTTcPr == null) {
			this.cTTcPr = cttc.addNewTcPr();
		}
		this.createVisibleBorder(this.cTTcPr, HTMLConstants.COLOR_GREY);
		//System.out.println("Created new CELL");

	}

	/**
	 * This method adds row span cell if needed.
	 */
	private void addRowSpanCellIfNeeded() {
		int rowNum = this.docxTableRow.getTable().getNumberOfRows();
		int cellNum = this.docxTableRow.getTableCells().size();
		if (this.tableRowParsingElement.getTableParsingElement()
				.containsRowCellAtPosition(rowNum, cellNum)) {

			// System.out.println("Creating new rowspan cell at positions:"
			// + this.docxTableRow.getTable().getNumberOfRows() + "; "
			// + this.docxTableRow.getTableCells().size());
			XWPFTableCell rowSpanCell = this.docxTableRow.createCell();

			TableCellParsingElement firstRowSpanCell = this.tableRowParsingElement
					.getTableParsingElement().getFirstRowSpanCell(rowNum,
							cellNum);

			CTTc cttc = rowSpanCell.getCTTc();
			CTTcPr pr = cttc.getTcPr();
			if (pr == null) {
				pr = cttc.addNewTcPr();
			}

			if (firstRowSpanCell.getcTTblWidth() != null) {
				pr.setTcW(firstRowSpanCell.getcTTblWidth());

			}

			if (firstRowSpanCell.getGridSpan() != null) {
				pr.setGridSpan(firstRowSpanCell.getGridSpan());
			}

			CTVMerge merge = pr.getVMerge();
			if (merge == null) {
				merge = pr.addNewVMerge();
			}

			merge.setVal(STMerge.CONTINUE);

			this.createVisibleBorder(pr, HTMLConstants.COLOR_GREY);

		}
	}

	/**
	 * This class creates visible border. The default color is grey.
	 * 
	 * @param pr
	 *            CTTcPr
	 * @param color
	 *            HTML color to use
	 */
	private void createVisibleBorder(CTTcPr pr, String color) {
		CTTcBorders borders = pr.addNewTcBorders();
		CTBorder border = borders.addNewBottom();
		border.setVal(STBorder.THICK);
		border.setColor(color);
		CTBorder border1 = borders.addNewLeft();
		border1.setVal(STBorder.THICK);
		border1.setColor(color);
		CTBorder border2 = borders.addNewRight();
		border2.setVal(STBorder.THICK);
		border2.setColor(color);
		CTBorder border3 = borders.addNewTop();
		border3.setVal(STBorder.THICK);
		border3.setColor(color);
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
		this.cTTcPr = cttc.getTcPr();
		if (this.cTTcPr == null) {
			this.cTTcPr = cttc.addNewTcPr();
		}

		this.cTTblWidth = this.cTTcPr.getTcW();
		if (cTTblWidth == null) {
			cTTblWidth = this.cTTcPr.addNewTcW();
		}

		if (usePercentage) {
			CTPageSz pageSize = this.getDocument().getDocument().getBody()
					.getSectPr().getPgSz();
			BigInteger documentWidth = pageSize.getW();
			int tableWidth = (documentWidth.intValue() / 100) * width;
			cTTblWidth.setW(BigInteger.valueOf(tableWidth));

		} else {
			cTTblWidth.setW(BigInteger.valueOf(ConversionUtil
					.convertTableCellPixelsToWidthUnits(width)));
		}
	}

	/**
	 * This method sets column span.
	 * 
	 * @param columnSpan
	 *            column span
	 */
	public void setColumnSpan(int columnSpan) {
		CTTc cttc = this.docxTableCell.getCTTc();
		this.cTTcPr = cttc.getTcPr();
		if (this.cTTcPr == null) {
			this.cTTcPr = cttc.addNewTcPr();
		}

		this.gridSpan = this.cTTcPr.getGridSpan();

		if (gridSpan == null) {
			gridSpan = this.cTTcPr.addNewGridSpan();
		}
		gridSpan.setVal(BigInteger.valueOf(columnSpan));

	}

	/**
	 * This method sets row span.
	 * 
	 * @param rowSpan
	 *            row span
	 */
	public void setRowSpan(int rowSpan) {

		// System.out.println("Setting rowSpan="+rowSpan);
		if (rowSpan > 1) {
			int currentRow = this.docxTableRow.getTable().getNumberOfRows();
			this.rowSpanCellNumber = this.docxTableRow.getTableCells().size() - 1;
			for (int i = 0; i < rowSpan; i++) {
				this.tableRowParsingElement.getTableParsingElement()
						.addRowSpanCell((currentRow + i), this);
			}
			CTTc cttc = this.docxTableCell.getCTTc();
			this.cTTcPr = cttc.getTcPr();
			if (this.cTTcPr == null) {
				this.cTTcPr = cttc.addNewTcPr();
			}

			CTVMerge merge = this.cTTcPr.getVMerge();
			if (merge == null) {
				merge = this.cTTcPr.addNewVMerge();
			}

			merge.setVal(STMerge.RESTART);

		}
	}

	/**
	 * @return the tableRowParsingElement
	 */
	public TableRowParsingElement getTableRowParsingElement() {
		return tableRowParsingElement;
	}

	/**
	 * @param tableRowParsingElement
	 *            the tableRowParsingElement to set
	 */
	public void setTableRowParsingElement(
			TableRowParsingElement tableRowParsingElement) {
		this.tableRowParsingElement = tableRowParsingElement;
	}

	/**
	 * @return the gridSpan
	 */
	public CTDecimalNumber getGridSpan() {
		return gridSpan;
	}

	/**
	 * @param gridSpan
	 *            the gridSpan to set
	 */
	public void setGridSpan(CTDecimalNumber gridSpan) {
		this.gridSpan = gridSpan;
	}

	/**
	 * @return the cTTblWidth
	 */
	public CTTblWidth getcTTblWidth() {
		return cTTblWidth;
	}

	/**
	 * @param cTTblWidth
	 *            the cTTblWidth to set
	 */
	public void setcTTblWidth(CTTblWidth cTTblWidth) {
		this.cTTblWidth = cTTblWidth;
	}

	/**
	 * @return the cTTcPr
	 */
	public CTTcPr getcTTcPr() {
		return cTTcPr;
	}

	/**
	 * @param cTTcPr
	 *            the cTTcPr to set
	 */
	public void setcTTcPr(CTTcPr cTTcPr) {
		this.cTTcPr = cTTcPr;
	}

}
