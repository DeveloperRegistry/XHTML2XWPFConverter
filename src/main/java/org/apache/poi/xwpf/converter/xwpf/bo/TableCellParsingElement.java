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
		System.out.println("Created new CELL");

	}

	/**
	 * This method adds row span cell if needed.
	 */
	private void addRowSpanCellIfNeeded() {
		 int rowNum = this.docxTableRow.getTable().getNumberOfRows();
		 int cellNum = this.docxTableRow.getTableCells().size();
		if (this.tableRowParsingElement.getTableParsingElement()
				.containsRowCellAtPosition(
						rowNum,
						cellNum)) {
			// TODO: remove
			System.out.println("Creating new rowspan cell at positions:"
					+ this.docxTableRow.getTable().getNumberOfRows() + "; "
					+ this.docxTableRow.getTableCells().size());
			XWPFTableCell rowSpanCell = this.docxTableRow.createCell();
			// TODO: remove
			rowSpanCell.setText("Row span cell");
			
			TableCellParsingElement firstRowSpanCell = this.tableRowParsingElement.getTableParsingElement().getFirstRowSpanCell(rowNum, cellNum);

			// TODO: set width and height from the mother cell
			CTTc cttc = rowSpanCell.getCTTc();
			CTTcPr pr = cttc.getTcPr();
			if (pr == null) {
				pr = cttc.addNewTcPr();
			}
			
			if( firstRowSpanCell.getcTTblWidth() != null )
			{
				pr.setTcW(firstRowSpanCell.getcTTblWidth());
			}
			
			if( firstRowSpanCell.getGridSpan() != null  )
			{
				pr.setGridSpan(firstRowSpanCell.getGridSpan());
			}

			CTVMerge merge = pr.getVMerge();
			if (merge == null) {
				merge = pr.addNewVMerge();
			}

			merge.setVal(STMerge.CONTINUE);
			// TODO: remove
		//	CTTcBorders borders = pr.addNewTcBorders();
		//	CTBorder border = borders.addNewTl2Br();
		//	border.setVal(STBorder.DOUBLE);
		}
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

		this.cTTblWidth = pr.getTcW();
		if (cTTblWidth == null) {
			cTTblWidth = pr.addNewTcW();
		}

		if (usePercentage) {
			CTPageSz pageSize = this.getDocument().getDocument().getBody()
					.getSectPr().getPgSz();
			BigInteger documentWidth = pageSize.getW();
			int tableWidth = (documentWidth.intValue() / 100) * width;
			cTTblWidth.setW(BigInteger.valueOf(tableWidth));

		} else {
			cTTblWidth.setW(BigInteger.valueOf(Units.toEMU(width)));
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
		CTTcPr pr = cttc.getTcPr();
		if (pr == null) {
			pr = cttc.addNewTcPr();
		}

		this.gridSpan = pr.getGridSpan();

		if (gridSpan == null) {
			gridSpan = pr.addNewGridSpan();
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
			CTTcPr pr = cttc.getTcPr();
			if (pr == null) {
				pr = cttc.addNewTcPr();
			}

			CTVMerge merge = pr.getVMerge();
			if (merge == null) {
				merge = pr.addNewVMerge();
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

}
