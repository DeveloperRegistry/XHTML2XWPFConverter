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

import org.apache.poi.xwpf.converter.xwpf.common.ElementType;
import org.apache.poi.xwpf.converter.xwpf.common.StyleConstants;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHighlight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;

/**
 * This class encapsulates a Paragraph parsing element.
 * 
 * @author Anton
 * 
 */
public class ParagraphParsingElement extends AbstractParsingElement {

	private AbstractParsingElement containingElement;
	private StringBuffer paragraphData;
	private XWPFParagraph docxParagraph;
	private boolean horizontalLine; // to support HR tag
	private String fontColor;
	private boolean highlightSpan;
	private boolean standAloneSpan;

	/**
	 * Constructor
	 * 
	 * 
	 * @param topLevel
	 *            if true, it is an independent top level element
	 * @param containingElement
	 *            containing element
	 * @param paragraphData
	 *            paragraph data
	 * @param document
	 *            document
	 */
	public ParagraphParsingElement(boolean topLevel,
			AbstractParsingElement containingElement,
			StringBuffer paragraphData, XWPFDocument document) {
		super(ElementType.PARAGRAPH, topLevel, document);
		this.initializeCommonParagraphFields();
		this.containingElement = containingElement;
		this.paragraphData = paragraphData;

		if (topLevel) {
			this.docxParagraph = document.createParagraph();
		}

		if (!topLevel && containingElement != null
				&& ElementType.TABLE_CELL.equals(containingElement.getType())) {
			TableCellParsingElement cell = (TableCellParsingElement) containingElement;
			this.docxParagraph = cell.getDocxTableCell().addParagraph();

		}

		// System.out.println("Created paragraph: "
		// + this.docxParagraph
		// + "; topLevel="
		// + topLevel
		// + "; containingElement: "
		// + ((this.containingElement != null) ? this.containingElement
		// .getType() : ""));
	}

	/**
	 * Constructor.
	 * 
	 * @param cursor
	 *            cursor where new paragraph will be inserted
	 * @param document
	 *            document
	 */
	public ParagraphParsingElement(XmlCursor cursor, XWPFDocument document) {
		super(ElementType.PARAGRAPH, false, document);
		this.initializeCommonParagraphFields();
		this.containingElement = null;
		this.paragraphData = null;

		this.docxParagraph = document.insertNewParagraph(cursor);
		// System.out.println("Created paragraph: "
		// + this.docxParagraph
		// + "; topLevel="
		// + topLevel
		// + "; containingElement: "
		// + ((this.containingElement != null) ? this.containingElement
		// .getType() : ""));

	}

	/**
	 * This method initializes common paragraph fields
	 */
	private void initializeCommonParagraphFields() {
		super.setMayContainText(true);
		super.setMayContainStrong(true);
		super.setMayContainItalic(true);
		super.setMayContainStrikeThrough(true);
		super.setMayContainBullet(true);
		super.setMayContainNumbering(true);
		super.setMayContainUnderline(true);
		super.setMayContainHeading(true);
	}

	/**
	 * @return the containingElement
	 */
	public AbstractParsingElement getContainingElement() {
		return containingElement;
	}

	/**
	 * @param containingElement
	 *            the containingElement to set
	 */
	public void setContainingElement(AbstractParsingElement containingElement) {
		this.containingElement = containingElement;
	}

	/**
	 * @return the paragraphData
	 */
	public StringBuffer getParagraphData() {
		return paragraphData;
	}

	@Override
	public void setParagraphData(StringBuffer paragraphData) {

		// System.out.println("Setting paragraphData="+paragraphData);
		this.paragraphData = paragraphData;

		String para = paragraphData.toString();

		if (this.isBullet()) {
			this.setListParagraphStyle();
			this.docxParagraph.setNumID(BigInteger.valueOf(1));
		}

		if (this.isNumbering()) {
			this.setListParagraphStyle();
			this.docxParagraph.setNumID(BigInteger.valueOf(this
					.getNumberedListValue()));
		}

		XWPFRun run = this.docxParagraph.createRun(); // create run object in
														// the paragraph
		run.setBold(this.isStrong());
		run.setItalic(this.isItalic());
		run.setStrike(this.isStrikeThrough());
		if (this.isUnderline()) {
			run.setUnderline(UnderlinePatterns.THICK);
		}
		if (this.isFontColorSet()) {
			run.setColor(this.fontColor);
		}

		if (this.isHeadingLevelSet()) {
			this.docxParagraph.setStyle(this.getHeadingLevel());
		}

		if (this.isHighlightSpan()) {
			CTRPr cTRPr = run.getCTR().getRPr() != null ? run.getCTR().getRPr()
					: run.getCTR().addNewRPr();
			CTHighlight highlight = cTRPr.addNewHighlight();
			highlight.setVal(STHighlightColor.YELLOW);
			this.highlightSpan = false;
		}

		run.setText(para);
		// System.out.println("Created new run for paragraph: " + para
		// + "; docxPara=" + this.docxParagraph);

	}

	/**
	 * This method sets list paragraph style.
	 */
	private void setListParagraphStyle() {
		CTPPr ppr = this.docxParagraph.getCTP().addNewPPr();
		CTString style = ppr.addNewPStyle();
		style.setVal(StyleConstants.LIST_PARAGRAPH);
	}

	/**
	 * This method creates an empty run.
	 */
	public void createEmptyRun() {
		this.docxParagraph.createRun();
	}

	/**
	 * @return the docxParagraph
	 */
	public XWPFParagraph getDocxParagraph() {
		return docxParagraph;
	}

	/**
	 * @param docxParagraph
	 *            the docxParagraph to set
	 */
	public void setDocxParagraph(XWPFParagraph docxParagraph) {
		this.docxParagraph = docxParagraph;
	}

	/**
	 * This method adds a new line break.
	 */
	public void addLineBreak() {
		XWPFRun run = this.docxParagraph.createRun();
		run.addCarriageReturn();
	}

	/**
	 * @return the horizontalLine
	 */
	public boolean isHorizontalLine() {
		return horizontalLine;
	}

	/**
	 * @param horizontalLine
	 *            the horizontalLine to set
	 */
	public void setHorizontalLine(boolean horizontalLine) {
		this.horizontalLine = horizontalLine;

		if (this.horizontalLine) {
			super.setMayContainText(false);
			super.setMayContainStrong(false);
			super.setMayContainItalic(false);
			super.setMayContainStrikeThrough(false);
			super.setMayContainBullet(false);
			super.setMayContainNumbering(false);
			super.setMayContainUnderline(false);
			super.setMayContainHeading(false);
			this.docxParagraph.setBorderBottom(Borders.SINGLE);
		}
	}

	/**
	 * @return the fontColor
	 */
	public String getFontColor() {
		return fontColor;
	}

	/**
	 * @param fontColor
	 *            the fontColor to set
	 */
	public void setFontColor(String fontColor) {
		this.fontColor = fontColor;
	}

	/**
	 * This method returns true if font color is set.
	 * 
	 * @return true if font color is set
	 */
	public boolean isFontColorSet() {
		return (this.fontColor != null);
	}

	/**
	 * @return the highlightSpan
	 */
	public boolean isHighlightSpan() {
		return highlightSpan;
	}

	/**
	 * @param highlightSpan
	 *            the highlightSpan to set
	 */
	public void setHighlightSpan(boolean highlightSpan) {
		this.highlightSpan = highlightSpan;
	}

	/**
	 * @return the standAloneSpan
	 */
	public boolean isStandAloneSpan() {
		return standAloneSpan;
	}

	/**
	 * @param standAloneSpan
	 *            the standAloneSpan to set
	 */
	public void setStandAloneSpan(boolean standAloneSpan) {
		this.standAloneSpan = standAloneSpan;
	}

}
