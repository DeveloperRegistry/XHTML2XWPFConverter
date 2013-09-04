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
package org.apache.poi.xwpf.converter.xwpf.handler;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.converter.xwpf.bo.AbstractParsingElement;
import org.apache.poi.xwpf.converter.xwpf.bo.HyperlinkParsingElement;
import org.apache.poi.xwpf.converter.xwpf.bo.ImageParsingElement;
import org.apache.poi.xwpf.converter.xwpf.bo.ParagraphParsingElement;
import org.apache.poi.xwpf.converter.xwpf.bo.TableCellParsingElement;
import org.apache.poi.xwpf.converter.xwpf.bo.TableParsingElement;
import org.apache.poi.xwpf.converter.xwpf.bo.TableRowParsingElement;
import org.apache.poi.xwpf.converter.xwpf.bo.XWPFOptions;
import org.apache.poi.xwpf.converter.xwpf.common.ElementType;
import org.apache.poi.xwpf.converter.xwpf.common.HTMLConstants;
import org.apache.poi.xwpf.converter.xwpf.exception.XWPFDocumentConversionException;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * Content handler class for parsing XHTML document and mapping it to DocX
 * document.
 * 
 * @author Anton
 * 
 */
public class XWPFMapper extends DefaultHandler {
	private InputStream in;
	private XWPFDocumentContentHandler docxHandler;
	private XWPFOptions options;
	private AbstractParsingElement currentTopLevelElement;
	private TableRowParsingElement currentRow;
	private StringBuffer currentTextBuffer;
	private List<AbstractParsingElement> parsingTree;

	/**
	 * Private constructor to prevent initialization.
	 */
	@SuppressWarnings("unused")
	private XWPFMapper() {

	}

	/**
	 * Creates a new instance XHTML Content Handler for parsing XHTML document.
	 * 
	 * @param in
	 *            input stream
	 * @param docxHandler
	 *            content handler for handling DocX document
	 * @param options
	 *            processing options
	 */
	public XWPFMapper(InputStream in, XWPFDocumentContentHandler docxHandler,
			XWPFOptions options) {

		super();
		this.in = in;
		this.docxHandler = docxHandler;
		this.options = options != null ? options : XWPFOptions.getDefault();

	}

	/**
	 * Maps XHTML to DocX document.
	 * 
	 * @throws Exception
	 */
	public void map() throws Exception {
		XMLReader xr = XMLReaderFactory.createXMLReader();
		xr.setContentHandler(this);
		xr.setEntityResolver(new XHTMLEntityResolver());
		xr.setErrorHandler(this);
		xr.parse(new InputSource(in));
	}

	@Override
	public final void startDocument() {
		this.docxHandler.createDocument(true);
		this.docxHandler.startDocument();
		this.docxHandler.setDocumentSize(options);
		parsingTree = new ArrayList<AbstractParsingElement>();
	}

	@Override
	public final void startElement(String uri, String name, String qName,
			Attributes atts) {

		name = this.normalizeName(name);
		this.currentTextBuffer = new StringBuffer();
		AbstractParsingElement newElement = null;

		//System.out.println("Element: " + name);

		if (HTMLConstants.HTML_TAG.equals(name)) {
			// Do nothing
		} else if (HTMLConstants.TABLE_TAG.equals(name)) {
			newElement = this.handleTableStart(atts);
		} else if (HTMLConstants.TBODY_TAG.equals(name)) {
			// Do nothing. Not needed in DocX
		} else if (HTMLConstants.TR_TAG.equals(name)) {
			newElement = this.handleTableRowStart(atts);
		} else if (HTMLConstants.TD_TAG.equals(name)) {
			newElement = this.handleTableCellStart(atts);
		} else if (HTMLConstants.P_TAG.equals(name)) {
			newElement = this.handleParagraphStart(atts);
		} else if (HTMLConstants.STRONG_TAG.equals(name)) {
			this.handleStrongStart(atts);
		} else if (HTMLConstants.A_TAG.equals(name)) {
			newElement = this.handleHyperlinkStart(atts);
		} else if (HTMLConstants.UL_TAG.equals(name)) {
			this.handleListStart(atts);
		} else if (HTMLConstants.LI_TAG.equals(name)) {
			newElement = this.handleParagraphStart(atts);
			newElement.setBullet(true);
		} else if (HTMLConstants.IMG_TAG.equals(name)) {
			newElement = this.handleImageStart(atts);
		} else {
			// development only. Remove before releasing code
			// throw new XWPFDocumentConversionException(" Unsupported tag: "
			// + name + ". Implement the tag!");
		}

		if (newElement != null) {
			this.parsingTree.add(newElement);
		}

	}

	/**
	 * This method handles image start.
	 * 
	 * @param atts
	 *            attributes
	 * @return image parsing element
	 */
	private AbstractParsingElement handleImageStart(Attributes atts) {

		ParagraphParsingElement lastParagraph = this.findLastParagraphElement();
		ImageParsingElement imageParsingElement = new ImageParsingElement(
				lastParagraph, docxHandler.getDocument(),
				this.options.isWebBasedProcessing());
		this.handleImageAttributes(atts, imageParsingElement);
		imageParsingElement.createImage();

		return imageParsingElement;
	}

	/**
	 * This method handles image attributes
	 * 
	 * @param atts
	 *            attributes
	 * @param imageParsingElement
	 *            image parsing element
	 */
	private void handleImageAttributes(Attributes atts,
			ImageParsingElement imageParsingElement) {

		for (int i = 0; atts != null && i < atts.getLength(); i++) {

			if (HTMLConstants.HTML_ATTRIBUTE_STYLE.equalsIgnoreCase(atts
					.getQName(i)) && atts.getValue(i) != null) {
				String style = atts.getValue(i).toLowerCase();

				String[] styleVariables = style.split(";");
				String styleVariable = null;

				for (int j = 0; j < styleVariables.length; j++) {
					try {
						styleVariable = styleVariables[j];
						if (styleVariable
								.contains(HTMLConstants.HTML_ATTRIBUTE_VALUE_WIDTH)) {

							String number = styleVariable
									.substring(
											styleVariable
													.indexOf(HTMLConstants.HTML_ATTRIBUTE_VALUE_WIDTH)
													+ HTMLConstants.HTML_ATTRIBUTE_VALUE_WIDTH
															.length(),
											styleVariable
													.indexOf(HTMLConstants.HTML_ATTRIBUTE_VALUE_PX));
							int imageWidth = Integer.parseInt(number);
							imageParsingElement.setWidth(imageWidth);

						}

						if (styleVariable
								.contains(HTMLConstants.HTML_ATTRIBUTE_VALUE_HEIGHT)) {
							String number = styleVariable
									.substring(
											styleVariable
													.indexOf(HTMLConstants.HTML_ATTRIBUTE_VALUE_HEIGHT)
													+ HTMLConstants.HTML_ATTRIBUTE_VALUE_HEIGHT
															.length(),
											styleVariable
													.indexOf(HTMLConstants.HTML_ATTRIBUTE_VALUE_PX));
							int imageHeight = Integer.parseInt(number);
							imageParsingElement.setHeight(imageHeight);
						}

					} catch (NumberFormatException nfe) {
						System.out.println("Unable to parse style: " + style);
					}

				}

			}

			if (HTMLConstants.HTML_ATTRIBUTE_SRC.equalsIgnoreCase(atts
					.getQName(i)) && atts.getValue(i) != null) {

				String filePath = atts.getValue(i);
				filePath = filePath.toLowerCase();

				if (filePath.startsWith(HTMLConstants.HTML_ATTRIBUTE_FILE)) {
					filePath = filePath.substring(
							HTMLConstants.HTML_ATTRIBUTE_FILE.length(),
							filePath.length());
				}
				imageParsingElement.setFilePath(filePath);
			}

		}

	}

	/**
	 * This method handles list start
	 * 
	 * @param atts
	 *            attributes
	 */
	private void handleListStart(Attributes atts) {
		// Presently do nothing
	}

	/**
	 * This method handles hyperlink start.
	 * 
	 * @param atts
	 *            tributes
	 * @return hyperlink
	 */
	private AbstractParsingElement handleHyperlinkStart(Attributes atts) {

		ParagraphParsingElement lastParagraph = this.findLastParagraphElement();

		HyperlinkParsingElement hyperlink = new HyperlinkParsingElement(
				lastParagraph, null, null, docxHandler.getDocument());

		this.handleHyperlinkAttributes(atts, hyperlink);

		return hyperlink;

	}

	/**
	 * This method handles hyperlink attributes
	 * 
	 * @param atts
	 *            attributes
	 * @param hyperlink
	 *            hyperlink
	 */
	private void handleHyperlinkAttributes(Attributes atts,
			HyperlinkParsingElement hyperlink) {

		for (int i = 0; atts != null && i < atts.getLength(); i++) {

			if (HTMLConstants.HTML_ATTRIBUTE_HREF.equalsIgnoreCase(atts
					.getQName(i)) && atts.getValue(i) != null) {
				String url = atts.getValue(i);
				hyperlink.setUrl(url);
			}

		}

	}

	/**
	 * This method handles Strong tag start
	 * 
	 * @param atts
	 *            attributes
	 */
	private void handleStrongStart(Attributes atts) {
		AbstractParsingElement lastMayContainStrongElement = this
				.findLastMayContainStrongElement();
		lastMayContainStrongElement.setStrong(true);
	}

	/**
	 * This method handles Paragraph start.
	 * 
	 * @param atts
	 *            attributes
	 * @return paragraph parsing element
	 */
	private AbstractParsingElement handleParagraphStart(Attributes atts) {

		boolean topLevel = (this.currentTopLevelElement == null);
		AbstractParsingElement containingElement = null;

		if (!topLevel) {
			containingElement = this.findLastMayContainParagraphElement();
		}

		ParagraphParsingElement paragraph = new ParagraphParsingElement(
				topLevel, containingElement, null, docxHandler.getDocument());

		if (topLevel) {
			this.currentTopLevelElement = containingElement;
		}

		this.handleParagraphAttributes(atts, paragraph);

		return paragraph;

	}

	/**
	 * This method finds the last element in the parsing tree that may contain
	 * text
	 * 
	 * @return the last element in the parsing tree that may contain text
	 */
	private AbstractParsingElement findLastMayContainTextElement() {
		AbstractParsingElement result = null;

		for (int j = this.parsingTree.size() - 1; j >= 0; j--) {
			if (this.parsingTree.get(j).isMayContainText()) {
				result = this.parsingTree.get(j);
				break;
			}
		}

		return result;
	}

	/**
	 * This method finds the last element in the parsing tree that may contain
	 * paragraph.
	 * 
	 * @return the last element in the parsing tree that may contain paragraph
	 */
	private AbstractParsingElement findLastMayContainParagraphElement() {
		AbstractParsingElement result = null;

		for (int j = this.parsingTree.size() - 1; j >= 0; j--) {
			if (this.parsingTree.get(j).isMayContainParagraph()) {
				result = this.parsingTree.get(j);
				break;
			}
		}

		return result;
	}

	/**
	 * This method finds the last element in the parsing tree that may contain
	 * text
	 * 
	 * @return the last element in the parsing tree that may contain text
	 */
	private AbstractParsingElement findLastMayContainStrongElement() {
		AbstractParsingElement result = null;

		for (int j = this.parsingTree.size() - 1; j >= 0; j--) {
			if (this.parsingTree.get(j).isMayContainStrong()) {
				result = this.parsingTree.get(j);
				break;
			}
		}

		return result;
	}

	/**
	 * This method finds the last paragraph element in the parsing tree s
	 * 
	 * @return the last paragraph element
	 */
	private ParagraphParsingElement findLastParagraphElement() {
		ParagraphParsingElement result = null;

		for (int j = this.parsingTree.size() - 1; j >= 0; j--) {
			if (ElementType.PARAGRAPH.equals(this.parsingTree.get(j).getType())) {
				result = (ParagraphParsingElement) this.parsingTree.get(j);
				break;
			}
		}

		return result;
	}

	/**
	 * This method handles Paragraph attributes.
	 * 
	 * @param atts
	 *            attributes
	 * @param paragraph
	 *            paragraph
	 */
	private void handleParagraphAttributes(Attributes atts,
			ParagraphParsingElement paragraph) {
		// Do nothing for now
	}

	/**
	 * This method handles table row start.
	 * 
	 * @param atts
	 *            attributes
	 * @return table cell parsing element
	 */
	private AbstractParsingElement handleTableCellStart(Attributes atts) {
		TableCellParsingElement cell = new TableCellParsingElement(
				this.currentRow.getDocxTableRow(), docxHandler.getDocument());
		this.handleTableCellAttributes(atts, cell);
		return cell;

	}

	/**
	 * This method handles table cell attributes.
	 * 
	 * @param atts
	 *            attributes
	 * @param cell
	 *            cells
	 */
	private void handleTableCellAttributes(Attributes atts,
			TableCellParsingElement cell) {
		// Do nothing for now
	}

	/**
	 * This method handles table row start.
	 * 
	 * @param atts
	 *            attributes
	 * @return table row parsing element
	 */
	private AbstractParsingElement handleTableRowStart(Attributes atts) {

		TableRowParsingElement row = new TableRowParsingElement(
				((TableParsingElement) this.currentTopLevelElement)
						.getDocxTable(),
				docxHandler.getDocument());
		this.currentRow = row;
		this.handleTableRowAttributes(atts, row);

		return row;
	}

	/**
	 * This method handles table attributes.
	 * 
	 * @param atts
	 *            attributes
	 * @param tableRowParsingElement
	 *            table row element
	 */
	private void handleTableRowAttributes(Attributes atts,
			TableRowParsingElement tableRowParsingElement) {
		// Presently, not attributes to handle
	}

	/**
	 * This method handles table start.
	 * 
	 * @param atts
	 *            attributes
	 * @return table element
	 */
	private AbstractParsingElement handleTableStart(Attributes atts) {

		boolean isTopLevel = (this.currentTopLevelElement == null);
		TableParsingElement tableElement = new TableParsingElement(isTopLevel,
				docxHandler.getDocument());

		this.handleTableAttributes(atts, tableElement);

		if (isTopLevel) {
			this.currentTopLevelElement = tableElement;
		}

		return tableElement;

	}

	/**
	 * This method handles table attributes.
	 * 
	 * @param atts
	 *            attributes
	 * @param tableElement
	 *            table element
	 */
	private void handleTableAttributes(Attributes atts,
			TableParsingElement tableElement) {
		for (int i = 0; atts != null && i < atts.getLength(); i++) {

			if (HTMLConstants.HTML_ATTRIBUTE_BORDER.equalsIgnoreCase(atts
					.getQName(i))
					&& atts.getValue(i) != null
					&& HTMLConstants.HTML_ATTRIBUTE_VALUE_BORDER_PRESENT
							.equals(atts.getValue(i))) {
				tableElement.setBorder(false, false, 4, 0);
			}
			if (HTMLConstants.HTML_ATTRIBUTE_CELLPADDING.equalsIgnoreCase(atts
					.getQName(i)) && atts.getValue(i) != null) {
				try {
					int padding = Integer.parseInt(atts.getValue(i));
					tableElement.setCellMargins(padding, padding, padding,
							padding);
				} catch (NumberFormatException nfe) {
					System.out.println("Unable to parse cellpadding: "
							+ atts.getValue(i));
				}
			}

			if (HTMLConstants.HTML_ATTRIBUTE_STYLE.equalsIgnoreCase(atts
					.getQName(i)) && atts.getValue(i) != null) {
				String style = atts.getValue(i).toLowerCase();

				String[] styleVariables = style.split(";");
				String styleVariable = null;

				for (int j = 0; j < styleVariables.length; j++) {
					try {
						styleVariable = styleVariables[j];
						if (styleVariable
								.contains(HTMLConstants.HTML_ATTRIBUTE_VALUE_WIDTH)) {

							this.handleTableWidthAttribute(tableElement,
									styleVariable);
						}
					} catch (NumberFormatException nfe) {
						System.out.println("Unable to parse style: " + style);
					}

				}

			}

		}
	}

	/**
	 * This method handles table Width attribute.
	 * 
	 * @param tableElement
	 *            table element
	 * @param styleVariable
	 *            xhtml width attribute
	 */
	private void handleTableWidthAttribute(TableParsingElement tableElement,
			String styleVariable) {
		boolean usePercentage = false;
		String type = null;

		if (styleVariable.contains(HTMLConstants.HTML_ATTRIBUTE_VALUE_PX)) {
			type = HTMLConstants.HTML_ATTRIBUTE_VALUE_PX;
		} else if (styleVariable
				.contains(HTMLConstants.HTML_ATTRIBUTE_VALUE_PERCENTAGE)) {
			type = HTMLConstants.HTML_ATTRIBUTE_VALUE_PERCENTAGE;
			usePercentage = true;
		} else {
			throw new XWPFDocumentConversionException(
					"Unknown width attribute: " + styleVariable);
		}

		String number = styleVariable.substring(
				styleVariable.indexOf(HTMLConstants.HTML_ATTRIBUTE_VALUE_WIDTH)
						+ HTMLConstants.HTML_ATTRIBUTE_VALUE_WIDTH.length(),
				styleVariable.indexOf(type));
		int tableWidth = Integer.parseInt(number);
		tableElement.setWidth(tableWidth, usePercentage);
	}

	@Override
	public final void endElement(String uri, String name, String qName) {

		name = this.normalizeName(name);

		if (this.currentTextBuffer != null
				&& this.currentTextBuffer.length() > 0) {
			AbstractParsingElement lastElementThatMayContainText = this
					.findLastMayContainTextElement();
			lastElementThatMayContainText
					.setParagraphData(this.currentTextBuffer);
		}

		if (HTMLConstants.HTML_TAG.equals(name)) {
			// Do nothing
		} else if (HTMLConstants.TABLE_TAG.equals(name)) {
			this.handleTableEnd();
		} else if (HTMLConstants.TBODY_TAG.equals(name)) {
			// Do nothing. Not needed in DocX
		} else if (HTMLConstants.TR_TAG.equals(name)) {
			this.handleTableRowEnd();
		} else if (HTMLConstants.TD_TAG.equals(name)) {
			this.handleTableCellEnd();
		} else if (HTMLConstants.P_TAG.equals(name)) {
			this.handleParagraphEnd();
		} else if (HTMLConstants.A_TAG.equals(name)) {
			this.handleHyperlinkEnd();
		} else if (HTMLConstants.UL_TAG.equals(name)) {
			this.handleListEnd();
		} else if (HTMLConstants.IMG_TAG.equals(name)) {
			this.handleImageEnd();
		}

		this.currentTextBuffer = null;

	}

	/**
	 * This method handles image tag end.
	 */
	private void handleImageEnd() {
		// Presently, do nothing
	}

	/**
	 * This method handles list end.
	 */
	private void handleListEnd() {
		// Presently, do nothing
	}

	/**
	 * This method handles hyperlink end.
	 */
	private void handleHyperlinkEnd() {
		// Presently, do nothing

	}

	/**
	 * This method handles paragraph end.
	 */
	private void handleParagraphEnd() {

		ParagraphParsingElement lastParagraph = this.findLastParagraphElement();
		if (lastParagraph.getParagraphData() == null) {
			lastParagraph.createEmptyRun();
		}
	}

	/**
	 * This method handles table cell end;
	 */
	private void handleTableCellEnd() {
		// Presently, do nothing

	}

	/**
	 * This method handles table row end;
	 */
	private void handleTableRowEnd() {
		this.currentRow = null;

	}

	/**
	 * This method handles table end.
	 */
	private void handleTableEnd() {
		if (this.currentTopLevelElement != null
				&& this.currentTopLevelElement.getType().equals(
						ElementType.TABLE)) {
			this.currentTopLevelElement = null;
		}
	}

	@Override
	public final void characters(char ch[], int start, int length) {

		if (this.currentTextBuffer == null) {
			this.currentTextBuffer = new StringBuffer();
		}

		for (int i = start; i < start + length; i++) {
			this.currentTextBuffer.append(ch[i]);

		}
	}

	/**
	 * The method makes all tags lower case
	 * 
	 * @param name
	 *            tag name
	 * @return normalized tag name
	 */
	private String normalizeName(String name) {
		String result = "";
		if (name != null) {
			result = name.toLowerCase();
		}

		return result;
	}

	@Override
	public final void endDocument() {
		docxHandler.endDocument();

		// System.out.println("************Parsing Tree ***********************");
		// for (AbstractParsingElement element : this.parsingTree) {
		// System.out.println("*** " + element.getType().toString()
		// + "; isTopLevel: " + element.isTopLevel() + "; isBold:"
		// + element.isStrong() + "; text: "
		// + element.getParagraphData());
		// }
	}

}
