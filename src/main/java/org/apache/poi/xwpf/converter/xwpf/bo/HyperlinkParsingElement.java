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
import org.apache.poi.xwpf.converter.xwpf.common.HTMLConstants;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlink;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * This class encapsulates a Hyperlink parsing element.
 * 
 * @author Anton
 * 
 */
public class HyperlinkParsingElement extends AbstractParsingElement {

	private ParagraphParsingElement paragraphParsingElement;
	private XWPFHyperlink docxHyperlink;
	private String id;
	private String url;

	/**
	 * Constructor
	 * 
	 * 
	 * @param paragraphParsingElement
	 *            paragraph parsing element
	 * @param id
	 *            id
	 * @param url
	 *            url
	 * @param document
	 *            document
	 */
	public HyperlinkParsingElement(
			ParagraphParsingElement paragraphParsingElement, String id,
			String url, XWPFDocument document) {
		super(ElementType.HYPERLINK, false, document);
		super.setMayContainText(true);
		super.setMayContainStrong(true);
		this.paragraphParsingElement = paragraphParsingElement;
	}

	/**
	 * @return the paragraphParsingElement
	 */
	public ParagraphParsingElement getParagraphParsingElement() {
		return paragraphParsingElement;
	}

	/**
	 * @param paragraphParsingElement
	 *            the paragraphParsingElement to set
	 */
	public void setParagraphParsingElement(
			ParagraphParsingElement paragraphParsingElement) {
		this.paragraphParsingElement = paragraphParsingElement;
	}

	/**
	 * @return the docxHyperlink
	 */
	public XWPFHyperlink getDocxHyperlink() {
		return docxHyperlink;
	}

	/**
	 * @param docxHyperlink
	 *            the docxHyperlink to set
	 */
	public void setDocxHyperlink(XWPFHyperlink docxHyperlink) {
		this.docxHyperlink = docxHyperlink;
	}

	/**
	 * @return the id
	 */
	public String getId() {
		return id;
	}

	/**
	 * @param id
	 *            the id to set
	 */
	public void setId(String id) {
		// Critical Note: creation of hyperlinks presently not supported by POI
		this.id = id;
	}

	@Override
	public void setParagraphData(StringBuffer paragraphData) {
		//docxHyperlink = new XWPFHyperlink(this.id, this.url);
		// CTHyperlink hyperlink =
		// this.paragraphParsingElement.getDocxParagraph().getCTP().addNewHyperlink();
		// hyperlink.setId(paragraphData.toString());

		// XWPFHyperlinkRun run = null;
		// this.paragraphParsingElement.getDocxParagraph().addRun(run);
		XWPFRun run = this.paragraphParsingElement.getDocxParagraph()
				.createRun(); // create run object in the paragraph
		run.setBold(this.isStrong());	
		run.setColor(HTMLConstants.COLOR_HYPERLINK_BLUE);
		run.setText(paragraphData.toString());
		
		//System.out.println("Created new run for hyperlink on paragraph: "+this.paragraphParsingElement);

	}

	/**
	 * @return the url
	 */
	public String getUrl() {
		return url;
	}

	/**
	 * @param url
	 *            the url to set
	 */
	public void setUrl(String url) {
		this.url = url;
	}

}
