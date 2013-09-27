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

import java.util.ArrayList;

import org.apache.poi.xwpf.converter.xwpf.common.ElementType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * The abstract parsing element class.
 * 
 * 
 * @author Anton
 * 
 */
public class AbstractParsingElement {

	private boolean topLevel;
	private ElementType type;
	private XWPFDocument document;
	private boolean mayContainText;
	private boolean mayContainStrong;
	private boolean mayContainItalic;
	private boolean mayContainStrikeThrough;
	private boolean mayContainBullet;
	private boolean mayContainParagraph;
	private boolean strong;
	private boolean bullet;
	private boolean italic;
	private boolean strikeThrough;
	private StringBuffer paragraphData;
	private ArrayList<AbstractParsingElement> innerElements = new ArrayList<AbstractParsingElement>();

	/**
	 * Private constructor to prevent initialization.
	 */
	@SuppressWarnings("unused")
	private AbstractParsingElement() {

	}

	/**
	 * Constructor.
	 * 
	 * @param type
	 *            parsing element type
	 * @param topLevel
	 *            indicates if top level
	 * @param document
	 *            DocX document
	 */
	public AbstractParsingElement(ElementType type, boolean topLevel,
			XWPFDocument document) {
		this.type = type;
		this.topLevel = topLevel;
		this.document = document;
	}

	/**
	 * Adds new inner element.
	 * 
	 * @param innerElement
	 *            inner element to add
	 */
	public void addInnerParsingElement(AbstractParsingElement innerElement) {
		this.innerElements.add(innerElement);
	}

	/**
	 * @return the topLevel
	 */
	public boolean isTopLevel() {
		return topLevel;
	}

	/**
	 * @param topLevel
	 *            the topLevel to set
	 */
	public void setTopLevel(boolean topLevel) {
		this.topLevel = topLevel;
	}

	/**
	 * @return the type
	 */
	public ElementType getType() {
		return type;
	}

	/**
	 * @param type
	 *            the type to set
	 */
	public void setType(ElementType type) {
		this.type = type;
	}

	/**
	 * @return the innerElements
	 */
	public ArrayList<AbstractParsingElement> getInnerElements() {
		return innerElements;
	}

	/**
	 * @param innerElements
	 *            the innerElements to set
	 */
	public void setInnerElements(ArrayList<AbstractParsingElement> innerElements) {
		this.innerElements = innerElements;
	}

	/**
	 * @return the document
	 */
	public XWPFDocument getDocument() {
		return document;
	}

	/**
	 * @param document
	 *            the document to set
	 */
	public void setDocument(XWPFDocument document) {
		this.document = document;
	}

	/**
	 * @return the mayContainText
	 */
	public boolean isMayContainText() {
		return mayContainText;
	}

	/**
	 * @param mayContainText
	 *            the mayContainText to set
	 */
	public void setMayContainText(boolean mayContainText) {
		this.mayContainText = mayContainText;
	}

	/**
	 * @return the mayContainStrong
	 */
	public boolean isMayContainStrong() {
		return mayContainStrong;
	}

	/**
	 * @param mayContainStrong
	 *            the mayContainStrong to set
	 */
	public void setMayContainStrong(boolean mayContainStrong) {
		this.mayContainStrong = mayContainStrong;
	}

	/**
	 * @return the strong
	 */
	public boolean isStrong() {
		return strong;
	}

	/**
	 * @param strong
	 *            the strong to set
	 */
	public void setStrong(boolean strong) {
		this.strong = strong;
	}

	/**
	 * @return the paragraphData
	 */
	public StringBuffer getParagraphData() {
		return paragraphData;
	}

	/**
	 * @param paragraphData
	 *            the paragraphData to set
	 */
	public void setParagraphData(StringBuffer paragraphData) {
		this.paragraphData = paragraphData;
	}

	/**
	 * @return the mayContainBullet
	 */
	public boolean isMayContainBullet() {
		return mayContainBullet;
	}

	/**
	 * @param mayContainBullet
	 *            the mayContainBullet to set
	 */
	public void setMayContainBullet(boolean mayContainBullet) {
		this.mayContainBullet = mayContainBullet;
	}

	/**
	 * @return the bullet
	 */
	public boolean isBullet() {
		return bullet;
	}

	/**
	 * @param bullet
	 *            the bullet to set
	 */
	public void setBullet(boolean bullet) {
		this.bullet = bullet;
	}

	/**
	 * @return the mayContainParagraph
	 */
	public boolean isMayContainParagraph() {
		return mayContainParagraph;
	}

	/**
	 * @param mayContainParagraph
	 *            the mayContainParagraph to set
	 */
	public void setMayContainParagraph(boolean mayContainParagraph) {
		this.mayContainParagraph = mayContainParagraph;
	}

	/**
	 * @return the mayContainItalic
	 */
	public boolean isMayContainItalic() {
		return mayContainItalic;
	}

	/**
	 * @param mayContainItalic
	 *            the mayContainItalic to set
	 */
	public void setMayContainItalic(boolean mayContainItalic) {
		this.mayContainItalic = mayContainItalic;
	}

	/**
	 * @return the mayContainStrikeThrough
	 */
	public boolean isMayContainStrikeThrough() {
		return mayContainStrikeThrough;
	}

	/**
	 * @param mayContainStrikeThrough
	 *            the mayContainStrikeThrough to set
	 */
	public void setMayContainStrikeThrough(boolean mayContainStrikeThrough) {
		this.mayContainStrikeThrough = mayContainStrikeThrough;
	}

	/**
	 * @return the italic
	 */
	public boolean isItalic() {
		return italic;
	}

	/**
	 * @param italic
	 *            the italic to set
	 */
	public void setItalic(boolean italic) {
		this.italic = italic;
	}

	/**
	 * @return the strikeThrough
	 */
	public boolean isStrikeThrough() {
		return strikeThrough;
	}

	/**
	 * @param strikeThrough
	 *            the strikeThrough to set
	 */
	public void setStrikeThrough(boolean strikeThrough) {
		this.strikeThrough = strikeThrough;
	}

}
