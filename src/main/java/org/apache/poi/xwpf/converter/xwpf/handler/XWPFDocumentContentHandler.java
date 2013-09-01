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

import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.xwpf.bo.XWPFOptions;
import org.apache.poi.xwpf.converter.xwpf.exception.XWPFDocumentConversionException;
import org.apache.poi.xwpf.converter.xwpf.template.DocXTemplate;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;

/**
 * This content handler supports the population of XWPFDocument from XHTML stream.
 * 
 * @author Anton
 * 
 */
public class XWPFDocumentContentHandler {

	private OutputStream out;
	private XWPFDocument document;

	/**
	 * Prevents initialization.
	 */
	@SuppressWarnings("unused")
	private XWPFDocumentContentHandler() {

	}

	/**
	 * XWPFDocumentContentHandler constructor
	 * 
	 * @param out
	 *            output stream
	 */
	public XWPFDocumentContentHandler(OutputStream out) {
		this.out = out;
	}

	/**
	 * This method creates new document
	 * 
	 * @param useTemplate
	 *            if true, will use existing template
	 */
	public void createDocument(boolean useTemplate) {
		if (useTemplate) {
			this.document = this.createDocumentFromTemplate();
		} else {
			this.document = new XWPFDocument();
		}
	}

	/**
	 * This method creates a document from pre-existing template
	 * 
	 * @return new document
	 */
	private XWPFDocument createDocumentFromTemplate() {
		XWPFDocument bulletTemplate = null;
		try {
			bulletTemplate = DocXTemplate.getEmptyBulletDocumentTemplate();

			for (int i = 0; i < bulletTemplate.getBodyElements().size(); i++) {
				bulletTemplate.removeBodyElement(i);
			}
		} catch (IOException e) {
			e.printStackTrace();
			throw new XWPFDocumentConversionException(e);
		}
		return bulletTemplate;
	}

	/**
	 * This method starts new document by adding new SectPr.
	 */
	public void startDocument() {
		this.document.getDocument().getBody().addNewSectPr();
	}

	/**
	 * This method sets document size.
	 * 
	 * @param options
	 *            processing options
	 */
	public void setDocumentSize(XWPFOptions options) {
		this.document.getDocument().getBody().getSectPr().addNewPgSz();
		CTPageSz pageSize = this.document.getDocument().getBody().getSectPr()
				.getPgSz();
		pageSize.setOrient(options.getOrientation());
		pageSize.setH(options.getPageHeight());
		pageSize.setW(options.getPageWidth());
		this.document.getDocument().getBody().getSectPr().setPgSz(pageSize);
	}

	/**
	 * This method completes the document processing by writing it out to 
	 * the output stream.
	 * 
	 * @throws IOException
	 */
	public void endDocument() {
		if (out != null) {
			try {
				this.document.write(out);
				out.flush();
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
				throw new XWPFDocumentConversionException(e);
			}
		}
	}

	/**
	 * @return the out
	 */
	public OutputStream getOut() {
		return out;
	}

	/**
	 * @param out
	 *            the out to set
	 */
	public void setOut(OutputStream out) {
		this.out = out;
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

}
