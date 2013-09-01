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
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.xwpf.bo.XWPFOptions;
import org.apache.poi.xwpf.converter.xwpf.exception.XWPFDocumentConversionException;

/**
 * This class represents a converter from Strict XHTML to Microsoft Word (docx)
 * document.
 * 
 * @author Anton
 * 
 */
public class XWPFConverter {

	private static final XWPFConverter INSTANCE = new XWPFConverter();

	public static XWPFConverter getInstance() {
		return INSTANCE;
	}

	/**
	 * This method converts XHTML document into XWPFDocument document.
	 * 
	 * @param in
	 *            input stream with XML document
	 * @param out
	 *            output stream with DocX document
	 * @param options
	 *            conversion options.
	 */
	public void doConvert(InputStream in, OutputStream out, XWPFOptions options)
			throws XWPFDocumentConversionException, IOException {

		options = options != null ? options : XWPFOptions.getDefault();
		XWPFDocumentContentHandler contentHandler = new XWPFDocumentContentHandler(
				out);
		convert(in, contentHandler, options);
	}

	/**
	 * This method converts XHTML document into XWPFDocument document.
	 * 
	 * @param in
	 *            input stream with XML document
	 * @param contentHandler
	 *            content handler
	 * @param options
	 *            conversion options.
	 */
	protected void convert(InputStream in,
			XWPFDocumentContentHandler contentHandler, XWPFOptions options)
			throws XWPFDocumentConversionException, IOException {
		try {
			options = options != null ? options : XWPFOptions.getDefault();
			XWPFMapper mapper = new XWPFMapper(in, contentHandler, options);
			mapper.map();
		} catch (Exception e) {
			throw new XWPFDocumentConversionException(e);
		}
	}

}