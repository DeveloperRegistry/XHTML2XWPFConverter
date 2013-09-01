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
package org.apache.poi.xwpf.converter.xwpf;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Logger;

import org.apache.poi.xwpf.converter.xwpf.bo.XWPFOptions;
import org.apache.poi.xwpf.converter.xwpf.exception.XWPFDocumentConversionException;
import org.apache.poi.xwpf.converter.xwpf.handler.XWPFConverter;

public class XHTML2XWPFConverter {

	private static final XHTML2XWPFConverter INSTANCE = new XHTML2XWPFConverter();

	/**
	 * Logger for this class
	 */
	private static final Logger LOGGER = Logger
			.getLogger(XHTML2XWPFConverter.class.getName());

	/**
	 * Returns a static instance of the converter.
	 * 
	 * @return a static instance of the converter
	 */
	public static XHTML2XWPFConverter getInstance() {
		return INSTANCE;
	}

	/**
	 * This method converts XHTML stream into XWPFDocument.
	 * 
	 * @param in
	 *            input stream
	 * @param output
	 *            stream
	 * @param options
	 *            conversion options
	 */
	public void convert(InputStream in, OutputStream out, XWPFOptions options)
			throws XWPFDocumentConversionException {
		try {
			XWPFConverter.getInstance().doConvert(in, out, options);
		} catch (Exception e) {
			LOGGER.severe(e.getMessage());
			throw new XWPFDocumentConversionException(e);
		}
	}

}
