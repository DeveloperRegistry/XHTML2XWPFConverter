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
package org.apache.poi.xwpf.converter.xwpf.exception;

import java.io.Serializable;

/**
 * Document conversion exception.
 * 
 * @author Anton
 * 
 */
public class XWPFDocumentConversionException extends RuntimeException implements
		Serializable {
	/**
	 * 
	 */
	private static final long serialVersionUID = 7010911517715237915L;

	// ~ Constructors
	// -----------------------------------------------------------

	/**
	 * Default constructor.
	 */
	public XWPFDocumentConversionException() {
		super();
	}

	/**
	 * Constructor that takes a String message describing this exception.
	 * 
	 * @param messageArg
	 *            Exception message as a String.
	 */
	public XWPFDocumentConversionException(String messageArg) {
		super(messageArg);
	}

	/**
	 * Constructor that takes a {@link java.lang.Throwable} object that is the
	 * parent exception of this exception. The parent exception will be chained.
	 * 
	 * @param excepArg
	 *            A Throwable object
	 */
	public XWPFDocumentConversionException(Throwable excepArg) {
		super(excepArg);
	}

	/**
	 * Contructor that takes a String message and a Throwable object.
	 * 
	 * @param messageArg
	 *            Exception message.
	 * @param excepArg
	 *            Instance of Throwable object to nest in this exception.
	 */
	public XWPFDocumentConversionException(String messageArg, Throwable excepArg) {
		super(messageArg, excepArg);
	}
}
