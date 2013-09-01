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

import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

/**
 * This class encapsulates conversion options.
 * 
 * @author Anton
 * 
 */
public class XWPFOptions {

	private static final XWPFOptions DEFAULT = new XWPFOptions();
	private STPageOrientation.Enum orientation;
	private BigInteger pageHeight;
	private BigInteger pageWidth;
	// Images cannot be processed on the web
	private boolean webBasedProcessing;

	private XWPFOptions() {
		this.orientation = STPageOrientation.PORTRAIT;
		this.pageHeight = BigInteger.valueOf(16383);
		this.pageWidth = BigInteger.valueOf(11906);
		this.webBasedProcessing = false;

	}

	/**
	 * Returns the instance of the options.
	 * 
	 * @return instance of the options
	 */
	public static XWPFOptions create() {
		return new XWPFOptions();
	}

	/**
	 * Returns default instance of the options.
	 * 
	 * @return default instance of the options
	 */
	public static XWPFOptions getDefault() {
		return DEFAULT;
	}

	/**
	 * @return the orientation
	 */
	public STPageOrientation.Enum getOrientation() {
		return orientation;
	}

	/**
	 * @param orientation
	 *            the orientation to set
	 */
	public void setOrientation(STPageOrientation.Enum orientation) {
		this.orientation = orientation;
	}

	/**
	 * @return the pageHeight
	 */
	public BigInteger getPageHeight() {
		return pageHeight;
	}

	/**
	 * @param pageHeight
	 *            the pageHeight to set
	 */
	public void setPageHeight(BigInteger pageHeight) {
		this.pageHeight = pageHeight;
	}

	/**
	 * @return the pageWidth
	 */
	public BigInteger getPageWidth() {
		return pageWidth;
	}

	/**
	 * @param pageWidth
	 *            the pageWidth to set
	 */
	public void setPageWidth(BigInteger pageWidth) {
		this.pageWidth = pageWidth;
	}

	/**
	 * @return the webBasedProcessing
	 */
	public boolean isWebBasedProcessing() {
		return webBasedProcessing;
	}

	/**
	 * @param webBasedProcessing
	 *            the webBasedProcessing to set
	 */
	public void setWebBasedProcessing(boolean webBasedProcessing) {
		this.webBasedProcessing = webBasedProcessing;
	}

}
