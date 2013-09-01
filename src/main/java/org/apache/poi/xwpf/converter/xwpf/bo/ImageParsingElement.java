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

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.converter.xwpf.common.ElementType;
import org.apache.poi.xwpf.converter.xwpf.exception.XWPFDocumentConversionException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * This class encapsulates image parsing element.
 * 
 * @author Anton
 * 
 */
public class ImageParsingElement extends AbstractParsingElement {

	private ParagraphParsingElement paragraphParsingElement;
	private int pictureType;
	private String filePath;
	private int width;
	private int height;
	private boolean webBasedProcessing;

	/**
	 * Constructor
	 * 
	 * @param paragraphParsingElement
	 *            parent paragraph
	 * @param document
	 *            document
	 * @param webBasedProcessing
	 *            indicates if processed online where local images may be not
	 *            available
	 */
	public ImageParsingElement(ParagraphParsingElement paragraphParsingElement,
			XWPFDocument document, boolean webBasedProcessing) {
		super(ElementType.IMAGE, false, document);
		super.setMayContainText(false);
		this.paragraphParsingElement = paragraphParsingElement;
		this.webBasedProcessing = webBasedProcessing;
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
	 * @return the pictureType
	 */
	public int getPictureType() {
		return pictureType;
	}

	/**
	 * @param pictureType
	 *            the pictureType to set
	 */
	public void setPictureType(int pictureType) {
		this.pictureType = pictureType;
	}

	/**
	 * @return the filePath
	 */
	public String getFilePath() {
		return filePath;
	}

	/**
	 * @param filePath
	 *            the filePath to set
	 */
	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}

	/**
	 * @return the width
	 */
	public int getWidth() {
		return width;
	}

	/**
	 * @param width
	 *            the width to set
	 */
	public void setWidth(int width) {
		this.width = width;
	}

	/**
	 * @return the height
	 */
	public int getHeight() {
		return height;
	}

	/**
	 * @param height
	 *            the height to set
	 */
	public void setHeight(int height) {
		this.height = height;
	}

	/**
	 * This method creates image.
	 */
	public void createImage() {
		XWPFRun run = this.paragraphParsingElement.getDocxParagraph()
				.createRun();
		try {

			InputStream inputStream = null;
			File file = null;
			String fileName = null;
			byte[] picbytes = null;
			this.pictureType = this.getImageType(this.getFilePath());

			// If file is processed on the web (loaded from client's browser)
			// the image files will not be available for conversion. Thus, we
			// only create images place-holders.
			// We will also create only placeholder if the file is not found.
			if (webBasedProcessing) {
				inputStream = new ByteArrayInputStream("".getBytes());
			} else {
				try {
					file = new File(this.filePath);
					inputStream = new FileInputStream(file);
					fileName = file.getName();

				} catch (FileNotFoundException e) {
					picbytes = "".getBytes();
					inputStream = new ByteArrayInputStream(picbytes);
					fileName = "Unknown.jpg";
				}
			}

			run.addPicture(inputStream, this.pictureType, fileName,
					Units.toEMU(this.width), Units.toEMU(this.height));

		} catch (InvalidFormatException e) {
			e.printStackTrace();
			throw new XWPFDocumentConversionException(e);
		} catch (IOException e) {
			e.printStackTrace();
			throw new XWPFDocumentConversionException(e);
		}

	}

	/**
	 * This method returns image type based on provided file path.
	 * 
	 * @param filePath
	 *            file path
	 * @return image type
	 */
	private int getImageType(String filePath) {
		int result = 0;

		if (filePath.endsWith(".emf")) {
			result = XWPFDocument.PICTURE_TYPE_EMF;
		} else if (filePath.endsWith(".wmf")) {
			result = XWPFDocument.PICTURE_TYPE_WMF;
		} else if (filePath.endsWith(".pict")) {
			result = XWPFDocument.PICTURE_TYPE_PICT;
		} else if (filePath.endsWith(".jpeg") || filePath.endsWith(".jpg")) {
			result = XWPFDocument.PICTURE_TYPE_JPEG;
		} else if (filePath.endsWith(".png")) {
			result = XWPFDocument.PICTURE_TYPE_PNG;
		} else if (filePath.endsWith(".dib")) {
			result = XWPFDocument.PICTURE_TYPE_DIB;
		} else if (filePath.endsWith(".gif")) {
			result = XWPFDocument.PICTURE_TYPE_GIF;
		} else if (filePath.endsWith(".tiff")) {
			result = XWPFDocument.PICTURE_TYPE_TIFF;
		} else if (filePath.endsWith(".eps")) {
			result = XWPFDocument.PICTURE_TYPE_EPS;
		} else if (filePath.endsWith(".bmp")) {
			result = XWPFDocument.PICTURE_TYPE_BMP;
		} else if (filePath.endsWith(".wpg")) {
			result = XWPFDocument.PICTURE_TYPE_WPG;
		} else {
			throw new XWPFDocumentConversionException(
					"Unsupported picture: "
							+ filePath
							+ ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");

		}

		return result;

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
