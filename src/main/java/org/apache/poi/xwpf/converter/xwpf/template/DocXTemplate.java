/**
 * 
 */
package org.apache.poi.xwpf.converter.xwpf.template;

import java.io.IOException;

import org.apache.poi.xwpf.converter.xwpf.common.TemplateConstants;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * This class povides access to word template documents. 
 * @author Anton
 * 
 */
public class DocXTemplate {

	/**
	 * This method returns sample  DocX template with bullets. 
	 * @return
	 * @throws IOException
	 */
	public static XWPFDocument getBulletDocumentTemplate() throws IOException {
		XWPFDocument document = new XWPFDocument(
				DataAccess.class.getResourceAsStream(TemplateConstants.FILE_NAME_BULLET_TEMPLATE));
		return document;
	}
	
	/**
	 * This method returns empty  DocX template with bullets. 
	 * @return
	 * @throws IOException
	 */
	public static XWPFDocument getEmptyBulletDocumentTemplate() throws IOException {
		XWPFDocument document = new XWPFDocument(
				DataAccess.class.getResourceAsStream(TemplateConstants.FILE_NAME_EMPTY_BULLET_TEMPLATE));
		return document;
	}
	
	/**
	 * This method returns simple  DocX template. 
	 * @return
	 * @throws IOException
	 */
	public static XWPFDocument getSimpleDocumentTemplate() throws IOException {
		XWPFDocument document = new XWPFDocument(
				DataAccess.class.getResourceAsStream(TemplateConstants.FILE_NAME_SIMPLE_TEMPLATE));
		return document;
	}

}
