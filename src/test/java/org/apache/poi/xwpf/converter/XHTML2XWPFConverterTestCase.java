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
package org.apache.poi.xwpf.converter;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.xwpf.XHTML2XWPFConverter;
import org.apache.poi.xwpf.converter.xwpf.template.DataAccess;
import org.junit.Assert;
import org.junit.Test;

/**
 * This class tests XHTML to DocX conversions. 
 * @author Anton
 *
 */
public class XHTML2XWPFConverterTestCase {

	@Test
	public void testInMemoryConversion() throws Exception {

		ByteArrayOutputStream baos = new ByteArrayOutputStream();

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("text.xhtml"), baos, null);
		Assert.assertTrue(baos.size() > 0);
	}

	@Test
	public void testConversionWithOutputToFile() throws Exception {

		File outFile = new File("target/output.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("text.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testLetterConversionWithOutputToFile() throws Exception {

		File outFile = new File("target/output_letter.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("bo.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testTableConversionWithOutputToFile() throws Exception {

		File outFile = new File("target/output_table.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("cv.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testCVWithHyperlinkConversionWithOutputToFile() throws Exception {

		File outFile = new File("target/test_hyperlink_text.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("test_hyperlink_text.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testAdvancedTableWithOutputToFile() throws Exception {

		File outFile = new File("target/advanced_table.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("advanced_table.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testAdvancedCVWithOutputToFile() throws Exception {

		File outFile = new File("target/advanced_cv.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("advanced_cv.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testAdvancedTableSubsetWithOutputToFile() throws Exception {

		File outFile = new File("target/advanced_table_subset.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("advanced_table_subset.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testFormattingTestWithOutputToFile() throws Exception {

		File outFile = new File("target/formatting_test.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("formatting_test.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testNumbersTestWithOutputToFile() throws Exception {

		File outFile = new File("target/numbers_test.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("numbers_test.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testWebTableTestWithOutputToFile() throws Exception {

		File outFile = new File("target/web_table.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("web_table.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testAdvancedFormattingAndCharsTestWithOutputToFile() throws Exception {

		File outFile = new File("target/advanced_formatting_and_chars_test.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("advanced_formatting_and_chars_test.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}
	
	@Test
	public void testCVPWithinPWithOutputToFile() throws Exception {

		File outFile = new File("target/cv_p_within_p.docx");
		outFile.getParentFile().mkdirs();

		OutputStream out = new FileOutputStream(outFile);

		XHTML2XWPFConverter.getInstance().convert(
				DataAccess.class.getResourceAsStream("cv_p_within_p.xhtml"), out, null);
		Assert.assertTrue(outFile.exists());
	}

}
