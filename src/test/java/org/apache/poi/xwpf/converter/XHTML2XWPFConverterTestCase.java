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

}
