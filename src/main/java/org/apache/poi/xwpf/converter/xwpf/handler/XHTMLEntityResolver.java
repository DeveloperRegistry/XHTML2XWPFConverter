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

import java.util.logging.Logger;

import org.apache.poi.xwpf.converter.xwpf.entity.DataAccess;
import org.xml.sax.EntityResolver;
import org.xml.sax.InputSource;

/**
 * Entity resolver class for resolving local DTD.
 * 
 * @author Anton
 * 
 */
public class XHTMLEntityResolver implements EntityResolver {

	public static final Logger logger = Logger
			.getLogger(XHTMLEntityResolver.class.getName());

	/**
	 * This class loads processing DTDs from the classpath. The DTDs are packaged in the 
	 * resources folder.
	 * @param public Id
	 * @param system Id
	 * @return local input source
	 */
	public InputSource resolveEntity(String publicId, String systemId) {
		InputSource result = null;

		//Extract the file name only from the path
		String[] tokens = systemId.split("/");
		String dtdFile = tokens[tokens.length - 1];

		result = new InputSource(DataAccess.class.getResourceAsStream(dtdFile));

		return result;
	}

}
