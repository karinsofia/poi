/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.apache.poi.xssf.usermodel;

import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.POIXMLDocumentPart;
import static org.apache.poi.POIXMLDocumentPart.DEFAULT_XML_OPTIONS;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCache;

public class XSSFPivotCache extends POIXMLDocumentPart {
        
    private CTPivotCache ctPivotCache;
    
    public XSSFPivotCache(){
        super();
        ctPivotCache = CTPivotCache.Factory.newInstance();
    }
    
    public XSSFPivotCache(CTPivotCache ctPivotCache){
        super();
        this.ctPivotCache = ctPivotCache;
    }
    
     /**
     * Creates an XSSFPivotCache representing the given package part and relationship.
     *
     * @param part - The package part that holds xml data representing this pivot cache definition.
     * @param rel - the relationship of the given package part in the underlying OPC package
     */
    protected XSSFPivotCache(PackagePart part, PackageRelationship rel) throws IOException {
        super(part, rel);
        readFrom(part.getInputStream());
    }
    
    protected void readFrom(InputStream is) throws IOException {
	try {
        XmlOptions options  = new XmlOptions(DEFAULT_XML_OPTIONS);
        //Removing root element
        options.setLoadReplaceDocumentElement(null);
            ctPivotCache = CTPivotCache.Factory.parse(is, options); 
        } catch (XmlException e) {
            throw new IOException(e.getLocalizedMessage());
        }
    }
    
    public CTPivotCache getCTPivotCache() {
        return ctPivotCache;
    }
}
