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
