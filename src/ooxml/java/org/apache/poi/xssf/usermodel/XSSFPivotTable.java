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
import java.io.OutputStream;
import java.util.List;
import java.util.NoSuchElementException;

import javax.xml.namespace.QName;

import static org.apache.poi.POIXMLDocumentPart.DEFAULT_XML_OPTIONS;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Beta;

import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTLocation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRowFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheetSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STAxis;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STItemType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSourceType;

public class XSSFPivotTable extends POIXMLDocumentPart {
    
    protected final static short CREATED_VERSION = 3;
    protected final static short MIN_REFRESHABLE_VERSION = 3;
    protected final static short UPDATED_VERSION = 3;
    
    private CTPivotTableDefinition pivotTableDefinition;
    private XSSFPivotCacheDefinition pivotCacheDefinition;
    private XSSFPivotCache pivotCache;
    private XSSFPivotCacheRecords pivotCacheRecords;
    private Sheet parentSheet;
    private Sheet dataSheet;

    @Beta    
    protected XSSFPivotTable() {
        super();
        pivotTableDefinition = CTPivotTableDefinition.Factory.newInstance();
        pivotCache = new XSSFPivotCache();
        pivotCacheDefinition = new XSSFPivotCacheDefinition();
        pivotCacheRecords = new XSSFPivotCacheRecords();
    }

     /**
     * Creates an XSSFPivotTable representing the given package part and relationship.
     * Should only be called when reading in an existing file.
     * 
     * @param part - The package part that holds xml data representing this pivot table.
     * @param rel - the relationship of the given package part in the underlying OPC package
     */
    @Beta    
    protected XSSFPivotTable(PackagePart part, PackageRelationship rel) throws IOException {
        super(part, rel);
        readFrom(part.getInputStream());
    }

    @Beta    
    public void readFrom(InputStream is) throws IOException {
	try {
            XmlOptions options  = new XmlOptions(DEFAULT_XML_OPTIONS);
            //Removing root element
            options.setLoadReplaceDocumentElement(null);
            pivotTableDefinition = CTPivotTableDefinition.Factory.parse(is, options); 
        } catch (XmlException e) {
            throw new IOException(e.getLocalizedMessage());
        }
    }

    @Beta    
    public void setPivotCache(XSSFPivotCache pivotCache) {
        this.pivotCache = pivotCache;
    }

    @Beta    
    public XSSFPivotCache getPivotCache() {
        return pivotCache;
    }

    @Beta    
    public Sheet getParentSheet() {
        return parentSheet;
    }

    @Beta    
    public void setParentSheet(XSSFSheet parentSheet) {
        this.parentSheet = parentSheet;
    }

    @Beta    
    public CTPivotTableDefinition getCTPivotTableDefinition() {
        return pivotTableDefinition;
    }

    @Beta    
    public void setCTPivotTableDefinition(CTPivotTableDefinition pivotTableDefinition) {
        this.pivotTableDefinition = pivotTableDefinition;
    }

    @Beta    
    public XSSFPivotCacheDefinition getPivotCacheDefinition() {
        return pivotCacheDefinition;
    }

    @Beta    
    public void setPivotCacheDefinition(XSSFPivotCacheDefinition pivotCacheDefinition) {
        this.pivotCacheDefinition = pivotCacheDefinition;
    }

    @Beta    
    public XSSFPivotCacheRecords getPivotCacheRecords() {
        return pivotCacheRecords;
    }

    @Beta    
    public void setPivotCacheRecords(XSSFPivotCacheRecords pivotCacheRecords) {
        this.pivotCacheRecords = pivotCacheRecords;
    }
    
    @Beta    
    public Sheet getDataSheet() {
        return dataSheet;
    }
    
    @Beta
    private void setDataSheet(Sheet dataSheet) {
        this.dataSheet = dataSheet;
    }
    
    @Beta    
    @Override
    protected void commit() throws IOException {
        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        //Sets the pivotTableDefinition tag
        xmlOptions.setSaveSyntheticDocumentElement(new QName(CTPivotTableDefinition.type.getName().
                getNamespaceURI(), "pivotTableDefinition"));
        PackagePart part = getPackagePart();
        OutputStream out = part.getOutputStream();
        pivotTableDefinition.save(out, xmlOptions);
        out.close();
    }
    
    /**
     * Set default values for the table definition.
     */
    @Beta
    protected void setDefaultPivotTableDefinition() {
        //Not more than one until more created
        pivotTableDefinition.setMultipleFieldFilters(false);
        //Indentation increment for compact rows
        pivotTableDefinition.setIndent(0);
        //The pivot version which created the pivot cache set to default value
        pivotTableDefinition.setCreatedVersion(CREATED_VERSION);
        //Minimun version required to update the pivot cache
        pivotTableDefinition.setMinRefreshableVersion(MIN_REFRESHABLE_VERSION);
        //Version of the application which "updated the spreadsheet last"
        pivotTableDefinition.setUpdatedVersion(UPDATED_VERSION);
        //Titles shown at the top of each page when printed
        pivotTableDefinition.setItemPrintTitles(true);
        //Set autoformat properties      
        pivotTableDefinition.setUseAutoFormatting(true);
        pivotTableDefinition.setApplyNumberFormats(false);
        pivotTableDefinition.setApplyWidthHeightFormats(true);
        pivotTableDefinition.setApplyAlignmentFormats(false);
        pivotTableDefinition.setApplyPatternFormats(false);
        pivotTableDefinition.setApplyFontFormats(false);
        pivotTableDefinition.setApplyBorderFormats(false);
        pivotTableDefinition.setCacheId(pivotCache.getCTPivotCache().getCacheId());
        pivotTableDefinition.setName("PivotTable"+pivotTableDefinition.getCacheId());
        pivotTableDefinition.setDataCaption("Values");
  
        //Set the default style for the pivot table
        CTPivotTableStyle style = pivotTableDefinition.addNewPivotTableStyleInfo();
        style.setName("PivotStyleLight16");
        style.setShowLastColumn(true);
        style.setShowColStripes(false);
        style.setShowRowStripes(false);
        style.setShowColHeaders(true);
        style.setShowRowHeaders(true);
    }
    
    /**
     * Add a row label using data from the given column.
     * @param columnIndex, the index of the column to be used as row label.
     */
    @Beta
    public void addRowLabel(int columnIndex) {
        AreaReference pivotArea = new AreaReference(getPivotCacheDefinition().
                getCTPivotCacheDefinition().getCacheSource().getWorksheetSource().getRef());
        int lastRowIndex = pivotArea.getLastCell().getRow() - pivotArea.getFirstCell().getRow();
        int lastColIndex = pivotArea.getLastCell().getCol() - pivotArea.getFirstCell().getCol();
        
        if(columnIndex > lastColIndex) {
            throw new IndexOutOfBoundsException();
        }
        CTPivotFields pivotFields = pivotTableDefinition.getPivotFields();
    
        List<CTPivotField> pivotFieldList = pivotTableDefinition.getPivotFields().getPivotFieldList();
        CTPivotField pivotField = CTPivotField.Factory.newInstance();
        CTItems items = pivotField.addNewItems();

        pivotField.setAxis(STAxis.AXIS_ROW);
        pivotField.setShowAll(false);
        for(int i = 0; i <= lastRowIndex; i++) {
            items.addNewItem().setT(STItemType.DEFAULT);
        }
        items.setCount(items.getItemList().size());
        pivotFieldList.set(columnIndex, pivotField);
        
        pivotFields.setPivotFieldArray(pivotFieldList.toArray(new CTPivotField[pivotFieldList.size()]));
        
        CTRowFields rowFields;
        if(pivotTableDefinition.getRowFields() != null) {
            rowFields = pivotTableDefinition.getRowFields();
        } else {
            rowFields = pivotTableDefinition.addNewRowFields();
        }
        
        rowFields.addNewField().setX(columnIndex);
        rowFields.setCount(rowFields.getFieldList().size());
    }
    
    /**
     * Add a column label using data from the given column and specified function
     * @param columnIndex, the index of the column to be used as column label.
     * @param function, the function to be used on the data
     * The following functions exists:
     * Sum, Count, Average, Max, Min, Product, Count numbers, StdDev, StdDevp, Var, Varp
     */
    @Beta
    public void addColumnLabel(DataConsolidateFunction.Enum function, int columnIndex) {
        AreaReference pivotArea = new AreaReference(getPivotCacheDefinition().
                getCTPivotCacheDefinition().getCacheSource().getWorksheetSource().getRef());
        int lastColIndex = pivotArea.getLastCell().getCol() - pivotArea.getFirstCell().getCol();
        
        if(columnIndex > lastColIndex && columnIndex < 0) {
            throw new IndexOutOfBoundsException();
        }
        
        addDataColumn(columnIndex, true);
        addDataField(function, columnIndex);
        
        //Only add colfield if there is already one.
        if (pivotTableDefinition.getDataFields().getCount() > 1) {
            CTColFields colFields;
            if(pivotTableDefinition.getColFields() != null) {
                colFields = pivotTableDefinition.getColFields();
            } else {
                colFields = pivotTableDefinition.addNewColFields();
            }     
            colFields.addNewField().setX(-2);
            colFields.setCount(colFields.getFieldList().size());
        }
    }
    
    /**
     * Add data field with data from the given column and specified function.
     * @param function, the function to be used on the data
     * @param index, the index of the column to be used as column label.
     * The following functions exists:
     * Sum, Count, Average, Max, Min, Product, Count numbers, StdDev, StdDevp, Var, Varp
     */
    @Beta
    private void addDataField(DataConsolidateFunction.Enum function, int columnIndex) {
        AreaReference pivotArea = new AreaReference(getPivotCacheDefinition().
                getCTPivotCacheDefinition().getCacheSource().getWorksheetSource().getRef());
        int lastColIndex = pivotArea.getLastCell().getCol() - pivotArea.getFirstCell().getCol();
        
        if(columnIndex > lastColIndex && columnIndex < 0) {
            throw new IndexOutOfBoundsException();
        }
        CTDataFields dataFields;
        if(pivotTableDefinition.getDataFields() != null) {
            dataFields = pivotTableDefinition.getDataFields();
        } else {
            dataFields = pivotTableDefinition.addNewDataFields();
        }
        CTDataField dataField = dataFields.addNewDataField();
        dataField.setSubtotal(function);
        Cell cell = getDataSheet().getRow(pivotArea.getFirstCell().getRow()).getCell(columnIndex);
        cell.setCellType(Cell.CELL_TYPE_STRING);
        dataField.setName(getNameOfFunction(function) + " of " + cell.getStringCellValue());
        dataField.setFld(columnIndex);
        dataFields.setCount(dataFields.getDataFieldList().size());
    }

    /**
     * Gets the name to use for the corresponding function
     * @param function, the function which name is requested
     * @return the name
     */
    @Beta
    private String getNameOfFunction(DataConsolidateFunction.Enum function) {
        switch(function.intValue()) {
            case DataConsolidateFunction.INT_AVERAGE:
                return "Average";
            case DataConsolidateFunction.INT_COUNT:
                return "Count";
            case DataConsolidateFunction.INT_COUNT_NUMS:
                return "Count";
            case DataConsolidateFunction.INT_MAX:
                return "Max";
            case DataConsolidateFunction.INT_MIN:
                return "Min";
            case DataConsolidateFunction.INT_PRODUCT:
                return "Product";
            case DataConsolidateFunction.INT_STD_DEV:
                return "StdDev";
            case DataConsolidateFunction.INT_STD_DEVP:
                return "StdDevp";
            case DataConsolidateFunction.INT_SUM:
                return "Sum";
            case DataConsolidateFunction.INT_VAR:
                return "Var";
            case DataConsolidateFunction.INT_VARP:
                return "Varp";
        }
        throw new NoSuchElementException();
    }
    
    /**
     * Add column containing data from the referenced area.
     * @param columnIndex, the index of the column containing the data
     * @param isDataField, true if the data should be displayed in the pivot table.
     */
    @Beta
    public void addDataColumn(int columnIndex, boolean isDataField) {
        AreaReference pivotArea = new AreaReference(getPivotCacheDefinition().
                getCTPivotCacheDefinition().getCacheSource().getWorksheetSource().getRef());
        int lastColIndex = pivotArea.getLastCell().getCol() - pivotArea.getFirstCell().getCol();
        if(columnIndex > lastColIndex && columnIndex < 0) {
            throw new IndexOutOfBoundsException();
        }
        CTPivotFields pivotFields = pivotTableDefinition.getPivotFields();
        List<CTPivotField> pivotFieldList = pivotFields.getPivotFieldList();
        CTPivotField pivotField = CTPivotField.Factory.newInstance();
        
        pivotField.setDataField(isDataField);
        pivotField.setShowAll(false);
        pivotFieldList.set(columnIndex, pivotField);
        pivotFields.setPivotFieldArray(pivotFieldList.toArray(new CTPivotField[pivotFieldList.size()]));
    }
    
    /**
     * Add filter for the column with the corresponding index and cell value
     * @param columnIndex, index of column to filter on
     */
    @Beta
    public void addReportFilter(int columnIndex) {
        AreaReference pivotArea = new AreaReference(getPivotCacheDefinition().
                getCTPivotCacheDefinition().getCacheSource().getWorksheetSource().getRef());
        int lastColIndex = pivotArea.getLastCell().getCol() - pivotArea.getFirstCell().getCol();
        int lastRowIndex = pivotArea.getLastCell().getRow() - pivotArea.getFirstCell().getRow();
        
        if(columnIndex > lastColIndex && columnIndex < 0) {
            throw new IndexOutOfBoundsException();
        }
        CTPivotFields pivotFields = pivotTableDefinition.getPivotFields();
    
        List<CTPivotField> pivotFieldList = pivotTableDefinition.getPivotFields().getPivotFieldList();
        CTPivotField pivotField = CTPivotField.Factory.newInstance();
        CTItems items = pivotField.addNewItems();

        pivotField.setAxis(STAxis.AXIS_PAGE);
        pivotField.setShowAll(false);
        for(int i = 0; i <= lastRowIndex; i++) {
            items.addNewItem().setT(STItemType.DEFAULT);
        }
        items.setCount(items.getItemList().size());
        pivotFieldList.set(columnIndex, pivotField);
        
        CTPageFields pageFields;
        if (pivotTableDefinition.getPageFields()!= null) {
            pageFields = pivotTableDefinition.getPageFields();
            //Another filter has already been created
            pivotTableDefinition.setMultipleFieldFilters(true);
        } else {
            pageFields = pivotTableDefinition.addNewPageFields();
        }
        CTPageField pageField = pageFields.addNewPageField();
        pageField.setHier(-1);
        pageField.setFld(columnIndex);
        
        pageFields.setCount(pageFields.getPageFieldList().size());
        pivotTableDefinition.getLocation().setColPageCount(pageFields.getCount());
    }
    
    /**
     * Creates cacheSource and workSheetSource for pivot table and sets the source reference as well assets the location of the pivot table
     * @param source Source for data for pivot table
     * @param position Position for pivot table in sheet
     * @param sourceSheet Sheet where the source will be collected from
     */
    @Beta
    protected void createSourceReferences(AreaReference source, CellReference position, Sheet sourceSheet){
        //Get cell one to the right and one down from position, add both to AreaReference and set pivot table location.
        AreaReference destination = new AreaReference(position, new CellReference(position.getRow()+1, position.getCol()+1));
        
        CTLocation location;
        if(pivotTableDefinition.getLocation() == null) {
            location = pivotTableDefinition.addNewLocation();
            location.setFirstDataCol(1);
            location.setFirstDataRow(1);
            location.setFirstHeaderRow(1);
        } else {
            location = pivotTableDefinition.getLocation();
        }
        location.setRef(destination.formatAsString());
        pivotTableDefinition.setLocation(location);

        //Set source for the pivot table
        CTPivotCacheDefinition cacheDef = getPivotCacheDefinition().getCTPivotCacheDefinition();
        CTCacheSource cacheSource = cacheDef.addNewCacheSource();
        cacheSource.setType(STSourceType.WORKSHEET);
        CTWorksheetSource worksheetSource = cacheSource.addNewWorksheetSource();
        worksheetSource.setSheet(sourceSheet.getSheetName());
        setDataSheet(sourceSheet);
        
        String[] firstCell = source.getFirstCell().getCellRefParts();
        String[] lastCell = source.getLastCell().getCellRefParts();
        worksheetSource.setRef(firstCell[2]+firstCell[1]+':'+lastCell[2]+lastCell[1]);
    }
    
    @Beta
    protected void createDefaultDataColumns() {
        CTPivotFields pivotFields;
        if (pivotTableDefinition.getPivotFields() != null) {
            pivotFields = pivotTableDefinition.getPivotFields();
        } else {
            pivotFields = pivotTableDefinition.addNewPivotFields();
        }
        String source = pivotCacheDefinition.getCTPivotCacheDefinition().
                getCacheSource().getWorksheetSource().getRef();
        AreaReference sourceArea = new AreaReference(source);
        int firstColumn = sourceArea.getFirstCell().getCol();
        int lastColumn = sourceArea.getLastCell().getCol();
        CTPivotField pivotField;
        for(int i = 0; i<=lastColumn-firstColumn; i++) {
            pivotField = pivotFields.addNewPivotField();
            pivotField.setDataField(false);
            pivotField.setShowAll(false);
        }
        pivotFields.setCount(pivotFields.getPivotFieldList().size());
    }
}