package com.adaptris.core.poi;

import static com.adaptris.core.poi.ExcelHelper.XML_ATTR_POSITION;
import static com.adaptris.core.poi.ExcelHelper.XML_ATTR_ROW_NUMBER;
import static com.adaptris.core.poi.ExcelHelper.XML_ATTR_SHEET_NAME;
import static com.adaptris.core.poi.ExcelHelper.XML_ATTR_TYPE;
import static com.adaptris.core.poi.ExcelHelper.XML_ELEMENT_ROW;
import static com.adaptris.core.poi.ExcelHelper.XML_ELEMENT_SPREADSHEET;
import static com.adaptris.core.poi.ExcelHelper.XML_ELEMENT_WORKSHEET;
import static com.adaptris.core.poi.ExcelHelper.createCellName;
import static com.adaptris.core.poi.ExcelHelper.getCellCount;
import static com.adaptris.core.poi.ExcelHelper.getRowCount;

import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.adaptris.core.poi.ExcelHelper.CellHandler;
import com.adaptris.core.poi.XmlStyle.ElementNaming;

class ExcelConverter {

  interface ExcelConverterContext {
    boolean ignoreNullRows();

    Logger logger();
  }

  private transient ExcelConverterContext context;

  ExcelConverter(ExcelConverterContext context) {
    this.context = context;
  }

  final Document convertToXml(Workbook workbook, XmlStyle styleGuide) throws Exception {
    Document document = createDocument();

    context.logger().trace("workbook has {} sheets", workbook.getNumberOfSheets());
    Element rootElement = document.createElement(XML_ELEMENT_SPREADSHEET);
    document.appendChild(rootElement);
    for (int sheetCounter = 0; sheetCounter < workbook.getNumberOfSheets(); sheetCounter++) {
      Sheet sheet = workbook.getSheetAt(sheetCounter);
      processWorksheet(sheet, rootElement, styleGuide);
    }
    return document;
  }

  private void processWorksheet(Sheet sheet, Element parent, XmlStyle styleGuide) throws Exception {
    Document document = parent.getOwnerDocument();
    Element sheetElement = document.createElement(XML_ELEMENT_WORKSHEET);
    parent.appendChild(sheetElement);
    sheetElement.setAttribute(XML_ATTR_SHEET_NAME, sheet.getSheetName());
    int nRows = getRowCount(sheet);
    int nCells = getCellCount(sheet);
    context.logger().trace("Sheet {} has {} rows with {} cells", sheet.getSheetName(), nRows, nCells);
    if (nRows > 0) {
      int rowCounter = sheet.getFirstRowNum();
      String[] columnNames = createColumnNames(sheet, styleGuide);
      if (styleGuide.resolveNamingStrategy() == ElementNaming.HEADER_ROW) {
        rowCounter = styleGuide.headerRow();
      }
      // Now loop through and create each row.
      for (; rowCounter < nRows; rowCounter++) {
        Row row = sheet.getRow(rowCounter);
        if (row == null) {
          if (!context.ignoreNullRows()) {
            throw new Exception("Unable to process row " + (rowCounter + 1) + "; it's null");
          }
        } else {
          processRow(row, sheetElement, styleGuide, columnNames);
        }
      }
    }
  }

  private void processRow(Row row, Element parent, XmlStyle styleGuide, String[] columnNames) throws Exception {
    Document document = parent.getOwnerDocument();
    Element rowElement = document.createElement(XML_ELEMENT_ROW);
    parent.appendChild(rowElement);
    int rowCounter = row.getRowNum() + 1;
    context.logger().trace("Processing row {}", rowCounter);
    if (styleGuide.emitRowNumberAttr()) {
      rowElement.setAttribute(XML_ATTR_ROW_NUMBER, String.valueOf(rowCounter));
    }
    for (int i = 0; i < columnNames.length; i++) {
      context.logger().trace("Creating element [{}]", columnNames[i]);
      Element cellElement = document.createElement(columnNames[i]);
      rowElement.appendChild(cellElement);
      Cell cell = row.getCell(i);
      CellHandler handler = ExcelHelper.getHandler(cell);
      String value = handler.getValue(cell, styleGuide);
      String type = handler.getType();
      context.logger().trace("Cell ({},{}) is a [{}] and the computed value is [{}]", rowCounter, i + 1, type, value);
      // if showDataTypes is true, add the attribute to the xml node
      if (styleGuide.emitDataTypeAttr()) {
        cellElement.setAttribute(XML_ATTR_TYPE, type);
      }
      // if showCellPosition is true, add the positional element to the XML (e.g. BB0).
      if (styleGuide.emitCellPositionAttr()) {
        cellElement.setAttribute(XML_ATTR_POSITION, createCellName(i) + rowCounter);
      }
      // set the value of the xml node
      cellElement.setTextContent(value);
    }
  }

  private String[] createColumnNames(Sheet sheet, XmlStyle lf) {
    return lf.resolveNamingStrategy().createColumnNames(sheet, lf);
  }

  public static Document createDocument() throws Exception {
    DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
    return dbf.newDocumentBuilder().newDocument();
  }

}
