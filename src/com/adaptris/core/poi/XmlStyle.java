package com.adaptris.core.poi;

import static com.adaptris.core.poi.ExcelHelper.createCellName;
import static com.adaptris.core.poi.ExcelHelper.getCellCount;
import static com.adaptris.core.poi.ExcelHelper.safeName;
import static org.apache.commons.lang.StringUtils.isEmpty;

import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Controls the XML look and feel for ExcelToXml.
 *
 * @author lchan
 *
 */
public class XmlStyle {
  private boolean emitDataTypeAttr;
  private boolean emitRowNumberAttr;
  private boolean emitCellPositionAttr;
  private String dateFormat;
  private String numberFormat;
  private Integer headerRow;

  private String elementNamingStyle;
  private String xmlEncoding;

  private transient DateFormat dateFormatter = null;
  private transient NumberFormat numberFormatter = null;

  public enum ElementNaming {
    SIMPLE {
      @Override
      void applyNames(Sheet sheet, String[] names, XmlStyle style) {
        for (int i = 0; i < names.length; i++) {
          names[i] = "cell";
        }
      }

    },
    CELL_POSITION {

      @Override
      void applyNames(Sheet sheet, String[] names, XmlStyle style) {
        for (int i = 0; i < names.length; i++) {
          names[i] = createCellName(i);
        }
      }
    },

    HEADER_ROW {
      @Override
      void applyNames(Sheet sheet, String[] names, XmlStyle style) {
        Row headerRow = sheet.getRow(style.headerRow() - 1);
        for (int i = 0; i < names.length; i++) {
          Cell cell = headerRow.getCell(i);
          names[i] = safeName(cell.getRichStringCellValue().getString());
        }
      }
    };

    public String[] createColumnNames(Sheet sheet, XmlStyle style) {
      int nCells = getCellCount(sheet);
      String[] columnNames = new String[nCells];
      applyNames(sheet, columnNames, style);
      return columnNames;
    }

    abstract void applyNames(Sheet sheet, String[] placeholders, XmlStyle style);
  }

  public XmlStyle() {
    setElementNamingStyle(ElementNaming.SIMPLE.name());
    setXmlEncoding(null);
    setDateFormat("yyyy-MM-dd'T'HH:mm:ssZ");
  }

  public boolean getEmitDataTypeAttr() {
    return emitDataTypeAttr;
  }


  /**
   * Whether or not to emit a type attribute for each cell.
   *
   * @param b true or false, default false.
   */
  public void setEmitDataTypeAttr(boolean b) {
    emitDataTypeAttr = b;
  }

  public boolean getEmitRowNumberAttr() {
    return emitRowNumberAttr;
  }

  /**
   * Whether or not to emit a row number attribute for each row.
   *
   * @param b true or false, default false
   */
  public void setEmitRowNumberAttr(boolean b) {
    emitRowNumberAttr = b;
  }

  public boolean getEmitCellPositionAttr() {
    return emitCellPositionAttr;
  }

  /**
   * Whether or not to emit an absolute cell position attribute for each cell in a row.
   *
   * @param b true or false, default false
   */
  public void setEmitCellPositionAttr(boolean b) {
    emitCellPositionAttr = b;
  }

  public String getDateFormat() {
    return dateFormat;
  }

  /**
   * Set the date format for date fields.
   *
   * @param dateFormat the date format, default is "yyyy-MM-dd'T'HH:mm:ssZ"
   */
  public void setDateFormat(String dateFormat) {
    this.dateFormat = dateFormat;
  }


  protected String format(Date d) {
    if (dateFormatter == null) {
      dateFormatter = new SimpleDateFormat(getDateFormat());
    }
    return dateFormatter.format(d);
  }

  protected ElementNaming resolveNamingStrategy() {
    ElementNaming result = ElementNaming.SIMPLE;
    for (ElementNaming ns : ElementNaming.values()) {
      if (ns.name().equalsIgnoreCase(getElementNamingStyle())) {
        result = ns;
        break;
      }
    }
    return result;
  }

  public String getElementNamingStyle() {
    return elementNamingStyle;
  }

  /**
   * Set how element names are generated
   *
   * @param s the style; one of SIMPLE, CELL_POSITION, HEADER_ROW. Default is 'null' which is equivalent to SIMPLE
   * @see ElementNaming
   */
  public void setElementNamingStyle(String s) {
    elementNamingStyle = s;
  }

  public String getXmlEncoding() {
    return xmlEncoding;
  }

  /**
   * Set the XML Encoding for the document.
   *
   * @param encoding the encoding; default is null which implies <code>System.getProperty("file.encoding")</code>
   */
  public void setXmlEncoding(String encoding) {
    xmlEncoding = encoding;
  }

  protected String xmlEncoding() {
    return getXmlEncoding() != null ? getXmlEncoding() : System.getProperty("file.encoding");
  }

  public String getNumberFormat() {
    return numberFormat;
  }

  /**
   * Set the format for numeric fields
   *
   * @param s the format; default is null, which means to use {@link String#valueOf(double)}
   */
  public void setNumberFormat(String s) {
    numberFormat = s;
  }

  protected String format(double d) {
    if (Double.isNaN(d)) {
      throw new IllegalArgumentException("Not a double");
    }
    String result = String.valueOf(d);
    if (!isEmpty(getNumberFormat())) {
      if (numberFormatter == null) {
        numberFormatter = new DecimalFormat(getNumberFormat());
      }
      result = numberFormatter.format(d);
    }
    return result;
  }

  public Integer getHeaderRow() {
    return headerRow;
  }

  /**
   * If element name generation style is {@link ElementNaming#HEADER_ROW} then use this to specify which row is considered the
   * header.
   * <p>
   * If you specify a header row, then all rows preceding the header row will be skipped.
   * </p>
   *
   * @param i the header row (starts from 1); default is 1.
   */
  public void setHeaderRow(Integer i) {
    headerRow = i;
  }

  protected int headerRow() {
    return getHeaderRow() != null ? getHeaderRow().intValue() : 1;
  }

}
