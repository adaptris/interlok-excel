package com.adaptris.core.poi;

import static com.adaptris.core.poi.ExcelHelper.createCellName;
import static com.adaptris.core.poi.ExcelHelper.getCellCount;
import static com.adaptris.core.poi.ExcelHelper.safeName;
import static org.apache.commons.lang3.StringUtils.isEmpty;

import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.validation.constraints.NotNull;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.adaptris.annotation.AutoPopulated;
import com.adaptris.core.AdaptrisMessage;

/**
 * Controls the XML look and feel for ExcelToXml.
 *
 * @author lchan
 *
 */
public class XmlStyle {
  private Boolean emitDataTypeAttr;
  private Boolean emitRowNumberAttr;
  private Boolean emitCellPositionAttr;

  @NotNull
  @AutoPopulated
  private String dateFormat;
  private String numberFormat;
  private Integer headerRow;

  @NotNull
  private ElementNaming elementNamingStyle;
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
    setElementNamingStyle(ElementNaming.SIMPLE);
    setXmlEncoding(null);
    setDateFormat("yyyy-MM-dd'T'HH:mm:ssZ");
  }

  public Boolean getEmitDataTypeAttr() {
    return emitDataTypeAttr;
  }

  /**
   * Whether or not to emit a type attribute for each cell.
   *
   * @param b
   *          true or false, default false.
   */
  public void setEmitDataTypeAttr(Boolean b) {
    emitDataTypeAttr = b;
  }

  boolean emitDataTypeAttr() {
    return getEmitDataTypeAttr() != null ? getEmitDataTypeAttr().booleanValue() : false;
  }

  public Boolean getEmitRowNumberAttr() {
    return emitRowNumberAttr;
  }

  /**
   * Whether or not to emit a row number attribute for each row.
   *
   * @param b
   *          true or false, default false
   */
  public void setEmitRowNumberAttr(Boolean b) {
    emitRowNumberAttr = b;
  }

  boolean emitRowNumberAttr() {
    return getEmitRowNumberAttr() != null ? getEmitRowNumberAttr().booleanValue() : false;
  }

  public Boolean getEmitCellPositionAttr() {
    return emitCellPositionAttr;
  }

  /**
   * Whether or not to emit an absolute cell position attribute for each cell in a row.
   *
   * @param b
   *          true or false, default false
   */
  public void setEmitCellPositionAttr(Boolean b) {
    emitCellPositionAttr = b;
  }

  boolean emitCellPositionAttr() {
    return getEmitCellPositionAttr() != null ? getEmitCellPositionAttr().booleanValue() : false;
  }

  public String getDateFormat() {
    return dateFormat;
  }

  /**
   * Set the date format for date fields.
   *
   * @param dateFormat
   *          the date format, default is "yyyy-MM-dd'T'HH:mm:ssZ"
   */
  public void setDateFormat(String dateFormat) {
    this.dateFormat = dateFormat;
  }

  String format(Date d) {
    if (dateFormatter == null) {
      dateFormatter = new SimpleDateFormat(getDateFormat());
    }
    return dateFormatter.format(d);
  }

  ElementNaming resolveNamingStrategy() {
    return getElementNamingStyle() != null ? getElementNamingStyle() : ElementNaming.SIMPLE;
  }

  public ElementNaming getElementNamingStyle() {
    return elementNamingStyle;
  }

  /**
   * Set how element names are generated
   *
   * @param s
   *          the style; one of SIMPLE, CELL_POSITION, HEADER_ROW. Default is SIMPLE
   * @see ElementNaming
   */
  public void setElementNamingStyle(ElementNaming s) {
    elementNamingStyle = s;
  }

  public String getXmlEncoding() {
    return xmlEncoding;
  }

  /**
   * Set the encoding for the resulting XML document.
   * <p>
   * If not specified the following rules will be applied:
   * </p>
   * <ol>
   * <li>If the {@link AdaptrisMessage#getContentEncoding()} is non-null then that will be used.</li>
   * <li>UTF-8</li>
   * </ol>
   * <p>
   * As a result; the character encoding on the message is always set using {@link AdaptrisMessage#setContentEncoding(String)}.
   * </p>
   */
  public void setXmlEncoding(String encoding) {
    xmlEncoding = encoding;
  }

  String evaluateEncoding(AdaptrisMessage msg) {
    String encoding = "UTF-8";
    if (!isEmpty(getXmlEncoding())) {
      encoding = getXmlEncoding();
    } else if (!isEmpty(msg.getContentEncoding())) {
      encoding = msg.getContentEncoding();
    }
    return encoding;
  }

  public String getNumberFormat() {
    return numberFormat;
  }

  /**
   * Set the format for numeric fields
   *
   * @see DecimalFormat
   * @param s
   *          the format; default is null, which means to use {@link String#valueOf(double)}
   */
  public void setNumberFormat(String s) {
    numberFormat = s;
  }

  String format(double d) {
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
   * If element name generation style is {@link ElementNaming#HEADER_ROW} then use this to specify which row is considered the header.
   * <p>
   * If you specify a header row, then all rows preceding the header row will be skipped.
   * </p>
   *
   * @param i
   *          the header row (starts from 1); default is 1.
   */
  public void setHeaderRow(Integer i) {
    headerRow = i;
  }

  protected int headerRow() {
    return getHeaderRow() != null ? getHeaderRow().intValue() : 1;
  }

}
