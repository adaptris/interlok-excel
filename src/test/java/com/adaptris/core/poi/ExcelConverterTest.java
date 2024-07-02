package com.adaptris.core.poi;

import static org.apache.commons.lang3.StringUtils.isEmpty;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.fail;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.junit.jupiter.api.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

import com.adaptris.interlok.junit.scaffolding.BaseCase;
import com.adaptris.util.XmlUtils;
import com.adaptris.util.text.xml.XPath;

public class ExcelConverterTest extends BaseCase {
  public static final String TMP_DIR_KEY = "tmp.dir";
  public static final String KEY_SAMPLE_INPUT = "poi.sample.input";
  public static final String KEY_SAMPLE_INPUT_WITH_NULL = "poi.sample.input.null";
  public static final String KEY_SAMPLE_INPUT_WITH_HEADER = "poi.sample.input.header";
  public static final String KEY_SAMPLE_XLSX_INPUT_WITH_HEADER = "poi.sample.xlsx.header";

  @Test
  public void testDefaultConversion() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      Document d = c.convertToXml(workbook, new XmlStyle());
      assertNotNull(d);
      new XmlUtils().writeDocument(d, System.err);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell").getLength() > 0);
      assertNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell[@position='A1']"));
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testConversion_NullRow_IgnoreNullRowFalse() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT_WITH_NULL);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      c.convertToXml(workbook, new XmlStyle());
      fail("Conversion Succeeded with Null Row");
    } catch (Exception e) {
      assertTrue(e.getMessage().matches("Unable to process.*it's null"));
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testConversion_NullRow_IgnoreNullRow() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(true));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT_WITH_NULL);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      Document d = c.convertToXml(workbook, new XmlStyle());
      assertNotNull(d);
      new XmlUtils().writeDocument(d, System.err);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell").getLength() > 0);
      assertNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell[@position='A1']"));
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testSimpleConversion_emitRow() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.SIMPLE);
      style.setEmitRowNumberAttr(true);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell").getLength() > 0);
      assertNotNull(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row[@number='1']/cell"));
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testSimpleConversion_emitAllAttributes_WithNumberFormat() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      XmlStyle style = new XmlStyle();
      style.setEmitCellPositionAttr(true);
      style.setEmitDataTypeAttr(true);
      style.setEmitRowNumberAttr(true);
      style.setNumberFormat("0.###E0");
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell").getLength() > 0);
      String value = xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row[@number='3']/cell[@position='A3']");
      // Engineering notation!
      assertEquals("1E0", value);
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testSimpleConversion_emitAllAttributes() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      XmlStyle style = new XmlStyle();
      style.setEmitCellPositionAttr(true);
      style.setEmitDataTypeAttr(true);
      style.setEmitRowNumberAttr(true);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell").getLength() > 0);
      assertNotNull(xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell[@position='A1']"));
      assertNotNull(
          xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row[@number='1']/cell[@position='A1' and @type='string']"));
      assertTrue(isEmpty(xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row[@number='1']/cell[@position='A2']")));
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testPositionalConversion() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.CELL_POSITION);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/A").getLength() > 0);
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testPositionalConversion_emitAllAttributes() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.CELL_POSITION);
      style.setEmitCellPositionAttr(true);
      style.setEmitDataTypeAttr(true);
      style.setEmitRowNumberAttr(true);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertNotNull(xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row/AA[@position='AA2']"));
      assertNotNull(
          xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row[@number='1']/AA[@position='AA1' and @type='string']"));
      assertTrue(isEmpty(xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row[@number='1']/AA[@position='A2']")));

    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testHeaderRowConversion_XLSX() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_XLSX_INPUT_WITH_HEADER);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.HEADER_ROW);
      style.setHeaderRow(5);
      style.setEmitCellPositionAttr(true);
      style.setEmitDataTypeAttr(true);
      style.setEmitRowNumberAttr(true);

      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Source Data']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Source Data']/row/Project_Name").getLength() > 0);
      // This checks that formulas are being calculated.
      assertEquals("12300.0", xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Source Data']/row[@number='6']/Estimated_Days"));
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testHeaderRowConversion() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.HEADER_ROW);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/Column_Zee").getLength() > 0);
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testHeaderRowConversion_emitAllAttributes() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.HEADER_ROW);
      style.setEmitCellPositionAttr(true);
      style.setEmitDataTypeAttr(true);
      style.setEmitRowNumberAttr(true);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/Column_Zee").getLength() > 0);
      assertNotNull(
          xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row[@number='2']/Column_Zee[@position='ZZ1' and @type='date']"));
      assertTrue(isEmpty(xp.selectSingleTextItem(d,
          "/spreadsheet/sheet[@name='Sheet1']/row[@number='2']/Column_Zee[@position='ZZ1' and @type='string']")));
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testHeaderRowConversion_Offset() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT_WITH_HEADER);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.HEADER_ROW);
      style.setHeaderRow(5);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Source Data']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Source Data']/row/Project_Name").getLength() > 0);
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  @Test
  public void testHeaderRowConversion_Offset_EmitAllAttributes() throws Exception {
    ExcelConverter c = new ExcelConverter(new BasicExcelConverterContext(false));
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT_WITH_HEADER);
    try {
      Workbook workbook = WorkbookFactory.create(in);
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.HEADER_ROW);
      style.setHeaderRow(5);
      style.setEmitCellPositionAttr(true);
      style.setEmitDataTypeAttr(true);
      style.setEmitRowNumberAttr(true);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Source Data']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Source Data']/row/Project_Name").getLength() > 0);
      String projectName = xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Source Data']/row[@number='6']/Project_Name");
      assertEquals("Cyberdyne Systems", projectName);
      String projectType = xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Source Data']/row/Project_Type[@position='B11']");
      assertEquals("The Lazarus Project", projectType);
    } finally {
      IOUtils.closeQuietly(in);
    }
  }

  protected static InputStream createFromProperty(String key) throws IOException {
    File f = new File(PROPERTIES.getProperty(key));

    return new FileInputStream(f);
  }

  private class BasicExcelConverterContext implements ExcelConverter.ExcelConverterContext {

    private transient boolean ignoreNulls;

    private BasicExcelConverterContext(boolean b) {
      ignoreNulls = b;
    }

    @Override
    public boolean ignoreNullRows() {
      return ignoreNulls;
    }

    @Override
    public Logger logger() {
      return LoggerFactory.getLogger(ExcelConverterTest.class);
    }

  }

}
