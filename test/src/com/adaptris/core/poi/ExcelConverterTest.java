package com.adaptris.core.poi;

import static org.apache.commons.lang.StringUtils.isEmpty;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.IOUtils;
import org.w3c.dom.Document;

import com.adaptris.core.BaseCase;
import com.adaptris.util.XmlUtils;
import com.adaptris.util.text.xml.XPath;

public class ExcelConverterTest extends BaseCase {
  public static final String TMP_DIR_KEY = "tmp.dir";
  public static final String KEY_SAMPLE_INPUT = "poi.sample.input";
  public static final String KEY_SAMPLE_INPUT_WITH_HEADER = "poi.sample.input.header";
  
  public ExcelConverterTest(String name) {
    super(name);
  }

  @Override
  public void setUp() throws Exception {
    super.setUp();
  }

  @Override
  public void tearDown() throws Exception {
    super.tearDown();
  }

  public void testDefaultConversion() throws Exception {
    ExcelConverter c = new ExcelConverter();
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(in));
      Document d = c.convertToXml(workbook, new XmlStyle());
      assertNotNull(d);
      new XmlUtils().writeDocument(d, System.err);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell").getLength() > 0);
      assertNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell[@position='A1']"));
    }
    finally {
      IOUtils.closeQuietly(in);
    }
  }


  public void testSimpleConversion_emitRow() throws Exception {
    ExcelConverter c = new ExcelConverter();
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(in));
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.SIMPLE.name());
      style.setEmitRowNumberAttr(true);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell").getLength() > 0);
      assertNotNull(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row[@number='1']/cell"));
    }
    finally {
      IOUtils.closeQuietly(in);
    }
  }

  public void testSimpleConversion_emitAllAttributes_WithNumberFormat() throws Exception {
    ExcelConverter c = new ExcelConverter();
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(in));
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
    }
    finally {
      IOUtils.closeQuietly(in);
    }
  }

  public void testSimpleConversion_emitAllAttributes() throws Exception {
    ExcelConverter c = new ExcelConverter();
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(in));
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
      assertNotNull(xp.selectSingleTextItem(d,
          "/spreadsheet/sheet[@name='Sheet1']/row[@number='1']/cell[@position='A1' and @type='string']"));
      assertTrue(isEmpty(xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row[@number='1']/cell[@position='A2']")));
    }
    finally {
      IOUtils.closeQuietly(in);
    }
  }

  public void testPositionalConversion() throws Exception {
    ExcelConverter c = new ExcelConverter();
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(in));
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.CELL_POSITION.name());
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/A").getLength() > 0);
    }
    finally {
      IOUtils.closeQuietly(in);
    }
  }

  public void testPositionalConversion_emitAllAttributes() throws Exception {
    ExcelConverter c = new ExcelConverter();
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(in));
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.CELL_POSITION.name());
      style.setEmitCellPositionAttr(true);
      style.setEmitDataTypeAttr(true);
      style.setEmitRowNumberAttr(true);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertNotNull(xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row/AA[@position='AA2']"));
      assertNotNull(xp.selectSingleTextItem(d,
          "/spreadsheet/sheet[@name='Sheet1']/row[@number='1']/AA[@position='AA1' and @type='string']"));
      assertTrue(isEmpty(xp.selectSingleTextItem(d, "/spreadsheet/sheet[@name='Sheet1']/row[@number='1']/AA[@position='A2']")));

    }
    finally {
      IOUtils.closeQuietly(in);
    }
  }

  public void testHeaderRowConversion() throws Exception {
    ExcelConverter c = new ExcelConverter();
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(in));
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.HEADER_ROW.name());
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/Column_Zee").getLength() > 0);
    }
    finally {
      IOUtils.closeQuietly(in);
    }
  }

  public void testHeaderRowConversion_emitAllAttributes() throws Exception {
    ExcelConverter c = new ExcelConverter();
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT);
    try {
      HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(in));
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.HEADER_ROW.name());
      style.setEmitCellPositionAttr(true);
      style.setEmitDataTypeAttr(true);
      style.setEmitRowNumberAttr(true);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/Column_Zee").getLength() > 0);
      assertNotNull(xp.selectSingleTextItem(d,
          "/spreadsheet/sheet[@name='Sheet1']/row[@number='2']/Column_Zee[@position='ZZ1' and @type='date']"));
      assertTrue(isEmpty(xp.selectSingleTextItem(d,
          "/spreadsheet/sheet[@name='Sheet1']/row[@number='2']/Column_Zee[@position='ZZ1' and @type='string']")));
    }
    finally {
      IOUtils.closeQuietly(in);
    }
  }

  public void testHeaderRowConversion_Offset() throws Exception {
    ExcelConverter c = new ExcelConverter();
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT_WITH_HEADER);
    try {
      HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(in));
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.HEADER_ROW.name());
      style.setHeaderRow(5);
      Document d = c.convertToXml(workbook, style);
      assertNotNull(d);
      XPath xp = new XPath();
      assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Source Data']"));
      assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Source Data']/row/Project_Name").getLength() > 0);
    }
    finally {
      IOUtils.closeQuietly(in);
    }
  }

  public void testHeaderRowConversion_Offset_EmitAllAttributes() throws Exception {
    ExcelConverter c = new ExcelConverter();
    InputStream in = createFromProperty(KEY_SAMPLE_INPUT_WITH_HEADER);
    try {
      HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(in));
      XmlStyle style = new XmlStyle();
      style.setElementNamingStyle(XmlStyle.ElementNaming.HEADER_ROW.name());
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
    }
    finally {
      IOUtils.closeQuietly(in);
    }
  }

  protected static InputStream createFromProperty(String key) throws IOException {
    File f = new File(PROPERTIES.getProperty(key));

    return new FileInputStream(f);
  }

}
