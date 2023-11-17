package com.adaptris.core.poi;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.xml.xpath.XPathExpressionException;

import org.apache.commons.io.IOUtils;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;

import com.adaptris.core.AdaptrisMessage;
import com.adaptris.core.DefaultMessageFactory;
import com.adaptris.core.ServiceException;
import com.adaptris.core.util.DocumentBuilderFactoryBuilder;
import com.adaptris.core.util.XmlHelper;
import com.adaptris.interlok.junit.scaffolding.services.ExampleServiceCase;
import com.adaptris.util.text.xml.XPath;

public class PoiServiceTest extends ExampleServiceCase {
  public static final String KEY_SAMPLE_INPUT = "poi.sample.input";
  public static final String KEY_SAMPLE_INPUT_WITH_NULL = "poi.sample.input.null";
  private DefaultMessageFactory dMessageFactory = new DefaultMessageFactory();

  public PoiServiceTest() {
    super();
  }

  @Override
  protected Object retrieveObjectForSampleConfig() {
    return new ExcelToXml();
  }

  @Override
  protected Object retrieveObjectForCastorRoundTrip() {
    return new ExcelToXml();
  }

  protected static byte[] readFile(String path) throws IOException {
    try (ByteArrayOutputStream out = new ByteArrayOutputStream(); InputStream in = new FileInputStream(new File(path))){
      IOUtils.copy(in, out);
      return out.toByteArray();
    }
  }

  @Test
  public void testService() throws Exception {
    AdaptrisMessage msg = dMessageFactory.newMessage(readFile(PROPERTIES.getProperty(KEY_SAMPLE_INPUT)));
    ExcelToXml service = new ExcelToXml();
    service.getXmlStyle().setXmlEncoding("UTF-8");
    try {
      start(service);
      service.doService(msg);
      Document d = XmlHelper.createDocument(msg, DocumentBuilderFactoryBuilder.newInstance());
      assertDocument(d);
    }
    finally {
      stop(service);
    }
  }

  @Test
  public void testServiceIgnoreNullRowsFalse() throws Exception {
    AdaptrisMessage msg = dMessageFactory.newMessage(readFile(PROPERTIES.getProperty(KEY_SAMPLE_INPUT_WITH_NULL)));
    ExcelToXml service = new ExcelToXml();
    service.getXmlStyle().setXmlEncoding("UTF-8");
    try {
      start(service);
      ServiceException serviceException = assertThrows(ServiceException.class, () -> service.doService(msg));
      assertEquals(Exception.class, serviceException.getCause().getClass());
      assertEquals("Unable to process row 5; it's null", serviceException.getCause().getMessage());
    } finally {
      stop(service);
    }
  }

  @Test
  public void testServiceIgnoreNullRowsTrue() throws Exception {
    AdaptrisMessage msg = dMessageFactory.newMessage(readFile(PROPERTIES.getProperty(KEY_SAMPLE_INPUT_WITH_NULL)));
    ExcelToXml service = new ExcelToXml();
    service.getXmlStyle().setXmlEncoding("UTF-8");
    service.setIgnoreNullRows(true);
    try {
      start(service);
      service.doService(msg);
      Document d = XmlHelper.createDocument(msg, DocumentBuilderFactoryBuilder.newInstance());
      assertDocument(d);
    } finally {
      stop(service);
    }
  }

  private void assertDocument(Document d) throws XPathExpressionException {
    XPath xp = new XPath();
    assertNotNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']"));
    assertTrue(xp.selectNodeList(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell").getLength() > 0);
    assertNull(xp.selectSingleNode(d, "/spreadsheet/sheet[@name='Sheet1']/row/cell[@position='A1']"));
  }

}