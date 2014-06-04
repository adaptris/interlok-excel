package com.adaptris.core.poi;

import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.Document;

import com.adaptris.core.AdaptrisMessage;
import com.adaptris.core.CoreException;
import com.adaptris.core.ServiceException;
import com.adaptris.core.ServiceImp;
import com.adaptris.util.XmlUtils;
import com.adaptris.util.license.License;
import com.adaptris.util.license.License.LicenseType;
import com.thoughtworks.xstream.annotations.XStreamAlias;


/**
 * Service to extract data from an Excel spreadsheet.
 * 
 * @config excel-to-xml-service
 * @license STANDARD
 * @author lchan
 * 
 */
@XStreamAlias("excel-to-xml-service")
public class ExcelToXml extends ServiceImp {
  private XmlStyle xmlStyle;

  public ExcelToXml() {
    setXmlStyle(new XmlStyle());
  }


  @Override
  public void doService(AdaptrisMessage msg) throws ServiceException {
    Workbook workbook = null;
    ExcelConverter converter = new ExcelConverter();
    try (InputStream in = msg.getInputStream()) {
      workbook = WorkbookFactory.create(in);
      Document d = converter.convertToXml(workbook, getXmlStyle());
      writeXmlDocument(d, msg);
    }
    catch (Exception e) {
      throw new ServiceException(e);
    }
  }

  protected void writeXmlDocument(Document doc, AdaptrisMessage msg) throws Exception {
    try (OutputStream out = msg.getOutputStream()) {
      String encoding = getXmlStyle().evaluateEncoding(msg);
      new XmlUtils().writeDocument(doc, out, encoding);
      msg.setCharEncoding(encoding);
    }
  }

  @Override
  public void init() throws CoreException {
  }

  @Override
  public void close() {
  }

  @Override
  public void start() throws CoreException {
    super.start();
  }

  @Override
  public void stop() {
    super.stop();
  }

  public XmlStyle getXmlStyle() {
    return xmlStyle;
  }

  public void setXmlStyle(XmlStyle style) {
    xmlStyle = style;
  }

  @Override
  public boolean isEnabled(License license) throws CoreException {
    return license.isEnabled(LicenseType.Standard);
  }
  

}
