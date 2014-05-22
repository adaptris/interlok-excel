package com.adaptris.core.poi;

import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.IOUtils;
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
 * <p>
 * Requires a Standard License
 * </p>
 * 
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
    HSSFWorkbook workbook = null;
    InputStream in = null;
    OutputStream out = null;
    ExcelConverter converter = new ExcelConverter();
    try {
      in = msg.getInputStream();
      out = msg.getOutputStream();
      workbook = new HSSFWorkbook(new POIFSFileSystem(in));
      Document d = converter.convertToXml(workbook, getXmlStyle());
      new XmlUtils().writeDocument(d, out, getXmlStyle().xmlEncoding());
      msg.setCharEncoding(getXmlStyle().xmlEncoding());
    }
    catch (Exception e) {
      throw new ServiceException(e);
    }
    finally {
      IOUtils.closeQuietly(in);
      IOUtils.closeQuietly(out);
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
