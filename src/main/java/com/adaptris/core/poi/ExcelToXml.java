package com.adaptris.core.poi;

import java.io.InputStream;
import java.io.OutputStream;

import org.apache.commons.lang3.BooleanUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

import com.adaptris.annotation.AdapterComponent;
import com.adaptris.annotation.ComponentProfile;
import com.adaptris.core.AdaptrisMessage;
import com.adaptris.core.CoreException;
import com.adaptris.core.ServiceException;
import com.adaptris.core.ServiceImp;
import com.adaptris.core.util.ExceptionHelper;
import com.adaptris.util.XmlUtils;
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
@AdapterComponent
@ComponentProfile(summary = "Convert an Excel Spreadsheet to XML", tag = "service,transform,excel,xml")
public class ExcelToXml extends ServiceImp implements ExcelConverter.ExcelConverterContext {
  private XmlStyle xmlStyle;
  private Boolean ignoreNullRows;

  public ExcelToXml() {
    setXmlStyle(new XmlStyle());
  }


  @Override
  public void doService(AdaptrisMessage msg) throws ServiceException {
    Workbook workbook = null;
    ExcelConverter converter = new ExcelConverter(this);
    try (InputStream in = msg.getInputStream()) {
      workbook = WorkbookFactory.create(in);
      Document d = converter.convertToXml(workbook, getXmlStyle());
      writeXmlDocument(d, msg);
    }
    catch (Exception e) {
      throw ExceptionHelper.wrapServiceException(e);
    }
  }

  protected void writeXmlDocument(Document doc, AdaptrisMessage msg) throws Exception {
    try (OutputStream out = msg.getOutputStream()) {
      String encoding = getXmlStyle().evaluateEncoding(msg);
      new XmlUtils().writeDocument(doc, out, encoding);
      msg.setContentEncoding(encoding);
    }
  }

  @Override
  public void initService() throws CoreException {
  }

  @Override
  public void closeService() {
  }

  @Override
  public void prepare() throws CoreException {}
  
  public XmlStyle getXmlStyle() {
    return xmlStyle;
  }

  public void setXmlStyle(XmlStyle style) {
    xmlStyle = style;
  }

  public Boolean getIgnoreNullRows() {
    return ignoreNullRows;
  }

  /**
   * Set to true to ignore null rows.
   * <p>
   * In some spreadsheets it is possible to get a null Row object when doing {@code Row row = sheet.getRow(i)}. Set this to be true
   * to silently ignore errors, which means you may a mismatch between the number of rows in the spreadsheet vs the number of Row
   * elements in the resulting XML.
   * </p>
   * 
   * @param b true to ignore rows that are null; default null (false).
   * @since 3.0.3
   */
  public void setIgnoreNullRows(Boolean b) {
    this.ignoreNullRows = b;
  }
  
  public boolean ignoreNullRows() {
    return BooleanUtils.toBooleanDefaultIfNull(getIgnoreNullRows(), false);
  }

  public Logger logger() {
    return this.log;
  }


}
