package com.adaptris.core.poi;

import static org.apache.commons.lang.StringUtils.isEmpty;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.adaptris.util.GuidGenerator;

/**
 * Various helper methods for handling the horrible spreadsheet format.
 * 
 * @author lchan
 * 
 */
public class ExcelHelper {

  static final String XML_ATTR_POSITION = "position";

  static final String XML_ATTR_TYPE = "type";
  static final String XML_ATTR_TYPE_ERROR = "error";
  static final String XML_ATTR_TYPE_BOOLEAN = "boolean";
  static final String XML_ATTR_TYPE_STRING = "string";
  static final String XML_ATTR_TYPE_NUMERIC = "numeric";
  static final String XML_ATTR_TYPE_DATE = "date";
  static final String XML_ATTR_TYPE_FORMULA = "formula";
  static final String XML_ATTR_TYPE_BLANK = "blank";

  static final String XML_ATTR_ROW_NUMBER = "number";
  static final String XML_ATTR_SHEET_NAME = "name";
  static final String XML_ELEMENT_ROW = "row";
  static final String XML_ELEMENT_WORKSHEET = "sheet";
  static final String XML_ELEMENT_SPREADSHEET = "spreadsheet";

  private enum ColumnCells {
    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z
  }

  private static final String[] INVALID_CHARS =
  {
      "\\\\", "\\?", "\\*", "\\:", " ", "\\|", "&", "\\\"", "\\'", "<", ">", "\\)", "\\(", "\\/"
  };
  private static final String REPLACEMENT_VALUE = "_";

  private static final GuidGenerator guid = new GuidGenerator();

  // Obtuse use of enums as both an interface and factory...
  // Because the type handlers are very small classes, there's not much
  // point having an abstract class / interface and a bunch of separate implementations.
  public enum CellHandler {
    NUMERIC_CELL(XML_ATTR_TYPE_NUMERIC, Cell.CELL_TYPE_NUMERIC) {
      @Override
      public CellHandler getHandler(Cell cell) {
        if (myCellType == cell.getCellType()) {
          if (DateUtil.isCellDateFormatted(cell)) {
            return DATE_CELL;
          }
          return this;
        }
        return null;
      }

      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return style.format(cell.getNumericCellValue());
      }
    },
    FORMULA_CELL(XML_ATTR_TYPE_FORMULA, Cell.CELL_TYPE_FORMULA) {
      @Override
      public CellHandler getHandler(Cell cell) {
        if (myCellType == cell.getCellType()) {
          if (DateUtil.isCellDateFormatted(cell)) {
            return DATE_CELL;
          }
          return this;
        }
        return null;
      }

      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return style.format(cell.getNumericCellValue());
      }
    },
    // Use -1 to represent the type as it isn't really a type, it's just formatting type.
    DATE_CELL(XML_ATTR_TYPE_DATE, -1) {
      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return style.format(DateUtil.getJavaDate(cell.getNumericCellValue()));
      }
    },
    STRING_CELL(XML_ATTR_TYPE_STRING, Cell.CELL_TYPE_STRING) {
      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return cell.getRichStringCellValue().getString();
      }

    },
    BOOLEAN_CELL(XML_ATTR_TYPE_BOOLEAN, Cell.CELL_TYPE_BOOLEAN) {

      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return String.valueOf(cell.getBooleanCellValue());
      }

    },
    ERROR_CELL(XML_ATTR_TYPE_ERROR, Cell.CELL_TYPE_ERROR) {
      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return String.valueOf(cell.getErrorCellValue());
      }
    },
    BLANK_CELL(XML_ATTR_TYPE_BLANK, Cell.CELL_TYPE_BLANK) {
      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return "";
      }
    };

    String myType;
    int myCellType;

    CellHandler(String type, int cellType) {
      myType = type;
      myCellType = cellType;
    }

    public CellHandler getHandler(Cell cell) {
      if (myCellType == cell.getCellType()) {
        return this;
      }
      return null;
    }

    public abstract String getValue(Cell cell, XmlStyle style);

    public String getType() {
      return myType;
    }
  };

  protected static CellHandler getHandler(Cell cell) throws Exception {
    if (cell == null) {
      return CellHandler.BLANK_CELL;
    }
    CellHandler handler = null;
    for (CellHandler ch : CellHandler.values()) {
      CellHandler t = ch.getHandler(cell);
      if (t != null) {
        handler = t;
        break;
      }
    }
    if (handler == null) {
      throw new Exception("Couldn't find a handler for a cellType of " + cell.getCellType());
    }
    return handler;
  }

  protected static int getRowCount(Sheet sheet) {
    int result = sheet.getLastRowNum();
    // 0 means 0 rows on the sheet, or one at position zero.
    if (result == 0) {
      result = sheet.getPhysicalNumberOfRows();
    }
    else {
      result = result + 1;
    }
    return result;
  }

  protected static int getCellCount(Sheet sheet) {
    int result = 0;
    int highRow = getRowCount(sheet);
    for (int i = sheet.getFirstRowNum(); i < highRow; i++) {
      Row row = sheet.getRow(i);
      if (row != null && row.getPhysicalNumberOfCells() > result) {
        result = row.getLastCellNum();
      }
    }
    return result;
  }

  protected static String createCellName(final int cellNumber) {
    String cellName = "";
    int len = ColumnCells.values().length;
    int cn = cellNumber;
    // Handles the situation where you have more than 26 columns (see sample-input.xls) which will
    // end up giving (hopefully) the names AA,AB,AC etc.
    int first = cn / len;
    if (first == 0) {
      cellName += ColumnCells.values()[cn];
    }
    else {
      cellName += ColumnCells.values()[first - 1];
      cellName += ColumnCells.values()[cn % len];
    }
    return cellName;
  }

  protected static String safeName(String input) {
    String name = input;
    for (String invalid : INVALID_CHARS) {
      name = name.replaceAll(invalid, REPLACEMENT_VALUE);
    }
    if (isEmpty(name)) {
      name = "blank_" + guid.safeUUID();
    }
    return name;
  }
}
