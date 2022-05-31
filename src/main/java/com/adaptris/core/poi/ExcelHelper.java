package com.adaptris.core.poi;

import com.adaptris.core.util.Args;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.adaptris.core.util.XmlHelper;
import com.adaptris.util.GuidGenerator;

/**
 * Various helper methods for handling the horrible spreadsheet format.
 *
 * @author lchan
 *
 */
class ExcelHelper {

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

  static final int LETTERS_IN_ALPHABET = 26;

  private static final String[] INVALID_CHARS =
  {
      "\\\\", "\\?", "\\*", "\\:", " ", "\\|", "&", "\\\"", "\\'", "<", ">", "\\)", "\\(", "\\/"
  };
  private static final String REPLACEMENT_VALUE = "_";

  private static final GuidGenerator guid = new GuidGenerator();

  // Obtuse use of enums as both an interface and factory...
  // Because the type handlers are very small classes, there's not much
  // point having an abstract class / interface and a bunch of separate implementations.
  enum CellHandler {
    NUMERIC_CELL(XML_ATTR_TYPE_NUMERIC, CellType.NUMERIC) {
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
    FORMULA_CELL(XML_ATTR_TYPE_FORMULA, CellType.FORMULA) {
      @Override
      public CellHandler getHandler(Cell cell) {
        if (myCellType == cell.getCellType()) {
          try {
            if (DateUtil.isCellDateFormatted(cell)) {
              return DATE_CELL;
            }
          }
          catch (Exception e) {
            // Expect this error if the cell contains an invalid formula
          }
          return this;
        }
        return null;
      }

      @Override
      public String getValue(Cell cell, XmlStyle style) {
        try {
          return style.format(cell.getNumericCellValue());
        }
        catch (Exception e) {
          // Expect this error if the cell contains formula that doesn't create number
          try {
            return cell.getRichStringCellValue().getString();
          }
          catch (Exception e1) {
            // Expect this error if the cell contains an invalid formula
            String errorString = FormulaError.forInt(cell.getErrorCellValue()).getString();
            return errorString;
          }
        }
      }
    },
    // Use -1 to represent the type as it isn't really a type, it's just formatting type.
    DATE_CELL(XML_ATTR_TYPE_DATE, null) {
      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return style.format(DateUtil.getJavaDate(cell.getNumericCellValue()));
      }
    },
    STRING_CELL(XML_ATTR_TYPE_STRING, CellType.STRING) {
      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return cell.getRichStringCellValue().getString();
      }

    },
    BOOLEAN_CELL(XML_ATTR_TYPE_BOOLEAN, CellType.BOOLEAN) {

      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return String.valueOf(cell.getBooleanCellValue());
      }

    },
    ERROR_CELL(XML_ATTR_TYPE_ERROR, CellType.ERROR) {
      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return String.valueOf(cell.getErrorCellValue());
      }
    },
    BLANK_CELL(XML_ATTR_TYPE_BLANK, CellType.BLANK) {
      @Override
      public String getValue(Cell cell, XmlStyle style) {
        return "";
      }
    };

    String myType;
    CellType myCellType;

    CellHandler(String type, CellType cellType) {
      myType = type;
      myCellType = cellType;
    }

    public CellHandler getHandler(Cell cell) {
      if (cell.getCellType() == myCellType) {
        return this;
      }
      return null;
    }

    public abstract String getValue(Cell cell, XmlStyle style);

    public String getType() {
      return myType;
    }
  };

  static CellHandler getHandler(Cell cell) throws Exception {
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

  static int getRowCount(Sheet sheet) {
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

  static int getCellCount(Sheet sheet) {
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

  /** Create a standardised Excel column name based on the colNumber.
   *
   *  <p>Note that Excel has a limit of 16384 columns, this will cope with more columns than that, but accuracy
   *  isn't guaranteed after column 16383.
   *  </p>
   * @param colNumber the position of the column, zero-indexed (i.e. 0 == 'A')
   * @return the Excel column Name.
   */
  static String createCellName(final int colNumber) {
    // This is a bit of a fudge because cell numbers from poi
    // are zero indexed.
    int cell = colNumber + 1;
    StringBuilder cellName = new StringBuilder();
    int asciiIndex = 0;
    while (cell > 0) {
      asciiIndex = (cell - 1) % LETTERS_IN_ALPHABET;
      cellName.append((char) (asciiIndex + 'A'));
      cell = (int) ((cell - asciiIndex) / LETTERS_IN_ALPHABET);
    }
    return cellName.reverse().toString();
  }

  /** Return the computed cell name based on normal Excel Naming.
   *
   * @param columnName the name of the cell as showin in Excel
   * @return the positional index that logically is (zero index, i.e. 0 == 'A')
   */
  static int numericColumnName(String columnName) {
    int result = 0;
    for (int i = 0; i < Args.notBlank(columnName, "columnName").length(); i++) {
      result *= LETTERS_IN_ALPHABET;
      result += columnName.charAt(i) - 'A' + 1;
    }
    return result - 1;
  }

  static String safeName(String input) {
    return XmlHelper.safeElementName(input, "blank_" + guid.safeUUID());
  }
}
