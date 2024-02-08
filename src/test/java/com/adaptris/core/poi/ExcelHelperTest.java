package com.adaptris.core.poi;

import static com.adaptris.core.poi.ExcelHelper.createCellName;
import static com.adaptris.core.poi.ExcelHelper.numericColumnName;
import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

public class ExcelHelperTest {

  @Test
  public void testCreateCellName() throws Exception {
    // Cells are zero-indexed.
    int len = ExcelHelper.LETTERS_IN_ALPHABET;
    assertEquals("A", createCellName(0));
    assertEquals("AB", createCellName(len + 1));
    assertEquals("ZZ", createCellName(len * len + 25));
    assertEquals("AAA", createCellName(len * len + 26));
    assertEquals("AAD", createCellName(len * len + 29));

    // Some well known columns.
    assertEquals("CV", createCellName(99));
    assertEquals("ALL", createCellName(999));
    assertEquals("NTP", createCellName(9999));
    // This is the logical max columns according to MS for an excel worksheet, max is 16384
    assertEquals("XFD", createCellName(16383));
  }

  @Test
  public void testNumericColumnName() throws Exception {
    assertEquals(0, numericColumnName("A"));
    assertEquals(27, numericColumnName("AB"));
    assertEquals(702, numericColumnName("ZZ") + 1);

    assertEquals(100, numericColumnName("CV") + 1);
    assertEquals(1000, numericColumnName("ALL") + 1);
    assertEquals(10000, numericColumnName("NTP") + 1);
    assertEquals(16384, numericColumnName("XFD") + 1);
  }

}
