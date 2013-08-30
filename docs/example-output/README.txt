All output is from tests/src/resources/sample-input.xls unless otherwise specified.

Each output document is named after the XML-Style that is configured.
i.e. PositionalStyle-AllAttributes.xml would indicate an xml-style of
              <xml-style>
                <element-naming-style>CELL_POSITION</element-naming-style>
                <emit-data-type-attr>true</emit-data-type-attr>
                <emit-row-number-attr>true</emit-row-number-attr>
                <emit-cell-position-attr>true</emit-cell-position-attr>
              </xml-style>

And so on and so forth; 
SimpleStyle == <element-naming-style>SIMPLE</element-naming-style>
PositionalStyle = <element-naming-style>CELL_POSITION</element-naming-style>
HeaderRow = <element-naming-style>HEADER_ROW</element-naming-style>

HeaderRow-Offest-* are sourced from resources/tests/test-input-header-row.xls with a specific configuration
that includes a header-row which starts processing of the XLS at the point of the header-row
ignoring all rows prior to that row.

<xml-style>
  <header-row>5</header-row>
  ... Other config skipped
</xml-style>

