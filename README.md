# postOCR-Extraction
Extraction of user-specified data from OCR-processed Microsoft Word files to Microsoft Excel.

When one recieves a document as a picture or a PDF file, there might be a need to extract some information.
Time required to extract data by hand is relevantly low when there are few documents.
However, it becomes tiresome when one has to extract data from 100+ documents.

The following tool aims to reduce time in extracting data from same-type documents.

Tool requires OCR-processed documents only in docx format due to fact
that paragraphs, table and other Word document objects have relatively same position on Word document page comparing to preprocessed document.

Required software:
<ul>
  <li>Microsoft Word.</li>
  <li>Microsoft Excel.</li>
</ul>

# Launch
Tool can be launched from command-line.
It requires at least two arguments:
<ul>
  <li>A path to Excel file with configuration data.</li>
  <li>A path to Word file or directory with Word files.</li>
</ul>
There may be additional arguments that are considered of the same path type as the second argument (file or directory)

Tool can be launched as it is, then a form will appear where user can input file paths, explained above.

Tool can be launched in debug mode by providing an argument "-d" before existing arguments.

# Configuration data
A configuration file in Excel format must be provided in order to specify concrete data to be extracted.

Please follow guidance below to correctly compose configuration file.
<ol>
  <li>Create an Excel file with xlsx extension.</li>
  <li>Have only one worksheet in file. Worksheet would contain search parameters among single document type.</li>
  <li>Add the following data into first column: <br>
    <table>
      <tr><td>Similarity algorithm >>></td></tr>
      <tr><td>Grid size >>></td></tr>
      <tr><td>Field name >>></td></tr>
      <tr><td>Value type >>></td></tr>
      <tr><td>Grid coordinates >>></td></tr>
      <tr><td>Field expression >>></td></tr>
      <tr><td>Search values >>></td></tr>
    </table>
  </li>
  These are required to determine borders of configuration file.
  <li>
    Next to the 'Similarity algorithm' identifier set name of string similarity algorithm. Currently, the following algorithms are implemented:
  </li>
  <ul>
    <li>JaroAlgorithm</li>
    <li>JaroWinklerAlgorithm</li>
    <li>LevensteinAlgorithm</li>
    <li>RatcliffObreshelpAlgorithm</li>
  </ul>  
  <li>Next to the 'Grid size' field set a number that denotes a grid size that which document page would be distributed by.</li>
</ol>
The following actions are used to define an individual search parameter and can be applied to any number of parameters.
<ol>
  <li>Add basic description of field - its name and value type:</li>
  <table>
    <tr>
      <td>Field name >>></td>
      <td>Goods</td>
    </tr>
    <tr>
      <td>Value type >>></td>
      <td>String</td>
    </tr>
    <tr>
      <td>Grid coordinates >>></td>
      <td></td>
    </tr>
    <tr>
      <td>Field expression >>></td>
      <td></td>
    </tr>    
    <tr>
      <td>Search values >>></td>
      <td></td>
    </tr>    
  </table>
  <li>In the row with 'Grid coordinates' identifier add segment coordiantes, in which search should be made. Coordinates must be of two comma-separated numbers, where:</li>
  <ul>
    <li>first coordinate - segment row,</li>
    <li>second coordinate - segment column.</li>
  </ul>
  Both coordinates must be non-negative and less than "Grid size".
    <table>
    <tr>
      <td>Field name >>></td>
      <td>Goods</td>
    </tr>
    <tr>
      <td>Value type >>></td>
      <td>String</td>
    </tr>
    <tr>
      <td>Grid coordinates >>></td>
      <td>1,1</td>
    </tr>
    <tr>
      <td>Field expression >>></td>
      <td></td>
    </tr>    
    <tr>
      <td>Search values >>></td>
      <td></td>
    </tr>    
  </table>
  <li>In the row with 'Field expression' identifier add regular expression pattern, which is used to match text, and a string, used to compare with regular expression match. Pattern and string are separated by semicolon.</li>
  <table>
    <tr>
      <td>Field name >>></td>
      <td>Goods</td>
    </tr>
    <tr>
      <td>Value type >>></td>
      <td>String</td>
    </tr>
    <tr>
      <td>Grid coordinates >>></td>
      <td>1,1</td>
    </tr>
    <tr>
      <td>Field expression >>></td>
      <td>^[gods]{4,6};Goods</td>
    </tr>    
    <tr>
      <td>Search values >>></td>
      <td></td>
    </tr>    
  </table>
  This pattern is a staring point of single field search, from which a specific search is launched.
  
  A Soundex algorithm can be used instead of using a regular expression pattern. In order to use it there has to be the following notaion: "soundex(expression);expreession".
  For instance, an expression "soundex(Goods);Goods" can be used in place of "^[gods]{4,6};Goods" to find a word "Goods".
  <li>In the row with 'Search values' identifier add the following composite parameter:</li>
  <ul>
    <li>Regular expression pattern or Soundex expression - used to match strings.</li>
    <li>Number that would indicate how many rows should be offset from previously found match to perform search by current pattern (can be empty). Positive number indicates movement to the bottom, negative - to the top.</li>
    <li>Number that would indicate the following:</li>
    <ul>
      <li>if it is empty or equals 0 - the search is performed from offset row start to end.</li>
      <li>if it equals 1 - the search is performed from vertically offset point of previous match to offset row end.</li>
      <li>if it equals -1 - the search is performed from offset row start to vertically offset point of previous match.</li>
    </ul>
  </ul>
  Numeric parts have different meaning if the value type contains "Table": these are used as number of table rows and columns, which should be offset from previously matched table cell.
  Pattern and numeric parts are separated by semicolon.
    
  <table>
    <tr>
      <td>Field name >>></td>
      <td>Goods</td>
    </tr>
    <tr>
      <td>Value type >>></td>
      <td>String</td>
    </tr>
    <tr>
      <td>Grid coordinates >>></td>
      <td>1,1</td>
    </tr>
    <tr>
      <td>Field expression >>></td>
      <td>^[gods]{4,6};Goods</td>
    </tr>    
    <tr>
      <td>Search values >>></td>
      <td>\d{2}\.\d{2}\.\d{4};1;</td>
    </tr>    
  </table>
  <li>
    There can be any number of specific parameters.
        <table>
      <tr>
        <td>Field name >>></td>
        <td>Goods</td>
      </tr>
      <tr>
        <td>Value type >>></td>
        <td>String</td>
      </tr>
      <tr>
        <td>Grid coordinates >>></td>
        <td>1,1</td>
      </tr>
      <tr>
        <td>Field expression >>></td>
        <td>^[gods]{4,6};Goods</td>
      </tr>    
      <tr>
        <td>Search values >>></td>
        <td>\d{2}\.\d{2}\.\d{4};1;</td>
      </tr> 
      <tr>
        <td></td>
        <td>(.*OIL);-1;</td>
      </tr> 
      <tr>
        <td></td>
        <td>(.*OIL);;</td>
      </tr> 
    </table>
  </li>
  <li>The value matched to last specific parameter is considered terminal, therefore search by single field is done.</li>
  <strong>NB.</strong> Soundex algorithm is not implemented for terminal parameter and, therefore, regular expresion pattern must be used.
</ol>

# Configuration file template

<table>
  <tr>
    <td>Similarity algorithm >>></td>
    <td>algorithm_name</td>
  </tr>
  <tr>
    <td>Grid size >>></td>
    <td>grid_size</td>
  </tr>  
  <tr>
    <td>Field name >>></td>
    <td>Field1</td>
    <td>Field2</td>
    <td>...</td>
    <td>FieldN</td>
  </tr>
  <tr>
    <td>Value type >>></td>
    <td>String</td>
    <td>Table/String</td>
    <td>...</td>
    <td>Number</td>
  </tr>
  <tr>
    <td>Grid coordinates >>></td>
    <td>coordinate1,coordinate2</td>
    <td>coordinate1,coordinate2</td>
    <td>...</td>
    <td>coordinate1,coordinate2</td>
  </tr>
  <tr>
    <td>Field expression >>></td>
    <td>re_pattern;check_string</td>
    <td>soundex_expression;check_string</td>
    <td>...</td>
    <td>re_pattern;check_string</td>
  </tr>    
  <tr>
    <td>Search values >>></td>
    <td>re_pattern1;line_offset1;horizontal_search_status</td>
    <td>soundex_expression1;row_offset1;column_offset1</td>
    <td>...</td>
    <td>re_pattern1;line_offset1;horizontal_search_status1</td>
  </tr> 
  <tr>
    <td></td>
    <td></td>
    <td>re_pattern1;row_offset2;column_offset2</td>
    <td>...</td>
    <td>re_pattern2;line_offset2;horizontal_search_status2</td>
  </tr> 
  <tr>
    <td></td>
    <td></td>
    <td></td>
    <td>...</td>
    <td>re_pattern3;line_offset3;horizontal_search_status3</td>
  </tr> 
</table>
