# sheets2pdf_gs
Library for converting Google Sheets™ Into PDF. 

Featurs
 * ⚙️ All PDF settings including colontitles = custom headers and footers.
 * 🔌Input parameters is a single plain object.
 * 👀 Live preview! Copy [sample Sheet]([url](https://docs.google.com/spreadsheets/d/1HwUaZk86BtrPdQ1RYILwTcRwJUUClgqtAPEpMAsX0y8/edit#gid=475497297) with UI → Go to the menu `⚡ Test Automation > 🍋 Drive > 🍎Print 2 PDF`

## Install

Use library code:
`1_hqx3V0ZI0VOFaeJN0D1qYtdhpCeuC1ct10AnHnnbmDblW8m058hfH5n`

## Use
```
Sheets2Pdf.getBlob(options);
Sheets2Pdf.save2Drive(options);
Sheets2Pdf.getUrl(options);
Sheets2Pdf.getOriginalParameters(options); // for tests
```

## Set options
Options is a plain object. 
All parameters have a type `string` or `number`. String parameters are mostly case sensitive.
Some parameters have long names: `A4 (8.27" x 11.69")`. This is made for direct conversion of user UI values into options. TODO: add ability to use numbers.

Sample `options`:
```
{
    
    // General options
    // .....................................................................................
    "delimiter": "|",    // for parsing lists
    "file_id": "",       // if ommited will try ActiveWorkbook
    "folder_id": "",     // if ommited will save to Root Drive Folder
    "pdf_name": "",      // if blank, will create name of file name + stamp
    
    
    // PDF options
    // ......................................................................................
    "sheet_names": "Parameters",              // delimited: Sheet1|Sheet2
    
    "range_a1": "",                           // delimites: A1|A2.
                                              // If not present, prints all sheets
    
    "page_size": 'A4 (8.27" x 11.69")',       // Possible values:
                                              // Letter (8.5" x 11")
                                              // Tabloid (11" x 17")
                                              // Legal (8.5" x 14")
                                              // Statement (5.5" x 8.5")
                                              // Executive (7.25" x 10.5")
                                              // Folio (8.5" x 13")
                                              // A3 (11.569" x 16.54")
                                              // A4 (8.27" x 11.69")
                                              // A5 (5.83" x 8.27")
                                              // B4 (9.84" x 13.90")
                                              // B5 (6.93" x 9.84")
                                              // Custom size
                                              // 📐Fit to range
                                              // ...
                                              // Also possible syntax: WxH in inches: 12.23x10.582
                                              
    "width": "",                              // inches, if page size set to "Custom size"
    "height": "",                             // inches, if page size set to "Custom size"
    
    "is_landscape": true,                     // true,false
    
    "scale": "Normal (100%)",                 // Possible values:
                                              // Normal (100%)
                                              // Fit to width
                                              // Fit to height
                                              // Fit to page
                                              // Custom number
                                              // To page breaks
                                              
    "scale_percent": "",                      // Number between 0 and 1.
                                              // if scale was set to "Custom number"
    
    "margins": "Wide",                        // Possible values: 
                                              // Normal
                                              // Narrow
                                              // Wide
                                              // Custom numbers
                                              //
                                              //            Top     Bottom  Left  Rigth
                                              // "Normal": [0.75,   0.75,   0.7,  0.7]
                                              // "Wide":   [1,      1,      1,    1]
                                              // "Narrow": [0.75,   0.75,   0.25, 0.25]
                                              
    "margins_top": "",                        // inches, if margins set to "Custom numbers"
    "margins_bottom": "",                     // inches, if margins set to "Custom numbers"
    "margins_left": "",                       // inches, if margins set to "Custom numbers"
    "margins_rigth": "",                      // inches, if margins set to "Custom numbers"
    
    "hirizontal_alignment": "Rigth",          // Possible values: 
                                              // Left
                                              // Center
                                              // Rigth
                                              
    "vertical_alignment": "Bottom",           // Possible values: 
                                              // Top
                                              // Center
                                              // Bottom
    
    "page_breaks_rows": "8|9|14",             // delimited: 8|9. Set scale to "To page breaks"
    "page_breaks_columns": "1|2",             // delimited: 8|9. Set scale to "To page breaks"
    
    "show_gridlines": true,                   // true,false       
    "show_notes": true,                       // true,false
    
    "repeat_rows": false,                     // true,false 
    "repeat_columns": false,                  // true,false

    // Default colontitles
    "show_date": false,                       // true,false
    "show_time": false,                       // true,false
    "show_page_numbers": true,                // true,false
    "show_workbook": false,                   // true,false
    "show_sheet": false,                      // true,false
    
    // Custom colontitles
    "colontitles_top_left": "",               // Example: It's Thursday! 
    "colontitles_top_center": "",             // Example: Printed "{{sheet}}" from "{{book}}"
    "colontitles_top_rigth": "",              // Example: Page #{{page}}
    "colontitles_bottom_left": "",            // TODO: is there a way to get total number of pages?
    "colontitles_bottom_center": "",          // 👉 multiline colontitles are not allowed
    "colontitles_bottom_rigth": "",           // 👉 if set, colontitles will show even if all margins = 0
                                              // 👉 parse dates on your side to get more flexible uotput
    
    // Special tags for custom colontitles
    // ......................................................................................
    "tag_open": "{{",                         // optional, default - {{
    "tag_close": "}}",                        // optional, default - }}
    "call_page_num": "page",                  // optional, default - page
    "call_sheet": "sheet",                    // optional, default - sheet
    "call_workbook": "book"                   // optional, default - book
  }
```

## Test
```
function myTestDefaults() {
  var test = Sheets2Pdf.getOriginalParameters({
     file_id: '1ycWG5XxVK8ZI6MrEiumuNXP9MDvh4dbJbydVWYmaVis',
     sheet_names: 'Sheet1',
  });
  console.log(test);
}

function myTestBlob() {
  var blob = Sheets2Pdf.getBlob({
     file_id: '1ycWG5XxVK8ZI6MrEiumuNXP9MDvh4dbJbydVWYmaVis',
     sheet_names: 'Sheet1',
  });
  console.log(blob.getContentType());
}

function myTest2Drive() {
  var test = Sheets2Pdf.save2Drive({
     file_id: '1ycWG5XxVK8ZI6MrEiumuNXP9MDvh4dbJbydVWYmaVis',
     sheet_names: 'Sheet1',
  });
  console.log(test); // output object
}

function myTestUrl() {
  var url = Sheets2Pdf.getUrl({
     file_id: '1ycWG5XxVK8ZI6MrEiumuNXP9MDvh4dbJbydVWYmaVis',
     sheet_names: 'Sheet1',
  });
  console.log(url);
}
```

## Get the original parameters

Reproduce the folowing in Chrome:

1. Create a new Google Sheet (use fast url: sheets.new)
2. Press [Ctrl]+[P] to see printing settings
3. Press [Ctrl]+[Shift]+[i] to open developer tools.
4. Press [next] button in the printing settings
5. Open network in developer console and find the line starting with "pdf..."

Source: [this answer](https://stackoverflow.com/a/58503138/5372400) on StackOverflow.


Library of all possible parameters:

```
    var datajson = [
    null,null,null,null,null,null,null,null,null,0,
    [
      [
        "0", // Sheet Id
        
        // Optional Range Borders
        8,   // row 1
        13,  // row 2
        2,   // column 1
        7    // column 2
      ]
    ],
    10000000,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
    44776.5574952662, // timestamp
    null,
    null,
    [
      1,    // show notes 1/0
      null, /** TODO: find out what is it: allowad values: number or null */
      0,    // show gridlines 1/0
      0,    // show page numbers 1/0
      0,    // show workbook name 1/0
      0,    // show sheet name 1/0
      0,    // show current date
      0,    // show current time
      1,    // repeat freezed rows
      1,    // repeat freezed columns
      2,    // page order 1 = down, then over, 2 = over, then down
      2,    // show colontitles: 1 = not to show, 2 = show

      // Colontitles are optional. set to null if not needed
      // Colontitles|TOP
      [
        // Colontitles|TOP|LEFT
        [ "h:mm:ss am/pm🌵 h:mm am/pm🌵 H:mm:ss🌵 H:mm" ],
        // Colontitles|TOP|CENTER
        [ "MM-dd-yyyy👀 M/d/yy👀 MM-dd-yy👀 M/d👀 MM-dd" ],
        // Colontitles|TOP|RIGHT
        [ "d-MMM😅d-MMM-yyyy😅MMMM d, yyyy😅MMMM d🥸😅MMM-d" ]
      ],

      // Colontitles|BOTTOM
      [
        // Colontitles|BOTTOM|LEFT
        [ "🎯 " ],
        // Colontitles|BOTTOM|CENTER
        [  "🔥 🔥 " ],
        // Colontitles|BOTTOM|RIGHT
        [ "M/d/yyyy👉🏻 yyyy-MM-dd" ]
      ],

      2, // hirizontal alignment: 1 = left, 2 = rigth, 3 = center
      1  // vertical alignment: 1 = top, 2 = center, 3 = bottom
    ],
    [

      // size
      "letter", // letter,tabloid,Legal,statement,executive,folio,A3,A4,A5,B4,B5
                // or custom size in inches: 11x8.6 = W x H
                // Q: can we use an array here?

      0, // pare oriantation: 0 = horizontal, 1 = vertical
      2, // scale: 1 = 100%, 2 = to width, 3 = to height, 4 = to page size, 5 = user defined, 6 = to page breaks
      1, // user defined scale: 1 = 100%

      // margins, inches
      [
              //                  narrow  | wide  |  usual
        0.75, // margin|TOP         0.75       1      0.75
        0.75, // margin|BOTTOM      0.75       1      0.75
        0.7,  // margin|LEFT        0.25       1      0.7
        0.7   // margin|RIGTH       0.25       1      0.7
      ]
    ],
    null,0,

    // line breaks, if not used → set to null
    // can be represented by a pair of same numbers
    // line break starts after this line
    [
      [
        "0", // sheet id
        // horizontal line breaks
        [ [ 0, 24 ],
          [ 32, 32 ] ],
        // vertical line breaks
        [ [ 1, 1 ] ]
      ]
    ]
  ];

```

## Author
@max__makhrov

![CoolTables](https://raw.githubusercontent.com/cooltables/pics/main/logos/ct_logo_small.png) [cooltables.online](https://www.cooltables.online/)
