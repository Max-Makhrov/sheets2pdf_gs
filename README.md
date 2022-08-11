# sheets2pdf_gs
Library for converting Google Sheets Into PDF.

## Install

Use library code:
`under_constryuction`

## Use
```
Sheets2Pdf.getBlob(options);
Sheets2Pdf.save2Drive(options);
Sheets2Pdf.getUrl(options);
Sheets2Pdf.getOriginalParameters(options); // for tests
```
Live preview:
Copy sample Sheet with options.

## Set options
Sample `options`:
```
{
    // for parsing lists
    "delimiter": "|",
    // if ommited will try ActiveWorkbook
    "file_id": "",
    
    "sheet_names": "Parameters",
    "range_a1": "",
    "page_size": 'A4 (8.27" x 11.69")',
    "width": "",
    "height": "",
    "is_landscape": true,
    "scale": "Normal (100%)",
    "scale_percent": "",
    "margins": "Wide",
    "margins_top": "",
    "margins_bottom": "",
    "margins_left": "",
    "margins_rigth": "",
    "hirizontal_alignment": "Rigth",
    "vertical_alignment": "Bottom",
    "page_breaks_rows": "8|9|14",
    "page_breaks_columns": "1|2",
    "show_gridlines": true,
    "show_notes": true,
    "show_page_numbers": true,
    "repeat_rows": false,
    "repeat_columns": false,
    "show_workbook": false,
    "show_sheet": false,
    "show_date": false,
    "show_time": false,
    "tag_open": "{{",
    "tag_close": "}}",
    "call_page_num": "page",
    "call_sheet": "sheet",
    "call_workbook": "book",
    "colontitles_top_left": "",
    "colontitles_top_center": "",
    "colontitles_top_rigth": "",
    "colontitles_bottom_left": "",
    "colontitles_bottom_center": "",
    "colontitles_bottom_rigth": ""
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
