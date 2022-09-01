var Sheets2Pdf = this;

/**
 * ## About
 * Function `getOriginalParameters` gets the original PDF printing settings from Sheets.
 * These settings are then passed to a server.
 * 
 * 
 * ## Sample usage
 * Please copy this code and run the function:
 * ```javascript
 * function myTestDefaults() {
 *   var test = Sheets2Pdf.getOriginalParameters({
 *      file_id: '1ycWG5XxVK8ZI6MrEiumuNXP9MDvh4dbJbydVWYmaVis',
 *      sheet_names: 'Sheet1',
 *   });
 *   console.log(test); // get all default options
 * }
 * ```
 * 
 * ## Author
 * `@max__makhrov` 
 * 
 * ![CoolTables](https://raw.githubusercontent.com/cooltables/pics/main/logos/ct_logo_small.png) [cooltables.online](https://www.cooltables.online/) 
 * 
 * @param {Object} options `options` is a plain object with settings: `{param: value}`
 * 
 * 👀All `options` are listed in 🔗[GitHub Repo](https://github.com/Max-Makhrov/sheets2pdf_gs).
 * @return {Object[]} array
 * 
 */
var getOriginalParameters = function(options) {
  return getOriginalParameters_(options);
}

/**
 * ## About
 * Function `getBlob` gets the PDF blob.
 * 
 * 
 * ## Sample usage
 * Please copy this code and run the function:
 * ```javascript
 * function myTestBlob() {
 *   var blob = Sheets2Pdf.getBlob({
 *      file_id: '1ycWG5XxVK8ZI6MrEiumuNXP9MDvh4dbJbydVWYmaVis',
 *      sheet_names: 'Sheet1',
 *   });
 *   console.log(blob.getContentType());
 * }
 * ```
 * 
 * ## Author
 * `@max__makhrov` 
 * 
 * ![CoolTables](https://raw.githubusercontent.com/cooltables/pics/main/logos/ct_logo_small.png) [cooltables.online](https://www.cooltables.online/) 
 * 
 * @param {Object} options `options` is a plain object with settings: `{param: value}`
 * 
 * 👀All `options` are listed in 🔗[GitHub Repo](https://github.com/Max-Makhrov/sheets2pdf_gs).
 * @return {Object} blob
 * 
 */
var getBlob = function(options) {
  return getPdfBlob_(options);
} 

/**
 * ## About
 * Function `save2Drive` ⚠️**saves PDF to your Drive**⚠️.
 * 
 * 
 * ## Sample usage
 * Please copy this code and run the function:
 * ```javascript
 * function myTest2Drive() {
 *   var test = Sheets2Pdf.save2Drive({
 *      file_id: '1ycWG5XxVK8ZI6MrEiumuNXP9MDvh4dbJbydVWYmaVis',
 *      sheet_names: 'Sheet1',
 *   });
 *   console.log(JSON.stringify(test)); // output object
 * }
 * ```
 * 
 * ## Author
 * `@max__makhrov` 
 * 
 * ![CoolTables](https://raw.githubusercontent.com/cooltables/pics/main/logos/ct_logo_small.png) [cooltables.online](https://www.cooltables.online/) 
 * 
 * @param {Object} options `options` is a plain object with settings: `{param: value}`
 * 
 * 👀All `options` are listed in 🔗[GitHub Repo](https://github.com/Max-Makhrov/sheets2pdf_gs).
 * @return {Object} resulting data
 * 
 */
var save2Drive = function(options) {
  return save2Drive_(options);
}

/**
 * ## About
 * Function `getUrl` returns URL you can use to save PDF locally.
 * 
 * 
 * ## Sample usage
 * Please copy this code and run the function:
 * ```javascript
 * function myTestUrl() {
 *   var url = Sheets2Pdf.getUrl({
 *      file_id: '1ycWG5XxVK8ZI6MrEiumuNXP9MDvh4dbJbydVWYmaVis',
 *      sheet_names: 'Sheet1',
 *   });
 *   console.log(url);
 * }
 * ```
 * 
 * ## Author
 * `@max__makhrov` 
 * 
 * ![CoolTables](https://raw.githubusercontent.com/cooltables/pics/main/logos/ct_logo_small.png) [cooltables.online](https://www.cooltables.online/) 
 * 
 * @param {Object} options `options` is a plain object with settings: `{param: value}`
 * 
 * 👀All `options` are listed in 🔗[GitHub Repo](https://github.com/Max-Makhrov/sheets2pdf_gs).
 * @return {String} url
 * 
 */
var getUrl = function(options) {
  return getUrl_(options);
}





/** generates random Id */
function getEsId_() {
    return ( Math.round( Math.random() * 10000000 ) );
}


/**
 * get the PDF Url you can use to save PDF locally
 */
function getUrl_(sets) {
  /** get the orijinal parameters */
  var datajson = getOriginalParameters_(sets); 
  var bookid = sets.book.getId();
  /** create the url */
  var url = "https://docs.google.com/spreadsheets/d/" + bookid + "/pdf?" +
        "esid=" + getEsId_() + "&" +
        "id=" + bookid + "&" +
        "a=true&" +
        "pc=" + encodeURIComponent( JSON.stringify( datajson ) ) + "&" +
        "gf=[]&" +
        "lds=[]";
  return url;
}


/**
 * Creates PDF and saves in to Drive
 * 
 */
function save2Drive_(sets) {
  var blob = getPdfBlob_(sets);
  /** get pdf name */
  if (!sets.pdf_name) {
    sets.pdf_name = sets.book.getName() + 
      ' ' + 
      Utilities.formatDate(
        new Date(),
        Session.getTimeZone(),
        'yyyy-MM-dd HH:mm:ss'
        )
  }
  /** get folder */
  var folder;
  if (sets.folder_id) {
    folder = DriveApp.getFolderById(sets.folder_id);
  } else {
    folder = DriveApp.getRootFolder();
  }  
  /** create PDF */
  var pdf = folder.createFile( blob ).setName( sets.pdf_name + ".pdf" );
  return {
    pdf_name: sets.pdf_name,
    pdf_url: pdf.getUrl(),
    pdf_id: pdf.getId(),
    pdf_blob: blob,
    pdf: pdf,
    folder_name: folder.getName(),
    folder_url: folder.getUrl(),
    folder_id: folder.getId(),
    folder: folder
  } 
}



/**
 * creates PDF blob by sets
 */
function getPdfBlob_(sets) {

  /** get the orijinal parameters */
  var datajson = getOriginalParameters_(sets);

  /** Get the Blob */
  var urloptions = {
    'method': 'post',
    'payload': "a=true&pc=" + JSON.stringify(datajson) + "&gf=[]&lds=[]",
    'headers': {
      Authorization: "Bearer " + ScriptApp.getOAuthToken(),
      date: new Date()},
    'muteHttpExceptions': true
  };
  var bookid = sets.book.getId();
  var blob = UrlFetchApp.fetch( "https://docs.google.com/spreadsheets/d/" + bookid + "/pdf?" +
        "esid=" + getEsId_() + "&" +
        "id=" + bookid, urloptions ).getBlob();
  return blob;
}




/**
 * 
 * Get original printing parameters
 * 
 */
function getOriginalParameters_(sets) {
    /** 1️⃣ get sheet & range and borders */
  var book, err_msg;
  if (!sets.file_id || sets.file_id === '') {
    book = SpreadsheetApp.getActive();
  } else {
    err_msg = 'no workbook with id = ' + sets.file_id;
    try {
      book = SpreadsheetApp.openById(sets.file_id);
    } catch (err) {
      throw err_msg;
    }
    if (!book) { throw err_msg; }
  }
  sets.book = book; // add to sets for the future use
  // final result of this step
  var i_sheets = [];
  var getSheet_ = function(sheet_name) {
    var err_msg = 'no sheet ' + sheet_name;
    try {
      var sheet = book.getSheetByName(sheet_name);
    } catch (err) {
      throw err_msg;
    }
    if (!sheet) { throw err_msg; }
    return sheet;
  }
  // get sheet names
  var sheetNames = sets.sheet_names.toString().split(sets.delimiter);
  // get ranges if present
  var ranges_a1 = [];
  if (sets.range_a1 && sets.range_a1 !== '') {
    ranges_a1 = sets.range_a1.split(sets.delimiter);
    // assign same ranges as the last one
    if (ranges_a1.length < sheetNames.length) {
      var last_range_a1 = ranges_a1[ranges_a1.length-1];
      for (var i = ranges_a1.length; i < sheetNames.length; i++) {
        ranges_a1.push(last_range_a1); }
    }
  }
  // console.log(ranges_a1);
  /**
   * get sheet sets for 1 sheet:
   * 
   * [ {string}SheetId, r1,r2,c1,c2 ]
   * 
   */
  var sheets = [], ranges = [];
  var rownum, columnnum, rownum2, columnnum2;
  var rownums = [], columnnums = [], rownums2 = [], columnnums2 = [];
  var getSheetSets_ = function(sheet_name, index) {
    var sheet = getSheet_(sheet_name);
    sheets.push(sheet);
    var result = [sheet.getSheetId().toString()];
    /** add range if needed */
    var range = null;
    if (ranges_a1.length) {
      var range_a1 = ranges_a1[index];
      err_msg = 'no range ' + sets.range_a1;
      try {
        range = sheet.getRange(range_a1)
      } catch(err) {
        throw err_msg;
      }
      if (!range) { throw err_msg; }
      var maxRows = sheet.getMaxRows();
      var maxColumns =sheet.getMaxColumns();
      rownum = range.getRow();
      columnnum = range.getColumn();
      rownum2 = range.getLastRow();
      if (rownum2 > maxRows) {
        rownum2 = maxRows;
      }
      columnnum2 = range.getLastColumn();
      if (columnnum2 > maxColumns) {
        columnnum2 = maxColumns;
      }
      result = result.concat([
        rownum-1,
        rownum2,
        columnnum-1,
        columnnum2
      ]);
    }
    ranges.push(range);
    rownums.push(rownum);
    columnnums.push(columnnum);
    rownums2.push(rownum2);
    columnnums2.push(columnnum2);
    return result;
  }
  // loop sheets
  i_sheets = sheetNames.map(getSheetSets_);

  /** 2️⃣ get timestamp parsed */
  var encodeDate_ = function( dt ){
    var yy = dt.getFullYear(); 
    var mm = dt.getMonth();
    var dd = dt.getDate();
    var hh = dt.getHours();
    var ii = dt.getMinutes();
    var ss = dt.getSeconds();
    var days=[31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if(((yy % 4) == 0) && (((yy % 100) != 0) || ((yy % 400) == 0)))days[1]=29;
    for (var i=0; i<mm; i++) {
      dd += days[ i ];
    }
    yy--;
    var result = ((((yy * 365 + ((yy-(yy % 4)) / 4) - 
      ((yy-(yy % 100)) / 100) + 
      ((yy-(yy % 400)) / 400) + dd - 693594) * 24 + hh) * 60 + ii) * 60 + ss)/86400.0;
    return result;
  }
  var i_date = encodeDate_(new Date());
  // logJson_(i_date);

  /** 3️⃣ show... */
  var parseTrueFalse_ = function(value, iffalse, iftrue) {
    if (iftrue === undefined ) {
      iftrue = 1;
    }
    if (iffalse === undefined) {
      iffalse = 0;
    }
    if (value) {
      return iftrue; }
    return iffalse;
  }
  var i_notes = parseTrueFalse_(sets.show_notes);
  var i_gridlines = parseTrueFalse_(sets.show_gridlines, 1, 0);
  var i_page_numbers = parseTrueFalse_(sets.show_page_numbers);
  var i_workbook = parseTrueFalse_(sets.show_workbook);
  var i_sheet = parseTrueFalse_(sets.show_sheet);
  var i_show_date = parseTrueFalse_(sets.show_date);
  var i_show_time = parseTrueFalse_(sets.show_time);
  var i_freezed_rows = parseTrueFalse_(sets.repeat_rows);
  var i_freezed_columns = parseTrueFalse_(sets.repeat_columns);
  var i_page_order = parseTrueFalse_(sets.is_reverse_page_order,1,2);
  // console.log(i_page_order);
  // page oriantetion --  0 = horizontal, 1 = vertical
  var i_page_oriantation = parseTrueFalse_(sets.is_landscape, 1, 0);
  // console.log(sets.is_landscape);
  // console.log(i_page_oriantation);

  /** 4️⃣ Colontitles */
  /**
   * Replace all for ES5 Compatability
   * https://stackoverflow.com/a/1144788/5372400
   */
  var replaceAll_ = function(str, find, replace) {
    if (!str) {
      return str;
    }
    return str.replace(new RegExp(escapeRegExp_(find), 'g'), replace);
  }
  var escapeRegExp_ = function(string) {
    // $& means the whole matched string
    if (!string) { return string; }
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); 
  }
  /**
   * starts with for ES5 Compatability
   */
  function startsWith_(str, word) {
      return str.lastIndexOf(word, 0) === 0;
  }
  var i_colontitles_top = null;
  var i_colontitles_bottom = null;
  /**
   * creating a proper Sheets 2 DDF 
   * colontitle parameter
   * 
   */
  var parseColontitle_ = function(text) {
    /** options & defaults */
    var original_t = text;
    sets.tag_open = sets.tag_open || '{{';
    sets.tag_close = sets.tag_close || '}}'
    sets.call_page_num = sets.call_page_num || 'page';
    sets.call_sheet = sets.call_sheet || 'sheet';
    sets.call_workbook = sets.call_workbook || 'book';
    // var sets = {
    //   tag_open: '{{',
    //   tag_close: '}}',
    //   parameter_open: '(',
    //   parameter_close: ')',
    //   call_date: 'date',
    //   call_time: 'time',
    //   call_page_num: 'page',
    //   call_sheet: 'sheet',
    //   call_workbook: 'book'
    // }
    /** constants */
    // special chars
    var Dict = {
      // use TEXT(now(), "YYY-MM-DD hh:mm:ss")
      // call_date: {
      //   char1: '\uEE13',
      //   char2: '\uEE14' 
      // },
      // call_time: {
      //   char1: '\uEE17',
      //   char2: '\uEE18'
      // },
      call_page_num: {
        char1: '\uEE12'
      },
      // call_page_num2: {
      //   char1: '\uEE15'
      // },
      // call_page_num3: {
      //   char1: '\uEE16'
      // },
      call_workbook: {
        char1: '\uEE10'
      },
      call_sheet: {
        char1: '\uEE11'
      }
    };
    /** 
     *  parse main tags of string
     */
    var getFlatTextParts_ = function(text, open_tag, close_tag) {
      // compose some rare opening and closure
      var basicText = '\\\\';
      /**
       * repeat for ES5 Compatibility
       */
      var repeat_ = function(txt, times) {
        var res = [];
        for (var i = 0; i < times; i++) {
          res.push(txt);
        }
        return res.join('');
      }
      var pre = repeat_(basicText, 20);
      var open = 'open' + pre;
      var close = pre + 'close';
      text = replaceAll_(text, open_tag, open);
      text = replaceAll_(text, close_tag, close);
      var open_re = 'open' + pre + pre;
      var close_re = pre + pre + 'close';
      // compose regex to get what's inside tags
      var re_txt = open_re + '(.*?(?!' + close_re + ')*)' + close_re;
      var re = new RegExp(re_txt, 'gm');
      // get the results
      var result1 = [], result0 = [];
      while(null != (z=re.exec(text))) {
        // console.log(z);
        result0.push(z[0]);
        result1.push(z[1]); 
      }
      return {
        result0: result0,
        result1: result1
      };
    }
    var list_tagged = getFlatTextParts_(
      text, 
      sets.tag_open, 
      sets.tag_close).result1;
    /**
     * parse single parameter 
     * date(yyy-MM-dd)
     * page
     * 
     * get replacement and text to replace with
     */
    var getColontitleReplacemant_ = function(text) {
      var user_key = '', parameter;
      var result ={}, replacewith = ''
      for (var k in Dict) {
      user_key = sets[k];
        if (startsWith_(text, user_key)) {     
          parameter = getFlatTextParts_(
            text, 
            sets.parameter_open,
            sets.parameter_close).result1[0];
          replacewith = Dict[k].char1;
          if (Dict[k].char2) {
            replacewith += parameter + Dict[k].char2
          }
          result = {
            replace: sets.tag_open + text + sets.tag_close, 
            replace_with: encodeURIComponent(replacewith)
          }
          // logJson_(result)
          return result;
        } // found the result!
      }
      var err_text = '🤪Wrong Colontitle Syntax. Unknown key = "' + 
        text + '" in text \n' + original_t;
      throw err_text;
    }
    var oreplace = {}, newt = original_t;
    for (var i = 0; i < list_tagged.length; i++) {
      oreplace = getColontitleReplacemant_(list_tagged[i]);
      newt = replaceAll_(newt,
        oreplace.replace,
        oreplace.replace_with);
    }
    // console.log(newt);
    return newt;
  }
  /** end of parsing colontitles functiion */ 
  // Top Colontitles
  if (  
    (sets.colontitles_top_left !== '' && sets.colontitles_top_left) ||
    (sets.colontitles_top_center !== '' && sets.colontitles_top_center) ||
    (sets.colontitles_top_rigth !== '' && sets.colontitles_top_rigth)
  ) {
    i_colontitles_top = [
      [parseColontitle_(sets.colontitles_top_left)],
      [parseColontitle_(sets.colontitles_top_center)],
      [parseColontitle_(sets.colontitles_top_rigth)]
    ]
  }
  // Bottom Colontitles
  if (  
    (sets.colontitles_bottom_left !== '' && sets.colontitles_bottom_left) ||
    (sets.colontitles_bottom_center !== '' && sets.colontitles_bottom_center) ||
    (sets.colontitles_bottom_rigth !== '' && sets.colontitles_bottom_rigth)
  ) {
    i_colontitles_bottom = [
      [parseColontitle_(sets.colontitles_bottom_left)],
      [parseColontitle_(sets.colontitles_bottom_center)],
      [parseColontitle_(sets.colontitles_bottom_rigth)]
    ]
  }
  var i_show_colontitles = 1; // not show
  if (i_colontitles_top || i_colontitles_bottom) {
    i_show_colontitles = 2;
  }
  // logJson_(i_colontitles_top);
  // logJson_(i_colontitles_bottom);


  /** 4️⃣ Some custom options */
  var custom_pdf_options = {
      page_size: {
        'A4 (8.27" x 11.69")':       'A4', // default
        'Letter (8.5" x 11")':       'letter',
        'Tabloid (11" x 17")':       'tabloid',
        'Legal (8.5" x 14")':        'legal',
        'Statement (5.5" x 8.5")':   'statement',
        'Executive (7.25" x 10.5")': 'executive',
        'Folio (8.5" x 13")':        'folio',
        'A3 (11.569" x 16.54")':     'A3',
        'A5 (5.83" x 8.27")':        'A5',
        'B4 (9.84" x 13.90")':       'B4',
        'B5 (6.93" x 9.84")':        'B5'
      },

      scale: {
        'Normal (100%)': 1, // dafault
        'Fit to width': 2,
        'Fit to height': 3,
        'Fit to page': 4,
        'Custom number': 5,
        'To page breaks': 6
      },

      margins: {
        'Normal': [0.75, 0.75, 0.7, 0.7], // dafault
        'Narrow': [0.75, 0.75, 0.25, 0.25],
        'Wide':   [1, 1, 1, 1]
      },

      horizontal_alignment: {
        'Left': 1, // default
        'Center': 2,
        'Rigth': 3
      },

      vertical_alignment: {
        'Top': 1, // default
        'Bottom': 3,
        'Center': 2       
      }
    };

  var getCustomOption_ = function(obj, key) {
    var result = obj[key];
    if (!result) {
      // default option
      var key0 = Object.keys(obj)[0];
      result = obj[key0];
    }
    return result;
  }
  var i_horizontal_alignment = getCustomOption_(
    custom_pdf_options.horizontal_alignment,
    sets.horizontal_alignment);
  // console.log(i_horizontal_alignment);
  var i_vertical_alignment = getCustomOption_(
    custom_pdf_options.vertical_alignment,
    sets.vertical_alignment);
  // console.log(i_vertical_alignment);
  var i_page_size = getCustomOption_(
    custom_pdf_options.page_size,
    sets.page_size);
  var i_scale = getCustomOption_(
    custom_pdf_options.scale,
    sets.scale);
  // console.log(i_scale)
  var i_margins = getCustomOption_(
    custom_pdf_options.margins,
    sets.margins);
  // console.log(i_margins);


  // page size is custom
  if (
    sets.page_size === 'Custom size' &&
    sets.height && sets.height > 0 &&
    sets.width && sets.width > 0 ) {
      // W x H
      i_page_size = sets.width + 'x' + sets.height;
    }


  /** try to toast message to show users */
  var toastme_ = function(msg) {
    try {
      book.toast(msg);
    } catch(err) {
      console.log(msg);
    }
  }
  /** 5️⃣ Measure a range if needed */
  /** 
   * get size of a range 
   * 
   * */
  var getRangeSize_ = function(sheet, rownum, columnnum, rownum2, columnnum2) {
    var ratio = 96; // get inch from pixel
    options = {
      // max size allowed
      size_limit: 1100,  // rows/columns (in tests)
      // max size to measure
      measure_limit: 150 // rows/columns
    }
   toastme_('Please wait...', '📐Measuring Range...');

    
    // get width in pixels 
    var w = 0, size;
    for (var i = columnnum; i <= columnnum2; i++) {
      // console.log(i);
      if (i <= options.measure_limit) {
        size = sheet.getColumnWidth(i);
      }
      w += size;
      if ((i % 50) === 0 && i <= options.measure_limit) {
        toastme_(
          'Done ' + i + ' columns of ' + columnnum2,
          '↔📐Measuring width...');
      }
    }
    if (i > options.measure_limit) {
      toastme_(
        'Estimation: all other columns are the same size',
        '↔📐Measuring width...');
    }
  
    // get row height in pixels
    var h = 0;
    for (var i = rownum; i <= rownum2; i++) {
      if (i <= options.measure_limit) {
        size = sheet.getRowHeight(i);
      }
      h += size
      /** manual correction */
      if (size === 2) {
        h-=1;
      } else {
        // h -= 0.42; /** TODO → test the range to make it fit any range */
      }
      
      if ((i % 50) === 0 &&  i <= options.measure_limit) {
        toastme_(
          'Done ' + i + ' rows of ' + rownum2,
          '↕📐Measuring height...');
      }
    }
    if (i > options.measure_limit) {
      toastme_(
        'Estimation: all other rows are the same size',
        '↕📐Measuring height...');
    }

    // add 0.1 inch to fit some ranges
    return {
      // sheetName: sheet.getName(), // test
      width: w/ratio + 0.1,
      height: h/ratio + 0.1
    };
  }
  if (
    sets.range_a1 !== '' &&
    sets.page_size === '📐Fit to range') {
    var sizes = []
    for (var i = 0; i < ranges.length; i++) {
      sizes.push(getRangeSize_(
        sheets[i],
        rownums[i],
        columnnums[i],
        rownums2[i],
        columnnums2[i]));
    }
    logJson_(sizes);
    // put ranges vertically
    var h_i = 0, hhh = 0, w_i = 0, www = 0;
    for (var i = 0; i < sizes.length; i++) {
      w_i = sizes[i].width;
      if (w_i > www) {
        www = w_i;
      }
      h_i = sizes[i].height;
      if (h_i > hhh) {
        hhh = h_i;
      }
    }
    var s_www = Math.round(www * 1000) / 1000;
    var s_hhh = Math.round(hhh * 1000) / 1000;
    i_page_size = s_www + 'x' + s_hhh;
  }
  // console.log(i_page_size);

  /** 6️⃣ Other */
  // scale percent
  var i_scale_percent = null;
  // is custom scale
  if (i_scale === 5) {
    i_scale_percent = 1; // defautlt = 100%
    if (sets.scale_percent && sets.scale_percent !== '') {
      i_scale_percent = sets.scale_percent
    }
  }
  // console.log(i_scale_percent);
  if (
    sets.margins === 'Custom numbers'&&
    sets.margins_top >= 0 &&
    sets.margins_bottom >= 0 &&
    sets.margins_left >= 0 &&
    sets.margins_rigth >= 0) {
      i_margins = [
        sets.margins_top,
        sets.margins_bottom,
        sets.margins_left,
        sets.margins_rigth
      ];
    }


  /** 7️⃣ */
  var i_line_breaks = null;
  if (
    (sets.page_breaks_rows && sets.page_breaks_rows !== '') || 
    (sets.page_breaks_columns && sets.page_breaks_columns !== '')) {
      // [
      //   [
      //     "0", // sheet id
      //     // horizontal line breaks
      //     [ [ 0, 24 ],
      //       [ 32, 32 ] ],
      //     // vertical line breaks
      //     [ [ 1, 1 ] ]
      //   ]
      // ]
      i_line_breaks = [];
      var i_break_rows = [], i_break_columns = [];
      var getBreakValue_ = function(breakvalue) {
            var intval = parseInt(breakvalue);
            return [intval, intval]; }
      if (sets.page_breaks_rows && sets.page_breaks_rows !== '') {
        i_break_rows = sets.page_breaks_rows.toString().split(sets.delimiter)
          .map(getBreakValue_);
      }
      if (sets.page_breaks_columns && sets.page_breaks_columns !== '') {
        i_break_columns = sets.page_breaks_columns.toString().split(sets.delimiter)
          .map(getBreakValue_); 
      }
      // console.log(i_break_rows);
      // console.log(i_break_columns);
      for (var i = 0; i < i_sheets.length; i++) {
        i_line_breaks.push(
          [i_sheets[i][0],
          i_break_rows,
          i_break_columns]);
      }
    }
  // logJson_(i_line_breaks)


  /** 8️⃣ Combine the object  */
  var data_out = [
    null,null,null,null,null,null,null,null,null,0, /** TODO */
    i_sheets,
    10000000,null,null,null,null,null,null,null,null,null,null,null,null,null,null, /** TODO */
    i_date,
    null,
    null,
    [
      i_notes,
      null,    /** TODO */
      i_gridlines,
      i_page_numbers, 
      i_workbook,
      i_sheet,
      i_show_date,
      i_show_time,
      i_freezed_rows,   
      i_freezed_columns,    
      i_page_order,  /** TODO : what is the difference? */
      i_show_colontitles,
      i_colontitles_top,
      i_colontitles_bottom,
      i_horizontal_alignment,
      i_vertical_alignment
    ],
    [
      i_page_size,
      i_page_oriantation,  
      i_scale, 
      i_scale_percent, 
      i_margins
    ],
    null, /** TODO */
    0,    /** TODO */
    i_line_breaks
  ];

  // logJson_(datajson);
  return data_out;
}

