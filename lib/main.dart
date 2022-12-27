import 'dart:convert';
import 'dart:developer';

import 'package:excel/excel.dart';
import 'dart:io';

import 'package:opening_excel/keys.dart';

void main() async {
  // Read the Excel file
  var bytes =
      File('/Users/abdullah/Desktop/waste_me/opening_excel/lib/FRENCH.xlsx')
          .readAsBytesSync();
  var excel = Excel.decodeBytes(bytes);

  // Get the first sheet
  var sheet = excel.tables.values.first;

  // Extract the data as a list of rows and cells
  var rows = sheet.rows;

  // Create a StringBuffer to hold the data for the Dart file
  // var buffer = StringBuffer();

  /* 
    Handle the following exceptions.
    1. Spaces and '(' ')' should be replaced with '_'
    2. Remove &,:!?./\-' and 1234567890 and leading space
  
  
  
   */

  Map<String, String> lowercaseMap = constantsMap.map((key, value) {
    var k = key.toLowerCase();
    var v = value;

    log('keyAndValue: $k - $v');

    return MapEntry(k, v);
  });

  var data = {};
  int totalRows = 0;

  // Iterate over the rows and cells and append the data to the buffer
  for (var row in rows) {
    if (row.first?.value != null &&
        row.last?.value != null &&
        row.last?.value.toString().trim() != '') {
      totalRows++;
      String? k = row.first?.value.toString();
      if (k != null) {
        // first remove special characters
        k = k.replaceAll('&', '');
        k = k.replaceAll(',', '');
        k = k.replaceAll(':', '');
        k = k.replaceAll('!', '');
        k = k.replaceAll('?', '');
        k = k.replaceAll('.', '');
        k = k.replaceAll('/', '');
        k = k.replaceAll('\\', '');
        k = k.replaceAll('-', '');
        k = k.replaceAll('\'', '');
        k = k.replaceAll('[', '');
        k = k.replaceAll(']', '');
        k = k.replaceAll('’', '');

        // remove any leading or trailing space
        k = k.trim();

        k = k.replaceAll(' ', '_');
        k = k.replaceAll('(', '_');
        k = k.replaceAll(')', '_');

        String? actualKey = lowercaseMap[k.toLowerCase()];

        log('${k.toLowerCase()} = $actualKey');
        if (actualKey != null) {
          data[actualKey] = row.last?.value.toString();
        } else {
          data['NOT_FOUND'] = row.last?.value.toString();
        }
      }
    }

    // if (row.first?.value != null && row.last?.value != null) {
    //   data[row.first?.value.toString()] = row.last?.value.toString();
    // }
  }

  var json = jsonEncode(data);

  // Write the buffer to a Dart file
  File('/Users/abdullah/Desktop/waste_me/opening_excel/lib/FRENCH.json')
      .writeAsStringSync(json);

  log('total rows: $totalRows');
}
