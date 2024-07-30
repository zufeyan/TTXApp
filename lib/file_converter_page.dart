import 'dart:async';
import 'dart:io';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';

class FileConverterPage extends StatefulWidget {
  @override
  _FileConverterPageState createState() => _FileConverterPageState();
}

class _FileConverterPageState extends State<FileConverterPage> {
  String? txtFilePath;
  String? xlsxFilePath;
  bool isConverting = false;
  double progress = 0.0;
  Future<void>? conversionFuture;

  Future<void> pickTxtFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['txt'],
    );

    if (result != null) {
      setState(() {
        txtFilePath = result.files.single.path;
      });
    }
  }

  Future<void> pickXlsxFilePath() async {
    String? outputPath;
    try {
      String? selectedDirectory = await FilePicker.platform.getDirectoryPath();

      if (selectedDirectory != null) {
        outputPath = '$selectedDirectory/ConvertedFile.xlsx';
        setState(() {
          xlsxFilePath = outputPath;
        });
      }
    } catch (e) {
      print("Failed to get output path: $e");
    }
  }

  Future<void> startConversionProcess() async {
    if (txtFilePath == null || xlsxFilePath == null) return;

    setState(() {
      isConverting = true;
      progress = 0.0;
      conversionFuture = convertFile(txtFilePath!, xlsxFilePath!);
    });

    try {
      await conversionFuture;
      showAlertDialog('Success', 'File converted successfully');
    } catch (e) {
      showAlertDialog('Error', 'File conversion failed: $e');
    } finally {
      setState(() {
        isConverting = false;
      });
    }
  }

  Future<void> convertFile(String txtPath, String xlsxPath) async {
    try {
      final file = File(txtPath);
      final lines = await file.readAsLines();
      final totalLines = lines.length;

      var excel = Excel.createExcel();
      Sheet sheetObject = excel['Sheet1'];

      List<List<dynamic>> data = [];
      for (int i = 0; i < lines.length; i++) {
        String line = lines[i].trim();
        if (line.isEmpty) continue;

        List<dynamic> rowData = [];
        int spaceCount = ' '.allMatches(line).length;
        List<String> columns = [];

        if (spaceCount == 1) {
          rowData.add(" ");
        } else if (spaceCount == 2) {
          int firstSpaceIndex = line.indexOf(' ');
          String firstPart = line.substring(0, firstSpaceIndex);
          String secondPart = line.substring(firstSpaceIndex + 1);
          rowData.add(firstPart);
          columns = secondPart.split(',');
        } else {
          int secondSpaceIndex = line.indexOf(' ', line.indexOf(' ') + 1);
          String firstPart = line.substring(0, secondSpaceIndex);
          String secondPart = line.substring(secondSpaceIndex + 1);
          rowData.add(firstPart);
          columns = secondPart.split(',');
        }

        int x = 0;
        String date_ = " ";
        String time_ = " ";
        double kw = 0;

        while (x < columns.length) {
          if (x == 0) {
            rowData.add(columns[x]);
          } else if (x == 1) {
            String datetimeStr = columns[x];
            List<String> datetimeParts = datetimeStr.split(" ");
            if (datetimeParts.length == 2) {
              String datePart = datetimeParts[0];
              String timePart = datetimeParts[1];

              DateTime date;
              try {
                date = DateTime.parse(datePart);
              } catch (e) {
                showAlertDialog('Error', 'Invalid date format in line: $line');
                return;
              }
              String newDateStr =
                  "${date.day.toString().padLeft(2, '0')}/${date.month.toString().padLeft(2, '0')}/${date.year}";

              date_ = newDateStr;
              time_ = timePart;
              rowData.add(newDateStr);
              rowData.add(timePart);
            } else {
              showAlertDialog(
                  'Error', 'Invalid datetime format in line: $line');
              return;
            }
          } else if (x == 6) {
            kw = int.tryParse(columns[x]) != null
                ? int.parse(columns[x]) / 1000
                : 0;
            rowData.add(kw);
          } else if (x == 7) {
            rowData.add(kw / 4);
            rowData.add("$date_ $time_");
          }
          x++;
        }

        data.add(rowData);
        await Future.delayed(Duration(microseconds: 100)); // Add delay here
        setState(() {
          progress = (i + 1) / totalLines;
        });
      }

      List<String> headers = [
        "ยี่ห้อ",
        "PEA",
        "ว/ด/ป",
        "เวลา",
        "kw",
        "kwh",
        "วัน-เวลา"
      ];
      data.insert(0, headers);
      for (var row in data) {
        sheetObject.appendRow(row);
      }

      final excelFile = File(xlsxPath)
        ..createSync(recursive: true)
        ..writeAsBytesSync(excel.save()!);
    } catch (e) {
      throw e;
    }
  }

  void showAlertDialog(String title, String message) {
    showDialog(
      context: context,
      builder: (BuildContext context) {
        return AlertDialog(
          title: Text(title),
          content: Text(message),
          actions: [
            TextButton(
              child: Text("OK"),
              onPressed: () {
                Navigator.of(context).pop();
              },
            ),
          ],
        );
      },
    );
  }

  void clearData() {
    setState(() {
      txtFilePath = null;
      xlsxFilePath = null;
      progress = 0.0;
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Row(
          children: [
            Image.asset(
              'assets/icon.png', // Path to your icon file in assets
              height: 30,
            ),
            SizedBox(width: 10),
            Text('Text to XLSX Converter'),
          ],
        ),
        backgroundColor: Colors.deepPurple, // สีใหม่สำหรับ AppBar
        actions: [
          Builder(
            builder: (context) => IconButton(
              icon: Icon(Icons.menu),
              onPressed: () => Scaffold.of(context).openEndDrawer(),
            ),
          ),
        ],
      ),
      endDrawer: Drawer(
        child: ListView(
          padding: EdgeInsets.zero,
          children: <Widget>[
            Container(
              height: 100, // ลดความสูงของ DrawerHeader
              child: DrawerHeader(
                decoration: BoxDecoration(
                  color: Colors
                      .deepPurple, // สีใหม่สำหรับพื้นหลังของเมนูแฮมเบอร์เกอร์
                ),
                child: Text(
                  'Menu',
                  style: TextStyle(
                    color: Colors.white,
                    fontSize: 24,
                  ),
                ),
              ),
            ),
            Container(
              color: Colors.deepPurple[300], // สีพื้นหลังสำหรับปุ่ม
              child: Column(
                children: [
                  ListTile(
                    leading: Icon(Icons.clear_all, color: Colors.white),
                    title: Text('Clear path file',
                        style: TextStyle(color: Colors.white)),
                    onTap: () {
                      clearData();
                      Navigator.of(context).pop();
                    },
                  ),
                  ListTile(
                    leading: Icon(Icons.exit_to_app, color: Colors.white),
                    title: Text('Exit', style: TextStyle(color: Colors.white)),
                    onTap: () {
                      exit(0);
                    },
                  ),
                ],
              ),
            ),
          ],
        ),
      ),
      body: Stack(
        children: [
          Center(
            child: Opacity(
              opacity: 0.1,
              child: Image.asset(
                'assets/logo.png', // Path to your logo file in assets
                fit: BoxFit.cover,
              ),
            ),
          ),
          SingleChildScrollView(
            padding: const EdgeInsets.all(16.0),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'Select TXT File:',
                  style:
                      Theme.of(context).textTheme.headlineSmall, // updated here
                ),
                Row(
                  children: [
                    Expanded(
                      child: TextField(
                        readOnly: true,
                        style: TextStyle(color: Colors.deepPurple),
                        decoration: InputDecoration(
                          hintText: txtFilePath ?? '',
                        ),
                      ),
                    ),
                    SizedBox(width: 10),
                    ElevatedButton(
                      onPressed: pickTxtFile,
                      child: Text('Browse TXT File'),
                    ),
                  ],
                ),
                SizedBox(height: 20),
                Text(
                  'Save XLSX File:',
                  style:
                      Theme.of(context).textTheme.headlineSmall, // updated here
                ),
                Row(
                  children: [
                    Expanded(
                      child: TextField(
                        readOnly: true,
                        style: TextStyle(color: Colors.deepPurple),
                        decoration: InputDecoration(
                          hintText: xlsxFilePath ?? '',
                        ),
                      ),
                    ),
                    SizedBox(width: 10),
                    ElevatedButton(
                      onPressed: pickXlsxFilePath,
                      child: Text('Select Directory'),
                    ),
                  ],
                ),
                SizedBox(height: 20),
                Visibility(
                  visible: txtFilePath != null && xlsxFilePath != null,
                  child: Center(
                    child: GestureDetector(
                      onTap: startConversionProcess,
                      child: Column(
                        children: [
                          Image.asset(
                            'assets/icon_data.png', // Path to the icon image
                            height: 100, // Adjust the size as needed
                          ),
                          SizedBox(height: 10),
                          Text(
                            'Start Conversion Process',
                            style: TextStyle(
                              color: Colors.deepPurple,
                              fontSize: 16,
                              fontWeight: FontWeight.bold,
                            ),
                          ),
                        ],
                      ),
                    ),
                  ),
                ),
                SizedBox(height: 20),
                if (isConverting)
                  Center(
                    child: Column(
                      children: [
                        LinearProgressIndicator(value: progress),
                        SizedBox(height: 20),
                        Text(
                          '${(progress * 100).toStringAsFixed(1)}%',
                          style: Theme.of(context)
                              .textTheme
                              .bodyMedium, // updated here
                        ),
                        SizedBox(height: 20),
                        Text(
                          'Converting...',
                          style: Theme.of(context)
                              .textTheme
                              .bodyMedium, // updated here
                        ),
                      ],
                    ),
                  ),
              ],
            ),
          ),
        ],
      ),
    );
  }
}
