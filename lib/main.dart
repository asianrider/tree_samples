import 'dart:io';

import 'package:csv/csv.dart';
import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:pluto_menu_bar/pluto_menu_bar.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData(
        // This is the theme of your application.
        //
        // Try running your application with "flutter run". You'll see the
        // application has a blue toolbar. Then, without quitting the app, try
        // changing the primarySwatch below to Colors.green and then invoke
        // "hot reload" (press "r" in the console where you ran "flutter run",
        // or simply save your changes to "hot reload" in a Flutter IDE).
        // Notice that the counter didn't reset back to zero; the application
        // is not restarted.
        primarySwatch: Colors.blue,
      ),
      home: const MyHomePage(title: 'Tree samples report'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  // This widget is the home page of your application. It is stateful, meaning
  // that it has a State object (defined below) that contains fields that affect
  // how it looks.

  // This class is the configuration for the state. It holds the values (in this
  // case the title) provided by the parent (in this case the App widget) and
  // used by the build method of the State. Fields in a Widget subclass are
  // always marked "final".

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class CsvTable extends DataTableSource {
  List<List<dynamic>> csvData;
  bool firstLineIsHeader;
  CsvTable(this.csvData, {this.firstLineIsHeader = true});

  @override
  DataRow? getRow(int index) {
    int i = firstLineIsHeader ? index + 1 : index;
    return DataRow(cells: csvData[i].map((cell) => DataCell(Text(cell.toString()))).toList());
  }

  @override
  bool get isRowCountApproximate => false;

  @override
  int get rowCount => csvData.length;

  @override
  int get selectedRowCount => 0;
}

class _MyHomePageState extends State<MyHomePage> {
  int _counter = 0;

  List<List>? csvData;

  bool isLoading = false;

  void message(context, String text) {
    ScaffoldMessenger.of(context).hideCurrentSnackBar();

    final snackBar = SnackBar(
      content: Text(text),
    );

    ScaffoldMessenger.of(context).showSnackBar(snackBar);
  }

  @override
  Widget build(BuildContext context) {
    print("OK build");
    // This method is rerun every time setState is called, for instance as done
    // by the _incrementCounter method above.
    //
    // The Flutter framework has been optimized to make rerunning build methods
    // fast, so that you can just rebuild anything that needs updating rather
    // than having to individually change instances of widgets.
    return Scaffold(
      // appBar: AppBar(
      //   title: Text(widget.title),
      // ),
      body: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
        PlutoMenuBar(
            backgroundColor: Colors.blue,
            itemStyle: const PlutoMenuItemStyle(
              activatedColor: Colors.white,
              indicatorColor: Colors.deepOrange,
              textStyle: TextStyle(color: Colors.white),
              iconColor: Colors.white,
              moreIconColor: Colors.white,
            ),
            menus: [
              PlutoMenuItem(title: 'Fichier', children: [
                PlutoMenuItem(
                  title: 'Ouvrir..',
                  // icon: Icons.import_export,
                  onTap: () async {
                    FilePickerResult? result = await FilePicker.platform.pickFiles(
                      type: FileType.custom,
                      allowedExtensions: ['txt', 'csv', 'xlsx', 'xlsm'],
                    );
                    if (result != null) {
                      loadAndProcessFile(result.files.first);
                    }
                    // processCsv();
                  },
                ),
                PlutoMenuItem(
                    title: 'Entegistrer..',
                    // icon: Icons.import_export,
                    onTap: () async {
                      final path = await FilePicker.platform.saveFile(
                        allowedExtensions: ['txt', 'csv', 'xlsx', 'xlsm'],
                      );
                      if (path != null) {
                        saveResults(path);
                      }
                    })
              ]),
            ]),
        Expanded(
            child: Stack(children: [
          if (isLoading) const Center(child: CircularProgressIndicator()),
          SingleChildScrollView(
            child: csvData == null
                ? const Center(child: Text('Chargez des données avec "Fichier->ouvrir"'))
                : PaginatedDataTable(
                    columns: csvData![0]
                        .map(
                          (item) => DataColumn(
                            label: Text(
                              item.toString(),
                            ),
                          ),
                        )
                        .toList(),
                    source: CsvTable(csvData!),
                  ),
          )
        ]))
      ]),
      floatingActionButton: FloatingActionButton(
        onPressed: () {
          filterRowsSync();
        },
        tooltip: 'Increment',
        child: const Text("filtrer"),
      ), // This trailing comma makes auto-formatting nicer for build methods.
    );
  }

  void processCsv() async {
    setState(() => isLoading = true);
    final result = await DefaultAssetBundle.of(context).loadString(
      "assets/test.csv",
    );
    csvData = const CsvToListConverter().convert(result, eol: "\n", shouldParseNumbers: true);
    setState(() => isLoading = false);
  }

  void loadAndProcessFile(PlatformFile file) async {
    if (file.extension == 'txt' || file.extension == 'csv') {
      setState(() => isLoading = true);
      final contents = await File(file.path!).readAsString();
      csvData = const CsvToListConverter().convert(contents, fieldDelimiter: ' ', eol: "\n", shouldParseNumbers: true);
      setState(() => isLoading = false);
    } else if (file.extension == 'xlsx' || file.extension == 'xlsm') {
      setState(() => isLoading = true);
      try {
        final contents = await File(file.path!).readAsBytes();
        final excel = Excel.decodeBytes(contents);
        final tables = excel.tables.values;
        if (tables.length == 0) {
          message(context, "Fichier vide");
          return;
        }
        final table = tables.first;
        final result = table.rows
            .map((row) => row.map((e) {
                  final text = e?.value.toString();
                  return text;
                }).toList())
            .toList();
        setState(() {
          csvData = result;
          setState(() => isLoading = false);
        });
      } catch (e) {
        message(context, e.toString());
        setState(() => isLoading = false);
      }
    }
  }

  void filterRowsSync() {
    print("Do filter...");
    setState(() {
      isLoading = true;
    });
    final result = doFilterRows(csvData!);
    if (result is String) {
      print(result);
      message(context, result);
    } else {
      print("Résultat: ${result.length} lignes");
      message(context, "Résultat: ${result.length} lignes");
      setState(() {
        csvData = [csvData![0], ...result];
      });
    }
    setState(() {
      isLoading = false;
    });
  }

  Future filterRows() async {
    print("Spawn filter...");
    setState(() {
      isLoading = true;
    });
    final result = await compute(doFilterRows, csvData!);
    if (result is String) {
      print(result);
    } else {
      print("Got result: ${result.length} lines");
      setState(() {
        csvData = [csvData![0], ...result];
      });
    }
    setState(() {
      isLoading = false;
    });
  }

  saveResults(String fileName) async {
    final saveFile = File(fileName);
    // final sink = saveFile.openWrite();
    // sink.write(ListToCsvConverter().convert(csvData));
    // await sink.flush();
    // await sink.close();
    saveFile.writeAsString(const ListToCsvConverter().convert(csvData));
  }
}

dynamic doFilterRows(List<List<dynamic>> csvData) {
  try {
    final firstRow = csvData![0];
    int targetColumn = firstRow.indexOf('%CC');
    if (targetColumn == -1) targetColumn = firstRow.indexOf('Glk');
    final dateL = firstRow.indexOf('DateL');
    final dateR = firstRow.indexOf('DateR');
    if (targetColumn == -1) {
      // message(context, "Can't find %CC or Glk column");
      return "Pas trouvé de colonne %CC ou Glk";
    }
    if (dateL == -1) {
      // message(context, "Can't find DateL");
      return "Pas trouvé de colonne DateL";
    }
    if (dateR == -1) {
      // message(context, "Can't find DateR");
      return "Pas trouvé de colonne DateR";
    }
    print("Target CC is $targetColumn, year columns are $dateL, $dateR");
    Map<String, List<List<dynamic>>> treeGroups = {};
    List<List<dynamic>> result = [];
    csvData.sublist(1).forEach((row) {
      if (row.length > 0 && row[0] != null) {
        final treeName = row[0] as String;
        if (treeGroups[treeName] == null) treeGroups[treeName] = [];
        treeGroups[treeName]!.add(row);
      }
    });
    print("Found ${treeGroups.keys.length} trees");
    treeGroups.forEach((key, rowList) {
      try {
        final ownLine = rowList.firstWhere((row) => row[targetColumn] == 100 || row[targetColumn] == "100");
        final start = ownLine[dateL];
        final end = ownLine[dateR];
        print("Tree: $key: start=$start, end=$end, has ${rowList.length} entries");
        rowList.forEach((row) {
          if (row[dateL] == start && row[dateR] == end) {
            result.add(row);
          }
        });
      } catch (e) {
        // message(context, "Can't find 100 in column $targetColumn of tree: $key");
        return;
      }
    });
    // message(context, "Trouvé ${result.length} résultats");
    return result;
  } catch (e) {
    return e.toString();
  }
}