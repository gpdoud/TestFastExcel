using FastExcel;

var newFile = "/Users/gpdoud/Repos/TestFastExcel/NewExcel.xlsx";
if (System.IO.File.Exists(newFile))
    System.IO.File.Delete(newFile);

using FastExcel.FastExcel fe = new FastExcel.FastExcel(new FileInfo("/Users/gpdoud/Repos/TestFastExcel/blank.xlsx"), new FileInfo("/Users/gpdoud/Repos/TestFastExcel/NewExcel.xlsx"));
var worksheet = new Worksheet();
var rows = new List<Row>();
for(var row = 1; row <= 100; row++) {
    var cells = new List<Cell>();
    for(var cell = 1; cell <= 15; cell++) {
        cells.Add(new Cell(cell, cell * DateTime.Now.Millisecond));
    }
    cells.Add(new Cell(16, "FileFormat"+row));

    rows.Add(new Row(row, cells));
}
worksheet.Rows = rows;

fe.Write(worksheet, "sheet1");
Console.WriteLine("Done ...");