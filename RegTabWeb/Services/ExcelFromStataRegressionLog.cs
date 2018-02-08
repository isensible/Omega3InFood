using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace RegTabWeb.Services
{
    public interface IExcelFromStataRegressionLog
    {
        MemoryStream GenerateExcel(string fileContent);
    }

    public class ExcelFromStataRegressionLog : IExcelFromStataRegressionLog
    {
        private readonly ILogger<ExcelFromStataRegressionLog> _logger;
        private string[] _inputLines;
        private SubArray<string>[] _tables;
        private SubArray<string>[] _commands;
        private IOrderedEnumerable<SubArray<string>> _orderedCommandsAndTables;
        private readonly string HeaderOddsRatioLowHigh = "OR (95% CI)";
        private readonly string HeaderPValue = "P";
        private SubArray<string>[] _numberOfObs;

        public ExcelFromStataRegressionLog(ILogger<ExcelFromStataRegressionLog> logger)
        {
            _logger = logger;
        }

        public MemoryStream GenerateExcel(string fileContent)
        {
            var options = new Options
            {
                Input = fileContent
            };

            return PutRegressionTablesIntoExcel(options);
        }
        
        private MemoryStream PutRegressionTablesIntoExcel(Options options)
        {
            PrepareInput(options);

            var stream = new MemoryStream();
            
            WriteStataLogTablesToExcelUsingStream(stream);

            return stream;
        }

        private void PrepareInput(Options options)
        {
            _inputLines = options.Input.Split(
                new[] {"\r\n", "\r", "\n"},
                StringSplitOptions.None);

            var numberedLines = NumberTheLines();

            LocateStataLogTables(numberedLines);

            LocateStataCommands(numberedLines);
            
            LocateStataStats(numberedLines);

            _orderedCommandsAndTables = _commands
                .Concat(_numberOfObs)
                .Concat(_tables)
                .OrderBy(x => x.Offset);

            if (options.List)
            {
                foreach (var t in _orderedCommandsAndTables)
                {
                    PrintSubArray(t);
                }
            }
        }

        private void WriteStataLogTablesToExcelUsingStream(Stream stream)
        {
            using (var excel = new ExcelPackage(stream))
            {
                var worksheet = excel.Workbook.Worksheets.Add("Regression_Tables");
                
                worksheet.Cells.Style.Font.Name = "Times New Roman";
                worksheet.Cells.Style.Font.Size = 12;
                
                var colA = worksheet.Column(1);
                colA.Width = 120;
                colA.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                
                var colB = worksheet.Column(2);
                colB.Width = 50;
                colB.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                
                var colC = worksheet.Column(3);
                colC.Width = 40;
                colC.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                
                var currentAddress = new ExcelAddress("A1");
                const string col = "A";

                ExcelAddress NextAddress(string column, int lastRow, int gapSize = 0) => new ExcelAddress(column + (lastRow + gapSize + 1));

                foreach (var o in _orderedCommandsAndTables)
                {
                    var command = _commands.FirstOrDefault(x => x.Offset == o.Offset);
                    if (command != null)
                    {
                        var range = AppendStataCommandToWorksheet(worksheet, currentAddress, command);
                        range.Style.Font.Bold = true;
                        range.Style.WrapText = true;
                        currentAddress = NextAddress(col, range.End.Row);
                        continue;
                    }

                    var obs = _numberOfObs.FirstOrDefault(x => x.Offset == o.Offset);
                    if (obs != null)
                    {
                        var split = obs[0]
                            .Split("=", StringSplitOptions.RemoveEmptyEntries)
                            .Select(s => s.Trim())
                            .ToArray();

                        worksheet.Cells[currentAddress.Start.Row, currentAddress.Start.Column].Value = split[0];
                        worksheet.Cells[currentAddress.Start.Row, currentAddress.Start.Column+1].Value = split[1];
                        currentAddress = NextAddress(col, currentAddress.End.Row);
                    }
                    
                    var table = _tables.FirstOrDefault(x => x.Offset == o.Offset);
                    if (table != null)
                    {
                        var range = AppendRegressionTableToWorksheet(worksheet, currentAddress, table);
                        range.AutoFilter = false;
                        range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        
                        currentAddress = NextAddress(col, range.End.Row, gapSize:1);
                    }
                }
                
                excel.Save();
            }
        }

        private ExcelRangeBase AppendRegressionTableToWorksheet(ExcelWorksheet worksheet, ExcelAddress address, SubArray<string> table)
        {
            return worksheet.Cells[address.Address].LoadFromDataTable(BuildRegressionDataTable(table, tableName:address.Address), true, TableStyles.None);
        }

        private ExcelRangeBase AppendStataCommandToWorksheet(ExcelWorksheet worksheet, ExcelAddress address, SubArray<string> command)
        {
            var text = string.Join(Environment.NewLine, command);

            worksheet.SetValue(address.Address, text);

            var cell = worksheet.Cells[address.Address];

            cell.AutoFilter = false;
            cell.AutoFitColumns();
            
            return cell;
        }

        private NumberedLine[] NumberTheLines()
        {
            return _inputLines
                .Select((l, i) => new NumberedLine {Line = l.Trim(), LineNumber = i})
                .Skip(1)
                .SkipLastN(1)
                .ToArray();
        }
        
        private DataTable BuildRegressionDataTable(SubArray<string> table, string tableName = "RegressionTable")
        {
            var tokens = StataLogTableTokenizer(table);
            
            var headerTokens = StataLogTableHeaderTokenizer(tokens.Header).ToArray();
            
            var dt = new DataTable(tableName);
            dt.Columns.Add(headerTokens[0], typeof(string));
            dt.Columns.Add(HeaderOddsRatioLowHigh, typeof(string));
            dt.Columns.Add(HeaderPValue, typeof(string));
            
            foreach (var row in tokens.Rows)
            {
                var data = StataLogTableRowTokenizer(row).ToArray();
                var dtRow = dt.NewRow();
                dtRow[headerTokens[0]] = data[0];
                dtRow[HeaderOddsRatioLowHigh] = 
                    string.Format(@"{0:0.00} ({1:0.00}, {2:0.00})",
                        double.Parse(data[1]),
                        double.Parse(data[5]),
                        double.Parse(data[6]));
                dtRow[HeaderPValue] = data[4] == "0.000" ? "<0.001" : data[4];
                dt.Rows.Add(dtRow);
            }
            
            return dt;
        }

        
        private StringTable StataLogTableTokenizer(SubArray<string> table)
        {
            return new StringTable
            {
                Header = table[1],
                Rows = new SubArray<string>(_inputLines, table.Offset + 3, table.Count - 4)
            };
        }

        private static IEnumerable<string> StataLogTableHeaderTokenizer(string header)
        {
            var name = new string(header.TakeWhile(c => c != '|').ToArray());
            yield return name.Trim();
            yield return "Odds Ratio";
            yield return "Std. Err.";
            yield return "z";
            yield return "P>|z|";
            yield return "95% Conf. Low";
            yield return "95% Conf. High";
        }

        private static IEnumerable<string> StataLogTableRowTokenizer(string row)
        {
            var nameSplit = row.Split("|");

            yield return nameSplit[0].Trim();

            var dataSplit = nameSplit[1].Split(" ", StringSplitOptions.RemoveEmptyEntries);

            foreach (var column in dataSplit)
                yield return column;
        }

        private void LocateStataStats(NumberedLine[] numberedLines)
        {
            const string numberOfObs = "Number of obs";

            _numberOfObs = numberedLines
                .Where(a => a.Line.Contains(numberOfObs))
                .OrderBy(a => a.LineNumber)
                .Select(a => new SubArray<string>(_inputLines, a.LineNumber, 1))
                .ToArray();
        }

        private void LocateStataCommands(NumberedLine[] numberedLines)
        {
            var commandStrings = new []
            {
                ". xi:clogit", ". xi:logit"
            };
            
            var commandContinuationString = ">";
            
            var withIndices = numberedLines
                .Where(a => commandStrings.Any(s => a.Line.StartsWith(s)) || a.Line.StartsWith(commandContinuationString))
                .OrderBy(a => a.LineNumber)
                .ToArray();

            var list = new List<Tuple<int, int>>();
            
            var firstLine = -1;
            var lastLine = -1;
            bool expectingAnotherLine = false;
            foreach (var c in withIndices)
            {
                if (commandStrings.Any(s => c.Line.StartsWith(s)))
                {
                    firstLine = c.LineNumber;
                    lastLine = c.LineNumber;
                    expectingAnotherLine = c.Line.EndsWith("///");
                }
                
                if (c.Line.StartsWith(commandContinuationString) && expectingAnotherLine)
                {
                    lastLine = c.LineNumber;
                    expectingAnotherLine = c.Line.EndsWith("///");
                }

                if (!expectingAnotherLine && firstLine >= 0)
                {
                    list.Add(new Tuple<int, int>(firstLine, lastLine));
                    firstLine = -1;
                    lastLine = -1;
                }
            }

            _commands = list.Select(SliceInput).ToArray();
        }

        private void LocateStataLogTables(NumberedLine[] numberedLines)
        {
            var tableLines = numberedLines
                .Where(a => { return a.Line.Length != 0 && a.Line.All(c => c.CompareTo('-') == 0); })
                .ToArray();

            var tableBorderLineIndices = tableLines
                .Select(a => a.LineNumber)
                .ToList();

            _tables = SplitList(tableBorderLineIndices)
                .Where(t => t.Count == 2)
                .Select(t => new Tuple<int, int>(t[0], t[1]))
                .Select(SliceInput)
                .ToArray();
        }
        
        private SubArray<string> SliceInput(Tuple<int, int> tableIndexTuple)
        {
            var offset = tableIndexTuple.Item1;
            var count = 1 + tableIndexTuple.Item2 - tableIndexTuple.Item1;
            return new SubArray<string>(_inputLines, offset, count);
        }

        private void PrintSubArray(SubArray<string> subArray)
        {
            foreach (var line in subArray)
                Console.WriteLine(line);
        }
        
        private static IEnumerable<List<T>> SplitList<T>(List<T> input, int size=2)  
        {        
            for (var i = 0; i < input.Count; i+= size) 
            { 
                yield return input.GetRange(i, Math.Min(size, input.Count - i)); 
            }  
        }
    }
}