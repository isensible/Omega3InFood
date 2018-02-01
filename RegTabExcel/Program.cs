﻿using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using CommandLine;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace RegTabExcel
{
    public class Options
    {
        [Option('i', "input", Required = true, HelpText = "Full path to input file")]
        public string Input { get; set; }

        [Option('o', "output", Required = false, HelpText = "Full path of output Excel file")]
        public string Output { get; set; }

        [Option('l', "list", Required = false, HelpText = "List the tables to console")]
        public bool List { get; set; }

        [Option('t', "table", Required = false, HelpText = "Index of the table to output")]
        public int Table { get; set; }
        
        [Option('r', "row", Required = false, HelpText = "Number of rows to output from selected table")]
        public int Rows { get; set; }
    }

    public class NumberedLine
    {
        public string Line { get; set; }
        public int LineNumber { get; set; }
    }

    public class StringTable
    {
        public string Header { get; set; }
        public SubArray<string> Rows { get; set; }
    }
    
    class Program
    {
        private static string[] _inputLines;
        private static SubArray<string>[] _tables;
        private static SubArray<string>[] _commands;
        private static IOrderedEnumerable<SubArray<string>> _orderedCommandsAndTables;

        static int Main(string[] args)
        {
            var result = Parser.Default.ParseArguments<Options>(args);
            
            return result.MapResult(PutRegressionTablesIntoExcel, errors => 1);
        }

        private static int PutRegressionTablesIntoExcel(Options options)
        {
            _inputLines = File.ReadAllLines(options.Input);

            var numberedLines = NumberTheLines();

            IsolateTables(numberedLines);

            IsolateCommands(numberedLines);

            _orderedCommandsAndTables = _commands
                .Concat(_tables)
                .OrderBy(x => x.Offset);

            if (options.List)
            {
                foreach (var t in _orderedCommandsAndTables)
                {
                    PrintSubArray(t);
                }
            }

            var newFile = new FileInfo(options.Output);
            
            if (newFile.Exists)
                newFile.Delete();
            
            PutExcel1(newFile);

            Console.ReadKey();

            return 0;
        }
        
        private static void PutExcel1(FileInfo newFile)
        {
            using (var excel = new ExcelPackage(newFile))
            {
                ExcelTextFormat f = new ExcelTextFormat
                {
                    TextQualifier = '"',
                    Delimiter = ','
                };

                var worksheet = excel.Workbook.Worksheets.Add("Regression_Tables");

                ExcelAddress currentAddress = new ExcelAddress("A1");
                string col = "A";
                ExcelRangeBase lastRange;

                ExcelAddress NextAddress(string column, int lastRow) => new ExcelAddress(column + (lastRow + 1));

                foreach (var o in _orderedCommandsAndTables)
                {
                    SubArray<string> command = _commands.FirstOrDefault(x => x.Offset == o.Offset);
                    if (command != null)
                    {
                        lastRange = PutCommandToWorksheet(worksheet, currentAddress, command);
                        currentAddress = NextAddress(col, lastRange.End.Row);
                        continue;
                    }
                    
                    SubArray<string> table = _tables.FirstOrDefault(x => x.Offset == o.Offset);
                    if (table != null)
                    {
                        lastRange = PutTableToWorksheet(worksheet, currentAddress, table);
                        currentAddress = NextAddress(col, lastRange.End.Row);
                    }
                }
                
                excel.Save();
            }
        }

        private static ExcelRangeBase PutTableToWorksheet(ExcelWorksheet worksheet, ExcelAddress address, SubArray<string> table)
        {
            var csv = new StringBuilder();
            var c = TableToCsv(table, escapeCsv: false, writeHeader: true);
            foreach (var line in c)
            {
                csv.AppendLine(line);
            }
            
            var f = new ExcelTextFormat
            {
                TextQualifier = '"',
                Delimiter = ','
            };
            
            return worksheet.Cells[address.Address].LoadFromText(csv.ToString(), f, TableStyles.Medium10, true);
        }

        private static ExcelRangeBase PutCommandToWorksheet(ExcelWorksheet worksheet, ExcelAddress address, SubArray<string> command)
        {
            var text = string.Join(Environment.NewLine, command);

            worksheet.SetValue(address.Address, text);

            var cell = worksheet.Cells[address.Address];

            cell.AutoFilter = false;
            cell.AutoFitColumns();
            
            return cell;
        }

        private static void PutExcel(FileInfo newFile)
        {
            using (var excel = new ExcelPackage(newFile))
            {
                ExcelTextFormat f = new ExcelTextFormat
                {
                    TextQualifier = '"',
                    Delimiter = ','
                };

                // do work here
                ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("Regression_Tables");

                bool writtenHeader = false;
                StringBuilder csv = new StringBuilder();

                foreach (var t in _tables)
                {
                    var c = TableToCsv(t, escapeCsv: false, writeHeader: writtenHeader == false);
                    writtenHeader = true;
                    foreach (var line in c)
                    {
                        csv.AppendLine(line);
                    }
                }
                worksheet.Cells.LoadFromText(csv.ToString(), f, TableStyles.Dark10, true);
                excel.Save();
            }
        }

        private static NumberedLine[] NumberTheLines()
        {
            return _inputLines
                .Select((l, i) => new NumberedLine {Line = l.Trim(), LineNumber = i})
                .Skip(1)
                .SkipLastN(1)
                .ToArray();
        }
        
        /*private static DataTable GetDataTable(SubArray<string> table, string tableName = "RegressionTable")
        {
            var tokens = TableTokenizer(table);
            
            var headerTokens = HeaderTokenizer(tokens.Header).ToArray();
            
            DataTable dt = new DataTable(tableName);
            dt.Columns.Add(headerTokens[0], typeof(string));
            dt.Columns.Add("Odds Ratio (Low, High)", typeof(string));
            dt.Columns.Add("Created", typeof(DateTime));
            dt.Columns.Add("Modified", typeof(DateTime));
            foreach (var item in dir.GetDirectories())
            {
                var row=dt.NewRow();
                row["Name"]=item.Name;
                row["Created"]=item.CreationTime;
                row["Modified"]=item.LastWriteTime;

                dt.Rows.Add(row);
            }
            foreach (var item in dir.GetFiles())
            {
                var row = dt.NewRow();
                row["Name"] = item.Name;
                row["Size"] = item.Length;
                row["Created"] = item.CreationTime;
                row["Modified"] = item.LastWriteTime;

                dt.Rows.Add(row);
            }
            return dt;
        }*/

        private static IEnumerable<string> TableToCsv(SubArray<string> table, bool escapeCsv = true,  bool writeHeader = false)
        {
            var st = TableTokenizer(table);

            if (writeHeader)
            {
                var h = HeaderTokenizer(st.Header).ToArray();

                if (escapeCsv)
                {
                    yield return string.Join(",",
                        EscapeCsv(h[0]),
                        EscapeCsv("Odds Ratio (Low High)"),
                        EscapeCsv(h[4]));
                }
                else
                {
                    yield return string.Join(",",
                        h[0],
                        "Odds Ratio (Low High)",
                        h[4]);
                }
            }
            
            foreach (var row in st.Rows)
            {
                var r = RowTokenizer(row).ToArray();
                if (escapeCsv)
                {
                    yield return string.Join(",",
                        EscapeCsv(r[0]),
                        EscapeCsv(string.Format(@"{0} ({1} {2})", r[1], r[5], r[6])),
                        EscapeCsv(r[4]));
                }
                else
                {
                    yield return string.Join(",",
                        r[0],
                        string.Format(@"{0} ({1} {2})", r[1], r[5], r[6]),
                        r[4]);
                }
            }
        }

        private static string EscapeCsv(string s)
        {
            return @"""" + s + @"""";
        }

        private static StringTable TableTokenizer(SubArray<string> table)
        {
            return new StringTable
            {
                Header = table[1],
                Rows = new SubArray<string>(_inputLines, table.Offset + 3, table.Count - 4)
            };
        }

        private static IEnumerable<string> HeaderTokenizer(string header)
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

        private static IEnumerable<string> RowTokenizer(string row)
        {
            var nameSplit = row.Split("|");

            yield return nameSplit[0].Trim();

            var dataSplit = nameSplit[1].Split(" ", StringSplitOptions.RemoveEmptyEntries);

            foreach (var column in dataSplit)
                yield return column;
        }

        private static void IsolateCommands(NumberedLine[] numberedLines)
        {
            var commandString = ". xi:clogit";
            var commandContinuationString = ">";
            
            var withIndices = numberedLines
                .Where(a => a.Line.StartsWith(commandString) || a.Line.StartsWith(commandContinuationString))
                .OrderBy(a => a.LineNumber)
                .ToArray();

            var list = new List<Tuple<int, int>>();
            
            var firstLine = -1;
            var lastLine = -1;
            bool expectingAnotherLine = false;
            foreach (var c in withIndices)
            {
                if (c.Line.StartsWith(commandString))
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

        private static void IsolateTables(NumberedLine[] numberedLines)
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
        
        private static SubArray<string> SliceInput(Tuple<int, int> tableIndexTuple)
        {
            var offset = tableIndexTuple.Item1;
            var count = 1 + tableIndexTuple.Item2 - tableIndexTuple.Item1;
            return new SubArray<string>(_inputLines, offset, count);
        }

        private static void PrintSubArray(SubArray<string> subArray)
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