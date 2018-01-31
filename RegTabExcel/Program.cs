using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using CommandLine;

namespace RegTabExcel
{
    public class Options
    {
        [Option('i', "input", Required = true, HelpText = "Full path to input file")]
        public string Input { get; set; }

        [Option('o', "output", Required = true, HelpText = "Full path of output Excel file")]
        public string Output { get; set; }
    }

    public class LineWithIndex
    {
        public string Line { get; set; }
        public int LineIndex { get; set; }
    }
    
    class Program
    {
        private static string[] _inputLines;
        private static Tuple<int, int>[] _tableIndexTuples;
        private static Tuple<int, int>[] _commandLineTuples;

        static int Main(string[] args)
        {
            var result = Parser.Default.ParseArguments<Options>(args);
            
            return result.MapResult(PutRegressionTablesIntoExcel, errors => 1);
        }

        private static int PutRegressionTablesIntoExcel(Options options)
        {
            // 1. Read the input file into array of `lines`
            // 2. Find the locations in `lines` of each regression table
            // 3. Print the tables to console, giving each table an unique ordinal
            // 4. Ask user which table to export, and how many rows [X..Y] to export

            _inputLines = File.ReadAllLines(options.Input);

            var lineWithIndices = _inputLines
                .Select((l, i) => new LineWithIndex {Line = l.Trim(), LineIndex = i})
                .Skip(1)
                .SkipLastN(1)
                .ToArray();

            SetTableIndexTuples(lineWithIndices);

            SetCommandIndexTuples(lineWithIndices);

            var merged = _commandLineTuples
                .Concat(_tableIndexTuples)
                .OrderBy(t => t.Item1);

            foreach (var t in merged)
            {
                PrintLines(t.Item1, t.Item2);
            }

            Console.ReadKey();

            return 0;
        }

        private static void SetCommandIndexTuples(LineWithIndex[] lineWithIndices)
        {
            var withIndices = lineWithIndices
                .Where(a => a.Line.StartsWith(". xi:clogit") || a.Line.StartsWith(">"))
                .OrderBy(a => a.LineIndex)
                .ToArray();

            var list = new List<Tuple<int, int>>();
            
            var firstLine = -1;
            var lastLine = -1;
            bool expectingAnotherLine = false;
            foreach (var c in withIndices)
            {
                if (c.Line.StartsWith(". xi:clogit"))
                {
                    firstLine = c.LineIndex;
                    lastLine = c.LineIndex;
                    expectingAnotherLine = c.Line.EndsWith("///");
                }
                
                if (c.Line.StartsWith(">") && expectingAnotherLine)
                {
                    lastLine = c.LineIndex;
                    expectingAnotherLine = c.Line.EndsWith("///");
                }

                if (!expectingAnotherLine && firstLine >= 0)
                {
                    list.Add(new Tuple<int, int>(firstLine, lastLine));
                    firstLine = -1;
                    lastLine = -1;
                }
            }

            _commandLineTuples = list.ToArray();
        }

        private static void SetTableIndexTuples(LineWithIndex[] lineWithIndices)
        {
            var tableLines = lineWithIndices
                .Where(a => { return a.Line.Length != 0 && a.Line.All(c => c.CompareTo('-') == 0); })
                .ToArray();

            var tableBorderLineIndices = tableLines
                .Select(a => a.LineIndex)
                .ToList();

            _tableIndexTuples = SplitList(tableBorderLineIndices)
                .Where(t => t.Count == 2)
                .Select(t => new Tuple<int, int>(t[0], t[1]))
                .ToArray();
        }

        private static void PrintLines(int first, int last)
        {
            for (int i = first; i <= last; i++)
            {
                Console.WriteLine(_inputLines[i]);
            }
        }
        
        private static IEnumerable<List<T>> SplitList<T>(List<T> input, int size=2)  
        {        
            for (var i = 0; i < input.Count; i+= size) 
            { 
                yield return input.GetRange(i, Math.Min(size, input.Count - i)); 
            }  
        }
    }

    public static class EnumerableExtensions
    {
        public static IEnumerable<T> SkipLastN<T>(this IEnumerable<T> source, int n) {
            // ReSharper disable once GenericEnumeratorNotDisposed
            var  it = source.GetEnumerator();
            bool hasRemainingItems = false;
            var  cache = new Queue<T>(n + 1);

            do {
                // ReSharper disable once AssignmentInConditionalExpression
                if (hasRemainingItems = it.MoveNext()) {
                    cache.Enqueue(it.Current);
                    if (cache.Count > n)
                        yield return cache.Dequeue();
                }
            } while (hasRemainingItems);
        }

    }
}