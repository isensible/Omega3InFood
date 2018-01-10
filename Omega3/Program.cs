using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CommandLine;
using OfficeOpenXml;

namespace Omega3
{
    static class Program
    {
        private static readonly string StataLineBreak = " ///" + Environment.NewLine;
        
        private static readonly Omega3[] Omegas = {
            new Omega3 { Name = "n3mg", Column = "C" },
            new Omega3 { Name = "ALAg", Column = "D" },
            new Omega3 { Name = "EPAmg", Column = "E" },
            new Omega3 { Name = "DPAmg", Column = "F" },
            new Omega3 { Name = "DHAmg", Column = "G" }
        };

        private static readonly string WorksheetName = "All foods";


        static int Main(string[] args)
        {
            var result = Parser.Default.ParseArguments<Options>(args);
            
            return result.MapResult(GenerateStataFromExcel, errors => 1);
        }

        private static string GetCellRange(string column)
        {
            return column + "3:" + column + "59";
        }

        private static int GenerateStataFromExcel(Options options)
        {
            var file = new FileInfo(options.Input);

            if (!file.Exists)
            {
                Console.WriteLine("ERROR: input file not found");
                Console.ReadKey();
                return 1;
            }
            
            var outputText = new StringBuilder();

            using (var excelPackage = new ExcelPackage(file))
            {
                var allFoodsWorksheet = excelPackage.Workbook.Worksheets
                    .FirstOrDefault(w => w.Name == WorksheetName);

                if (allFoodsWorksheet == null)
                {
                    Console.WriteLine("ERROR: all foods worksheet not found");
                    Console.ReadKey();
                    return 1;
                }

                var foods = allFoodsWorksheet
                    .Cells[GetCellRange(Food.Column)]
                    .Select((f,i) => new { Name = f.Value, Row = i})
                    .ToArray();
                
                foreach (var omega3 in Omegas)
                {
                    outputText.AppendFormat("** {0}{1}", omega3.Name, Environment.NewLine);
                    outputText.AppendLine("** for each food:");
                    
                    var omega3Cells = allFoodsWorksheet
                        .Cells[GetCellRange(omega3.Column)]
                        .Select((o, i) => new { Omega3Value = o.Value, Row = i })
                        .ToArray();

                    var listOfFoodVarsForSumming = new List<string>();
                    
                    foreach (var food in foods)
                    {
                        var v = omega3Cells.FirstOrDefault(x => x.Row == food.Row);

                        listOfFoodVarsForSumming.Add(string.Format("{0}_{1}", food.Name, omega3.Name));
                        
                        var s = string.Format("gen {0}_{1}={0} * ({2}/100)", food.Name, omega3.Name, v?.Omega3Value);

                        outputText.AppendLine(s);
                        Console.WriteLine(s);
                    }

                    outputText.AppendFormat("** SUM of foods for {0}{1}", omega3.Name, Environment.NewLine);

                    var genSum = string.Format("gen {0}_day={1} ", omega3.Name, StataLineBreak);
                    outputText.Append(genSum);
                    Console.Write(genSum);

                    var splitList = SplitList(listOfFoodVarsForSumming).ToArray();
                    foreach (var split in splitList.Select((v,i) => new {Vars=v, Index=i}))
                    {
                        var nextLineOfSum = string.Format("{0}{1}{2}",
                            split.Index == 0 ? "  " : " + ",
                            string.Join(" + " , split.Vars),
                            split.Index == splitList.Length-1 ? Environment.NewLine : StataLineBreak);
                        
                        outputText.Append(nextLineOfSum);
                        Console.Write(nextLineOfSum);   
                    }
                    
                    outputText.AppendLine();
                    Console.WriteLine();
                }
            }

            File.WriteAllText(options.Output, outputText.ToString());
            
            Console.WriteLine("Done");

            return 0;
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