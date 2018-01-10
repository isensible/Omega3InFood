using CommandLine;

namespace Omega3
{
    internal class Options
    {
        [Option('i', "input", Required = true, HelpText = "Full path to input Excel file")]
        public string Input { get; set; }

        [Option('o', "output", Required = true, HelpText = "Full path of output file")]
        public string Output { get; set; }
    }
}