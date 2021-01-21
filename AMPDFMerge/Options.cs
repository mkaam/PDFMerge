using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using CommandLine.Text;

namespace AMPDFMerge
{
    class Options
    {
        [Option('d',"inputdir", Required = true , HelpText = "input file directory, ex: E:\\PDFDir")]
        public string InputDir { get; set; }

        [Option('i', "inputfile", Required = false, HelpText = "pdf file to be merged, ex: file.pdf file2.pdf file3.pdf ...")]
        public IEnumerable<string> InputFile { get; set; }

        [Option('p', "inputpassword", Required = false, HelpText = "input pdf password dictionary (must be set if input pdf file encrypted). ex : Passwowd1 Password2")]
        public IEnumerable<string> InputPassword { get; set; }

        [Option('o', "outputfile", Required = true, HelpText = "pdf file to be merged with (full path), ex: E:\\PDFMerge\\output.pdf")]
        public string OutputFile { get; set; }

        [Option('q', "outputpassword", Required = false, HelpText = "output pdf password (optional)")]
        public string OutputPassword { get; set; }

        [Option(HelpText = "TestMode")]
        public bool Testmode { get; set; }

        [Option(HelpText = "Print process output to console")]
        public bool Verbose { get; set; }

    }
}
