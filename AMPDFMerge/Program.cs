using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Security;
using PdfSharp.Pdf.IO;
using System.IO;
using CommandLine;

namespace AMPDFMerge
{
    class Program
    {        
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();        
        private static string AppPath;
        private static string pwd;

        static void Main(string[] args)
        {
            AppPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            
            var result = CommandLine.Parser.Default.ParseArguments<Options>(args).MapResult( 
                opts => RunOptions(opts), 
                errs => HandleParseError(errs)
                );

            Console.WriteLine("\nReturn code= {0}", result);

            //CommandLine.Parser.Default.ParseArguments<Options>(args)
            //    .WithParsed(RunOptions)
            //    .WithNotParsed(HandleParseError);

        }

        static void LoggerConfigure(Options opts)
        {
            var config = new NLog.Config.LoggingConfiguration();

            // Targets where to log to: File and Console
            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = AppPath + "\\logs\\AMPDFMerge-"+ DateTime.Now.ToString("yyyyMMdd") + ".log" };
            logfile.MaxArchiveFiles = 30;
            logfile.ArchiveAboveSize = 1024000;

            var logconsole = new NLog.Targets.ConsoleTarget("logconsole");

            config.AddRule(LogLevel.Info, LogLevel.Fatal, logfile);

            if (opts.Verbose) 
                config.AddRule(LogLevel.Info, LogLevel.Fatal, logconsole);            

            // Apply config           
            NLog.LogManager.Configuration = config;
        }

        public static void MergePDFs(string outputFile, string outputPassword, string[] pdfs)
        {
            Logger.Info("Merge Normal PDF");
            using (PdfDocument targetDoc = new PdfDocument())
            {
                if (outputPassword != null)
                {
                    Logger.Info("targetDoc.SecuritySettings.UserPassword = " + outputPassword);
                    targetDoc.SecuritySettings.OwnerPassword = outputPassword;
                    targetDoc.SecuritySettings.UserPassword = outputPassword;
                }

                foreach (string pdf in pdfs)
                {
                    using (PdfDocument pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                    {
                        bool hasOwnerAccess = pdfDoc.SecuritySettings.HasOwnerPermissions;
                        for (int i = 0; i < pdfDoc.PageCount; i++)
                        {
                            Logger.Info("targetDoc.AddPage(" + i.ToString() + ")");
                            targetDoc.AddPage(pdfDoc.Pages[i]);
                        }
                    }
                }
                PdfDocumentSecurityLevel level = targetDoc.SecuritySettings.DocumentSecurityLevel;
                Logger.Info("targetDoc.Save(" + outputFile + ")");
                targetDoc.Save(outputFile);
            }
        }

        // The 'get the password' call back function.        
        static void PasswordProvider(PdfPasswordProviderArgs args)
        {            
            args.Password = pwd;           
        }


        public static void MergeEncryptedPDFs(string outputFile, string outputPassword, string[] pdfs, string inputPassword)
        {
            Logger.Info("Merge Encrypted PDF");

            pwd = inputPassword;
            using (PdfDocument targetDoc = new PdfDocument())
            {
                if (outputPassword != null)
                {
                    Logger.Info("targetDoc.SecuritySettings.UserPassword = " + outputPassword);
                    targetDoc.SecuritySettings.OwnerPassword = outputPassword;
                    targetDoc.SecuritySettings.UserPassword = outputPassword;
                }

                foreach (string pdf in pdfs)
                {
                  
                    PdfDocument pdfDoc;
                    // checking valid password, throw error if invalid
                    pdfDoc = PdfReader.Open(pdf, inputPassword);

                    pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import, PasswordProvider);

                    bool hasOwnerAccess = pdfDoc.SecuritySettings.HasOwnerPermissions;
                    for (int i = 0; i < pdfDoc.PageCount; i++)
                    {
                        Logger.Info("targetDoc.AddPage(" + i.ToString() + ")");
                        targetDoc.AddPage(pdfDoc.Pages[i]);
                    }                                    
                }

                PdfDocumentSecurityLevel level = targetDoc.SecuritySettings.DocumentSecurityLevel;
                Logger.Info("targetDoc.Save(" + outputFile + ")");
                targetDoc.Save(outputFile);
            }
        }

        static void TestMode(Options opts)
        {
            LoggerConfigure(opts);
            //handle options
            Logger.Info("| --------- |");
            Logger.Info("| Test Mode |");
            Logger.Info("| _________ |");
            Logger.Info("SourceDir = {0}", opts.InputDir);
            Logger.Info("OutputFile = {0}", opts.OutputFile);

            string[] pdfs = Directory.GetFiles(opts.InputDir, "*.pdf");

            
            MergeEncryptedPDFs(opts.OutputFile , opts.OutputPassword, pdfs, opts.InputPassword);
        }

        static int RunOptions(Options opts)
        {
            var exitCode = 0;
            string[] pdfs;
            LoggerConfigure(opts);
            Logger.Info("LoggerConfigure(opts) successed!");

            if (opts.InputFile.Count() > 0)
            {
                Logger.Info("opts.InputFile is not null");
                pdfs = opts.InputFile.Cast<string>().ToArray();

                if (opts.InputDir != "")
                {
                    string[] tmppdfs = new string[pdfs.Length];
                    pdfs.CopyTo(tmppdfs, 0);

                    int cnt = 0;
                    foreach (string pdf in tmppdfs)
                    {
                        pdfs[cnt] = opts.InputDir + @"\" + pdf;
                        cnt++;
                    }
                }


            }
            else
            {
                Logger.Info("Directory.GetFiles("+ opts.InputDir + ")");
                pdfs = Directory.GetFiles(opts.InputDir, "*.pdf");
            }

            if (opts.Testmode)
            {
                TestMode(opts);
            }
            else
            {
                if (opts.InputPassword != null)
                {
                    try
                    {
                        MergeEncryptedPDFs(opts.OutputFile, opts.OutputPassword, pdfs, opts.InputPassword);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex);
                        exitCode = -2;
                    }
                }
                else
                {
                    try
                    {
                        MergePDFs(opts.OutputFile, opts.OutputPassword, pdfs);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex);
                        exitCode = -2;
                    }

                }

            }

            return exitCode;

        }
        static int HandleParseError(IEnumerable<Error> errs)
        {
            var exitCode=-2;
            //handle errors
            if (errs.Any(x => x is HelpRequestedError || x is VersionRequestedError))
                exitCode = -1;

            Console.WriteLine("Parameter unknown, please check the documentation or use parameter '--help' for more information");

            return exitCode;

        }


    }
}
