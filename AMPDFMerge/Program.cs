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
            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = $"{AppPath}\\logs\\{(opts.LogName == null ? "PDFMerge": opts.LogName)}-{DateTime.Now.ToString("yyyyMMdd")}.log" };
            if (opts.LogFile != "" && opts.LogFile != null)
            {
                logfile = new NLog.Targets.FileTarget("logfile") { FileName = opts.LogFile };
            }
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
            Logger.Info("Merge Normal PDF Start...");
            using (PdfDocument targetDoc = new PdfDocument())
            {
                if (outputPassword != null)
                {
                    Logger.Info("Set OwnerPassword : {0}", outputPassword);
                    targetDoc.SecuritySettings.OwnerPassword = outputPassword;
                    Logger.Info("Set UserPassword : {0}", outputPassword);
                    targetDoc.SecuritySettings.UserPassword = outputPassword;
                }

                foreach (string pdf in pdfs)
                {
                    try
                    {
                        using (PdfDocument pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                        {
                            bool hasOwnerAccess = pdfDoc.SecuritySettings.HasOwnerPermissions;
                            for (int i = 0; i < pdfDoc.PageCount; i++)
                            {
                                Logger.Info("Add PDF : {0}", pdf);
                                targetDoc.AddPage(pdfDoc.Pages[i]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, "MergePDFs");
                    }

                }
                PdfDocumentSecurityLevel level = targetDoc.SecuritySettings.DocumentSecurityLevel;                
                targetDoc.Save(outputFile);
                Logger.Info("Merge Normal PDF success!, Save to : {0}", outputFile);
            }
        }

        // The 'get the password' call back function.        
        static void PasswordProvider(PdfPasswordProviderArgs args)
        {            
            args.Password = pwd;           
        }


        public static void MergeEncryptedPDFs(string outputFile, string outputPassword, string[] pdfs, IEnumerable<string> inputPassword)
        {
            Logger.Info("Merge Encrypted PDF Starting...");
            var passpdfs = inputPassword.Cast<string>().ToArray();
            using (PdfDocument targetDoc = new PdfDocument())
            {
                if (outputPassword != null)
                {
                    Logger.Info("Set OwnerPassword = " + outputPassword);
                    targetDoc.SecuritySettings.OwnerPassword = outputPassword;
                    Logger.Info("Set UserPassword = " + outputPassword);
                    targetDoc.SecuritySettings.UserPassword = outputPassword;
                }

                foreach (string pdf in pdfs)
                {
                    Logger.Info("Adding PDF : {0} ...", pdf);
                    PdfDocument pdfDoc;
                    foreach (string mypass in inputPassword)
                    {

                        try
                        {
                            Logger.Info("Trying unlock PDF {0} using password : {1}", pdf, mypass);
                            pdfDoc = PdfReader.Open(pdf, mypass);
                            pwd = mypass;
                            Logger.Info("PDF {0} unlocked with password : {1}", pdf, mypass);

                            pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import, PasswordProvider);

                            bool hasOwnerAccess = pdfDoc.SecuritySettings.HasOwnerPermissions;
                            for (int i = 0; i < pdfDoc.PageCount; i++)
                            {                                
                                targetDoc.AddPage(pdfDoc.Pages[i]);
                            }
                            break;
                        }
                        catch (PdfReaderException pdfex) {                            
                            if (!pdfex.Message.Contains("The specified password is invalid"))
                                Logger.Error(pdfex);                                                            
                            else
                                Logger.Error("Invalid password : {0}", mypass);
                        }
                    }
                    // checking valid password, throw error if invalid  
                    Logger.Info("PDF : {0} successfully added", pdf);

                }
               
                PdfDocumentSecurityLevel level = targetDoc.SecuritySettings.DocumentSecurityLevel;
                targetDoc.Save(outputFile);
                Logger.Info("Merge Encrypted PDF success!, Save to : {0}", outputFile);
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

            
            //MergeEncryptedPDFs(opts.OutputFile , opts.OutputPassword, pdfs, "aaa");
        }

        static int RunOptions(Options opts)
        {
            var exitCode = 0;
            string[] pdfs;
            LoggerConfigure(opts);
            Logger.Debug("LoggerConfigure(opts) successed!");

            if (opts.InputFile.Count() > 0)
            {
                Logger.Debug("opts.InputFile is not null");
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
                Logger.Info("GetFiles PDF file from Directory : {0}", opts.InputDir);
                pdfs = Directory.GetFiles(opts.InputDir, "*.pdf");
            }

            if (opts.Testmode)
            {
                TestMode(opts);
            }
            else
            {
                if (opts.InputPassword.Count() > 0)
                {
                    //var passpdfs = opts.InputPassword.Cast<string>().ToArray();
                    MergeEncryptedPDFs(opts.OutputFile, opts.OutputPassword, pdfs, opts.InputPassword);
                    

                    //foreach (var passpdf in passpdfs)
                    //{
                    //    try
                    //    {
                    //        MergeEncryptedPDFs(opts.OutputFile, opts.OutputPassword, pdfs, passpdf);
                    //    }
                    //    catch (PdfReaderException pdfex)
                    //    {
                    //        Logger.Error("Invalid password : {0}", passpdf);
                    //        if (!pdfex.Message.Contains("The specified password is invalid")) {
                    //            Logger.Error(pdfex);
                    //            exitCode = -2;
                    //        }
                    //    }
                    //    catch(Exception ex)
                    //    {
                    //        Logger.Error(ex);
                    //        exitCode = -2;
                    //    }
                        
                    //}
                    
                    //try
                    //{
                    //    MergeEncryptedPDFs(opts.OutputFile, opts.OutputPassword, pdfs, opts.InputPassword);
                    //}
                    //catch (Exception ex)
                    //{
                    //    Logger.Error(ex);
                    //    exitCode = -2;
                    //}
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
