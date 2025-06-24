using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using CrystalDecisions.CrystalReports.Engine;

/*
 * Scan through folder containing Crystal Reports files (.rpt).  Looking for specific text.
 * 
 */

namespace ConsoleAppScan_rpt
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Crystal Reports Search Tool");
            Console.WriteLine("===========================");

            Console.WriteLine("Input location of crystal reports:");
            string location = Console.ReadLine();

            string folderPath = location;
            bool caseSensitive =  false;

            string searchTerm = "loss", searchTerm1 = "loss";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Error: Folder '{folderPath}' does not exist.");
                return;
            }

            Console.WriteLine($"Searching for '{searchTerm}' in Crystal Reports files...");
            Console.WriteLine($"Folder: {folderPath}");
            Console.WriteLine($"Case Sensitive: {caseSensitive}");
            Console.WriteLine();

            SearchCrystalReports(folderPath, searchTerm, searchTerm1, caseSensitive);

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        static void SearchCrystalReports(string folderPath, string searchTerm, string searchTerm1, bool caseSensitive = false)
        {
            var reportFiles = Directory.GetFiles(folderPath, "*.rpt", SearchOption.AllDirectories);
            int matchCount = 0;
            int totalFiles = reportFiles.Length;

            Console.WriteLine($"Found {totalFiles} Crystal Reports files to search.");
            Console.WriteLine();

            foreach (string reportFile in reportFiles)
            {
                try
                {
                    //Console.Write($"Searching: {Path.GetFileName(reportFile)}... ");

                    var matches = SearchInReport(reportFile, searchTerm, searchTerm1, caseSensitive);

                    if (matches.Any())
                    {
                        matchCount++;
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"FOUND ({matches.Count} matches)");
                        Console.ResetColor();

                        Console.WriteLine($"  File: {reportFile}");
                        foreach (var match in matches)
                        {
                            Console.WriteLine($"    - {match}");
                        }
                        Console.WriteLine();
                    }
                    //else
                    //{
                    //    Console.WriteLine("No matches");
                    //}
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"ERROR: {ex.Message}");
                    Console.ResetColor();
                }
            }

            Console.WriteLine($"\nSearch complete. Found matches in {matchCount} out of {totalFiles} files.");
        }

        static List<string> SearchInReport(string reportPath, string searchTerm, string searchTerm1, bool caseSensitive)
        {
            var matches = new List<string>();
            ReportDocument report = null;

            try
            {
                report = new ReportDocument();
                report.Load(reportPath);

                // Search in report sections
                foreach (Section section in report.ReportDefinition.Sections)
                {
                    foreach (ReportObject obj in section.ReportObjects)
                    {
                        string objectText = GetObjectText(obj);
                        if (!string.IsNullOrEmpty(objectText))
                        {
                            bool found = caseSensitive ?
                                objectText.Contains(searchTerm) :
                                objectText.ToLower().Contains(searchTerm.ToLower());

                            if (found)
                            {
                                matches.Add($"Section '{section.Name}' - {obj.GetType().Name}: {TruncateText(objectText, 100)}");
                            }
                        }
                    }
                }

                // Search in subreports
                foreach (Section section in report.ReportDefinition.Sections)
                {
                    foreach (ReportObject obj in section.ReportObjects)
                    {
                        if (obj is SubreportObject subreport)
                        {
                            var subreportMatches = SearchInSubreport(report, subreport, searchTerm, caseSensitive);
                            matches.AddRange(subreportMatches);
                        }
                    }
                }

                // Search in formulas
                var formulaMatches = SearchInFormulas(report, searchTerm, caseSensitive);
                matches.AddRange(formulaMatches);

                // Search in parameter fields
                var parameterMatches = SearchInParameters(report, searchTerm, caseSensitive);
                matches.AddRange(parameterMatches);
            }
            finally
            {
                report?.Close();
                report?.Dispose();
            }

            return matches;
        }

        static string GetObjectText(ReportObject obj)
        {
            try
            {
                switch (obj)
                {
                    case TextObject textObj:
                        return textObj.Text;
                    case FieldObject fieldObj:
                        return fieldObj.DataSource?.ToString() ?? "";
                    //case ParameterFieldObject paramObj:
                    //    return paramObj.ParameterFieldName;
                    default:
                        return obj.Name ?? "";
                }
            }
            catch
            {
                return "";
            }
        }

        static List<string> SearchInSubreport(ReportDocument mainReport, SubreportObject subreportObj, string searchTerm, bool caseSensitive)
        {
            var matches = new List<string>();

            try
            {
                ReportDocument subreport = subreportObj.OpenSubreport(subreportObj.SubreportName);

                foreach (Section section in subreport.ReportDefinition.Sections)
                {
                    foreach (ReportObject obj in section.ReportObjects)
                    {
                        string objectText = GetObjectText(obj);
                        if (!string.IsNullOrEmpty(objectText))
                        {
                            bool found = caseSensitive ?
                                objectText.Contains(searchTerm) :
                                objectText.ToLower().Contains(searchTerm.ToLower());

                            if (found)
                            {
                                matches.Add($"Subreport '{subreportObj.SubreportName}' - Section '{section.Name}' - {obj.GetType().Name}: {TruncateText(objectText, 100)}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                matches.Add($"Error searching subreport '{subreportObj.SubreportName}': {ex.Message}");
            }

            return matches;
        }

        static List<string> SearchInFormulas(ReportDocument report, string searchTerm, bool caseSensitive)
        {
            var matches = new List<string>();

            try
            {
                foreach (FormulaFieldDefinition formula in report.DataDefinition.FormulaFields)
                {
                    string formulaText = formula.Text;
                    if (!string.IsNullOrEmpty(formulaText))
                    {
                        bool found = caseSensitive ?
                            formulaText.Contains(searchTerm) :
                            formulaText.ToLower().Contains(searchTerm.ToLower());

                        if (found)
                        {
                            matches.Add($"Formula '{formula.FormulaName}': {TruncateText(formulaText, 100)}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                matches.Add($"Error searching formulas: {ex.Message}");
            }

            return matches;
        }

        static List<string> SearchInParameters(ReportDocument report, string searchTerm, bool caseSensitive)
        {
            var matches = new List<string>();

            try
            {
                foreach (ParameterFieldDefinition param in report.DataDefinition.ParameterFields)
                {
                    // Search parameter name
                    bool nameFound = caseSensitive ?
                        param.ParameterFieldName.Contains(searchTerm) :
                        param.ParameterFieldName.ToLower().Contains(searchTerm.ToLower());

                    if (nameFound)
                    {
                        matches.Add($"Parameter Name: {param.ParameterFieldName}");
                    }

                    // Search parameter prompt text
                    if (!string.IsNullOrEmpty(param.PromptText))
                    {
                        bool promptFound = caseSensitive ?
                            param.PromptText.Contains(searchTerm) :
                            param.PromptText.ToLower().Contains(searchTerm.ToLower());

                        if (promptFound)
                        {
                            matches.Add($"Parameter Prompt '{param.ParameterFieldName}': {param.PromptText}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                matches.Add($"Error searching parameters: {ex.Message}");
            }

            return matches;
        }

        static string TruncateText(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
                return text;

            return text.Substring(0, maxLength) + "...";
        }
    }
}