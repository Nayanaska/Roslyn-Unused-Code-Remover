using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Build.Locator;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.MSBuild;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;
using OfficeOpenXml;
using Microsoft.CodeAnalysis.FindSymbols;
using OfficeOpenXml.Style; // EPPlus for Excel
class Program
{
    static readonly List<string> IgnoredFolders = new List<string>
    {
        "Tests",
        "SampleSite",
        "PublicDomain"
    };

    static readonly string FailedChangesExcelPath = @"C:\Repo\FailedChanges.xlsx";
    static readonly string LogFilePath = @"C:\Repo\UnusedVariableLog.txt";

    static async Task Main(string[] args)
    {
        MSBuildLocator.RegisterDefaults();
        string solutionPath = @"Path/to/your/solution";

        // Start logging
        using StreamWriter logWriter = new(LogFilePath, append: false);
        Console.SetOut(logWriter);
        Console.SetError(logWriter);

        Console.WriteLine($"[INFO] Starting Unused Variable Detection: {DateTime.Now}");

        try
        {
            using var workspace = MSBuildWorkspace.Create();
            var solution = await workspace.OpenSolutionAsync(solutionPath);
            Console.WriteLine("find unused avariables");
            var unusedVariables = await FindUnusedVariables(solution);
            Console.WriteLine("found unused avariables, now start commenting");

            await CommentOutUnusedVariables(unusedVariables);

              var unusedFunctions = await FindUnusedFunctions(solutionPath);
            Console.WriteLine("found unused functions, now start commenting");

            await CommentOutUnusedFunctions(unusedFunctions);

            var unusedClasses = await FindUnusedClasses(solutionPath);
            Console.WriteLine("found unused classes, now start commenting");

            await CommentOutUnusedClasses(unusedClasses);
            Console.WriteLine("commenting done");

        }
        catch (Exception ex)
        {
            Console.WriteLine($"[ERROR] Unexpected error: {ex.Message}\n{ex.StackTrace}");
        }

        Console.WriteLine($"[INFO] Finished Processing: {DateTime.Now}");
    }

    static async Task<List<(string name, string filePath, int lineNumber, string fullLine)>> FindUnusedVariables(Solution solution)
    {
        var unusedVariables = new List<(string, string, int, string)>();

        foreach (var project in solution.Projects)
        {
            foreach (var document in project.Documents)
            {
                string filePath = document.FilePath ?? "";
                if (IgnoredFolders.Any(folder => filePath.Contains(folder, StringComparison.OrdinalIgnoreCase)))
                    continue;

                try
                {
                    var semanticModel = await document.GetSemanticModelAsync();
                    var syntaxTree = await document.GetSyntaxTreeAsync();
                    if (semanticModel == null || syntaxTree == null) continue;

                    var root = await syntaxTree.GetRootAsync();
                    var text = await document.GetTextAsync();

                    var variables = root.DescendantNodes()
                        .OfType<VariableDeclaratorSyntax>()
                        .Select(v => new { Symbol = semanticModel.GetDeclaredSymbol(v), SyntaxNode = v })
                        .Where(v => v.Symbol != null)
                        .ToList();

                    foreach (var variable in variables)
                    {
                        var references = await SymbolFinder.FindReferencesAsync(variable.Symbol, solution);
                        if (!references.Any(refs => refs.Locations.Any()))
                        {
                            var lineNumber = text.Lines.GetLineFromPosition(variable.SyntaxNode.SpanStart).LineNumber + 1;
                            var fullLine = text.Lines[lineNumber - 1].ToString(); // Get full line

                            unusedVariables.Add((variable.Symbol.ToDisplayString(), filePath, lineNumber, fullLine));
                            Console.WriteLine($"[INFO] Found Unused Variable: {variable.Symbol.ToDisplayString()} at {filePath} (Line {lineNumber})");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ERROR] Failed processing file: {filePath}\n{ex.Message}\n{ex.StackTrace}");
                }
            }
        }
        return unusedVariables;
    }

     static async Task<List<(string name, string filePath, int lineNumber)>> FindUnusedFunctions(string solutionPath)
    {
        using var workspace = MSBuildWorkspace.Create();
        var solution = await workspace.OpenSolutionAsync(solutionPath);
        var unusedList = new List<(string, string, int)>();

        foreach (var project in solution.Projects)
        {
            var compilation = await project.GetCompilationAsync();

            foreach (var syntaxTree in compilation.SyntaxTrees)
            {
                var semanticModel = compilation.GetSemanticModel(syntaxTree);
                var root = await syntaxTree.GetRootAsync();

                var methods = root.DescendantNodes().OfType<MethodDeclarationSyntax>();

                foreach (var method in methods)
                {
                    var symbol = semanticModel.GetDeclaredSymbol(method);
                    if (symbol == null) continue;

                    var refs = await SymbolFinder.FindReferencesAsync(symbol, solution);
                    if (!refs.Any(r => r.Locations.Any(l => !l.IsCandidateLocation)))
                    {
                        var location = method.GetLocation().GetLineSpan();
                        var filePath = syntaxTree.FilePath;
                        var line = location.StartLinePosition.Line + 1;

                        unusedList.Add((symbol.ToString(), filePath, line));
                        Console.WriteLine($"[UNUSED] {symbol} in {filePath} at line {line}");
                    }
                }
            }
        }

        return unusedList;
    }



    static async Task CommentOutUnusedVariables(
        List<(string name, string filePath, int lineNumber, string fullLine)> unusedVariables)
    {
        var unusedItems = unusedVariables.GroupBy(item => item.filePath)
                                         .ToDictionary(g => g.Key, g => g.ToList());

        List<(string name, string filePath, int lineNumber)> failedChanges = new();

        foreach (var file in unusedItems)
        {
            string filePath = file.Key;
            List<(string name, string filePath, int lineNumber, string fullLine)> items = file.Value;

            try
            {
                string[] lines = await File.ReadAllLinesAsync(filePath);
                string[] originalLines = (string[])lines.Clone();

                foreach (var item in items.OrderByDescending(i => i.lineNumber)) // Process from bottom up
                {
                    if (item.lineNumber - 1 < 0 || item.lineNumber - 1 >= lines.Length)
                        continue;

                    lines[item.lineNumber - 1] = "// Remove unused code" + lines[item.lineNumber - 1];

                    await File.WriteAllLinesAsync(filePath, lines);
                     Console.WriteLine($"[INFO] Successfully commented unused variable in {filePath} (Line {item.lineNumber})");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] Could not process file {filePath}\n{ex.Message}\n{ex.StackTrace}");
            }
        }

        SaveFailedChangesToExcel(failedChanges);
    }

    static async Task CommentOutUnusedFunctions(
        List<(string name, string filePath, int lineNumber)> unusedItems)
    {
        List<(string name, string filePath, int lineNumber)> failedChanges = new();
        var grouped = unusedItems.GroupBy(i => i.filePath);

        foreach (var fileGroup in grouped)
        {
            string filePath = fileGroup.Key;
            string[] originalLines = await File.ReadAllLinesAsync(filePath);
            var updatedLines = originalLines.ToList();
            bool fileChanged = false;

            try
            {
                var text = await File.ReadAllTextAsync(filePath);
                var syntaxTree = CSharpSyntaxTree.ParseText(text);
                var root = await syntaxTree.GetRootAsync();

                foreach (var item in fileGroup.OrderByDescending(i => i.lineNumber))
                {
                    var textLines = syntaxTree.GetText().Lines;
                    if (item.lineNumber - 1 >= textLines.Count)
                        continue;

                    var methodNode = root.DescendantNodes()
                        .OfType<MethodDeclarationSyntax>()
                        .FirstOrDefault(n =>
                            syntaxTree.GetLineSpan(n.Span).StartLinePosition.Line == item.lineNumber - 1);

                    if (methodNode == null)
                        continue;

                    var startLine = syntaxTree.GetLineSpan(methodNode.Span).StartLinePosition.Line;
                    var endLine = syntaxTree.GetLineSpan(methodNode.Span).EndLinePosition.Line;

                    for (int i = startLine; i <= endLine; i++)
                    {
                        if (!updatedLines[i].TrimStart().StartsWith("//"))
                        {
                            updatedLines[i] = "// " + updatedLines[i];
                            fileChanged = true;
                        }
                    }
                    if (startLine > 0)
                    {
                        updatedLines.Insert(startLine, "//  Remove unused code");
                    }
                    else
                    {
                        updatedLines.Insert(0, "// Remove unused code");
                    }

                    Console.WriteLine($"[INFO] Commented unused method at {filePath} (Lines {startLine + 1}-{endLine + 1})");
                }

                if (fileChanged)
                {
                    await File.WriteAllLinesAsync(filePath, updatedLines);

                    // if (!BuildSolution(solutionPath))
                    // {
                    //     await File.WriteAllLinesAsync(filePath, originalLines);
                    //     failedChanges.AddRange(fileGroup);
                    //     Console.WriteLine($"[ERROR] Reverted {filePath} — build failed.");
                    // }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] Failed to comment in {filePath}: {ex.Message}");
                failedChanges.AddRange(fileGroup);
            }
        }
    }

    static bool BuildSolution(string solutionPath)
    {
        try
        {
            ProcessStartInfo psi = new("dotnet", $"build \"{solutionPath}\"")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using Process process = Process.Start(psi);
            process.WaitForExit();

            return process.ExitCode == 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[ERROR] Build execution failed: {ex.Message}");
            return false;
        }
    }

    static void SaveFailedChangesToExcel(List<(string name, string filePath, int lineNumber)> failedChanges)
    {
        if (failedChanges.Count == 0)
            return;

        try
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Failed Variables");

            worksheet.Cells[1, 1].Value = "Unused Variable";
            worksheet.Cells[1, 2].Value = "File Path";
            worksheet.Cells[1, 3].Value = "Line Number";

            for (int i = 0; i < failedChanges.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = failedChanges[i].name;
                worksheet.Cells[i + 2, 2].Value = failedChanges[i].filePath;
                worksheet.Cells[i + 2, 3].Value = failedChanges[i].lineNumber;
            }

            File.WriteAllBytes(FailedChangesExcelPath, package.GetAsByteArray());
            Console.WriteLine($"[INFO] Saved failed variables to Excel: {FailedChangesExcelPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[ERROR] Failed to write Excel file: {ex.Message}");
        }
    }


    static async Task<List<(string name, string filePath, int lineNumber, string fullLine)>> FindUnusedClasses(
    string solutionPath)
    {
        var unusedClasses = new List<(string, string, int, string)>();

        try
        {
            Console.WriteLine($"[INFO] Loading solution: {solutionPath}");
            var workspace = MSBuildWorkspace.Create();
            var solution = await workspace.OpenSolutionAsync(solutionPath);

            foreach (var project in solution.Projects)
            {
                Console.WriteLine($"[INFO] Analyzing project: {project.Name}");
                Compilation? compilation = null;

                try
                {
                    compilation = await project.GetCompilationAsync();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ERROR] Failed to compile project {project.Name}: {ex.Message}");
                    continue;
                }

                if (compilation == null) continue;

                foreach (var syntaxTree in compilation.SyntaxTrees)
                {
                    string filePath = syntaxTree.FilePath;

                    if (IgnoredFolders.Any(folder => filePath.Replace('\\', '/').Contains($"/{folder}/")))
                    {
                        Console.WriteLine($"[INFO] Skipped file in ignored folder: {filePath}");
                        continue;
                    }

                    try
                    {
                        var semanticModel = compilation.GetSemanticModel(syntaxTree);
                        var root = await syntaxTree.GetRootAsync();

                        var classDeclarations = root.DescendantNodes().OfType<ClassDeclarationSyntax>();

                        foreach (var classDecl in classDeclarations)
                        {
                            var classSymbol = semanticModel.GetDeclaredSymbol(classDecl);
                            if (classSymbol == null || classSymbol.DeclaredAccessibility != Accessibility.Public)
                                continue;

                            bool isUsed = false;

                            // Check reference to class itself
                            var classRefs = await SymbolFinder.FindReferencesAsync(classSymbol, solution);
                            if (classRefs.Any(r => r.Locations.Any(l => !l.IsCandidateLocation)))
                            {
                                isUsed = true;
                            }

                            // Check references to any members (methods, props, etc.)
                            if (!isUsed)
                            {
                                foreach (var member in classSymbol.GetMembers().Where(m => !m.IsImplicitlyDeclared))
                                {
                                    var refs = await SymbolFinder.FindReferencesAsync(member, solution);
                                    if (refs.Any(r => r.Locations.Any(l => !l.IsCandidateLocation)))
                                    {
                                        isUsed = true;
                                        break;
                                    }
                                }
                            }

                            if (!isUsed)
                            {
                                var lineSpan = syntaxTree.GetLineSpan(classDecl.Span);
                                int lineNumber = lineSpan.StartLinePosition.Line + 1;
                                string fullLine = syntaxTree.GetText().Lines[lineNumber - 1].ToString();

                                unusedClasses.Add((classSymbol.Name, filePath, lineNumber, fullLine));
                                Console.WriteLine($"[INFO] Found unused class: {classSymbol.Name} at {filePath} (Line {lineNumber})");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[ERROR] Error analyzing file {filePath}: {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[ERROR] Failed to process solution: {ex.Message}");
        }

        Console.WriteLine($"[INFO] Total unused classes found: {unusedClasses.Count}");
        SaveUnusedClassesToExcel(unusedClasses,@"C:\Repo\UnusedClasses.xlsx");
        return unusedClasses;
    }



    static async Task CommentOutUnusedClasses(
    List<(string name, string filePath, int lineNumber, string fullLine)> unusedClasses)
    {
        var grouped = unusedClasses.GroupBy(c => c.filePath);
        List<(string name, string filePath, int lineNumber)> failedChanges = new();

        foreach (var fileGroup in grouped)
        {
            string filePath = fileGroup.Key;

            try
            {
                Console.WriteLine($"[INFO] Processing file: {filePath}");
                var text = await File.ReadAllTextAsync(filePath);
                var syntaxTree = CSharpSyntaxTree.ParseText(text);
                var root = await syntaxTree.GetRootAsync();
                var updatedLines = text.Split('\n').ToList();

                foreach (var item in fileGroup.OrderByDescending(i => i.lineNumber))
                {
                    try
                    {
                        var classNode = root.DescendantNodes()
                            .OfType<ClassDeclarationSyntax>()
                            .FirstOrDefault(n =>
                                syntaxTree.GetLineSpan(n.Span).StartLinePosition.Line == item.lineNumber - 1 &&
                                n.Identifier.Text == item.name);

                        if (classNode == null)
                        { 
                            Console.WriteLine($"[WARN] Could not find class node for {item.name} in {filePath}");
                            continue;
                        }

                        var span = syntaxTree.GetLineSpan(classNode.Span);
                        int startLine = span.StartLinePosition.Line;
                        int endLine = span.EndLinePosition.Line;

                        // Insert comment above class
                        updatedLines.Insert(startLine, "// Remove unused code");
                        endLine++; // Adjust because of insert

                        for (int i = startLine + 1; i <= endLine; i++)
                        {
                            if (!updatedLines[i].TrimStart().StartsWith("//"))
                            {
                                updatedLines[i] = "// " + updatedLines[i];
                            }
                        }

                        Console.WriteLine($"[INFO] Commented class {item.name} in {filePath} (Lines {startLine + 1}-{endLine + 1})");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[ERROR] Failed to comment class {item.name} in {filePath}: {ex.Message}");
                        failedChanges.Add((item.name, item.filePath, item.lineNumber));
                    }
                }

                await File.WriteAllLinesAsync(filePath, updatedLines);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] Could not process file {filePath}: {ex.Message}");
                failedChanges.AddRange(fileGroup.Select(i => (i.name, i.filePath, i.lineNumber)));
            }
        }

        // SaveFailedChangesToExcel(failedChanges);
        Console.WriteLine("[INFO] Finished commenting unused classes.");
    }

    static void SaveUnusedClassesToExcel(List<(string name, string filePath, int lineNumber, string fullLine)> unusedClasses, string outputPath)
    {
        try
        {
            // ✅ Required in EPPlus 5+ (especially 8+)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Unused Classes");

            // Header
            worksheet.Cells[1, 1].Value = "Class Name";
            worksheet.Cells[1, 2].Value = "File Path";
            worksheet.Cells[1, 3].Value = "Line Number";
            worksheet.Cells[1, 4].Value = "Code Line";

            // Style
            using (var range = worksheet.Cells[1, 1, 1, 4])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            // Data
            for (int i = 0; i < unusedClasses.Count; i++)
            {
                var (name, filePath, lineNumber, fullLine) = unusedClasses[i];
                worksheet.Cells[i + 2, 1].Value = name;
                worksheet.Cells[i + 2, 2].Value = filePath;
                worksheet.Cells[i + 2, 3].Value = lineNumber;
                worksheet.Cells[i + 2, 4].Value = fullLine;
            }

            worksheet.Cells.AutoFitColumns();

            package.SaveAs(new FileInfo(outputPath));
            Console.WriteLine($"[INFO] Saved unused class list to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[ERROR] Failed to export to Excel: {ex.Message}");
        }
    }


}