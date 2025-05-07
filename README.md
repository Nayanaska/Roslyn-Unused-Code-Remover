# Roslyn Analyzer Tool

This project uses the [Roslyn](https://github.com/dotnet/roslyn) compiler platform to analyze C# code and extract all ununsed variables, classes and functions and also you can comment it out. The analysis results are also exported to Excel using [EPPlus](https://github.com/EPPlusSoftware/EPPlus) there is one sample code added for saving it in excel for unused classes.

## 🚀 Features

- Load and analyze a C# Solution or Project using `MSBuildWorkspace`
- find all unused variables, functions and classes.
- comment out all unused code.
- Export analyzed data to Excel (`.xlsx`) using EPPlus
- Supports both CLI-based and extensible Roslyn workflows

## 🛠️ Technologies Used

- [.NET SDK](https://dotnet.microsoft.com/download) (6/7/8)
- Roslyn (`Microsoft.CodeAnalysis`)
- EPPlus (Excel file generation)
- MSBuild Workspace support (`Microsoft.Build.Locator`)

## 📦 Dependencies

Make sure the following NuGet packages are installed:

```bash
dotnet add package Microsoft.Build.Locator
dotnet add package Microsoft.CodeAnalysis.Workspaces.MSBuild
dotnet add package Microsoft.CodeAnalysis.CSharp.Workspaces
dotnet add package Microsoft.CodeAnalysis.FindSymbols
dotnet add package EPPlus
```

## 🧑‍💻 Getting Started

### 1. Clone the Repo

```bash
git clone https://github.com/your-username/roslyn-analyzer.git
cd roslyn-analyzer
```

### 2. Build the Project

```bash
dotnet build
```

### 3. Run the Analyzer

Update the path to your `.sln` or `.csproj` in `Program.cs` and excel file path, then:

```bash
dotnet run
```

The output Excel file will be saved to the project directory.

## 📝 License

This project uses EPPlus under its [Polyform Noncommercial License](https://epplussoftware.com/en/LicenseOverview). You must set the license context:

```csharp
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
```


## 🙋‍♀️ Contributions

Feel free to fork, improve, and submit PRs. Open issues for bugs or feature requests!

---

Made with ❤️ using .NET and Roslyn
