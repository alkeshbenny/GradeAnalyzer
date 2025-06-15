# 📊 GradeAnalyzer

**GradeAnalyzer** is a simple C# console application that reads an Excel file of student grades, calculates the average for each student, determines pass/fail status, and exports the results to a new Excel file using the [ClosedXML](https://github.com/ClosedXML/ClosedXML) library.

---

## 🚀 Features

- ✅ Reads grades from an input Excel file (`Grades.xlsx`)
- 📈 Calculates average marks per student
- 🏆 Assigns "Pass" or "Fail" based on average
- 📤 Writes results to a new Excel file (`Results_ClosedXML_Only.xlsx`)
- 📍 Keeps output file in the same directory as the input

---

---

## 🛠️ Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)
- [ClosedXML NuGet package](https://www.nuget.org/packages/ClosedXML/)
  > Install with:
  ```bash
  dotnet add package ClosedXML


