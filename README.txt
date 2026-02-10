ProductionCalc Prototype (Login + Dashboard + Excel Reports)

Projects:
- ProductionCalc.Core (calculations)
- ProductionCalc.Desktop (WPF app)

How to run:
1. Open ProductionCalc.sln in Visual Studio 2022
2. Restore NuGet packages (LiveChartsCore.SkiaSharpView.WPF 2.0.0-rc3, ClosedXML 0.95.4)
   - If LiveCharts fails, enable 'Include prerelease' in NuGet Package Manager.
3. Build solution
4. Run the app (Login window appears)
   - Use admin/password123 or engineer/engine2025
5. Use tabs to calculate and export reports to Excel (Reports folder)

Notes:
- LiveCharts is prerelease; if package fails to restore try enabling prerelease in NuGet Package Manager.
- ClosedXML is used for Excel export.
