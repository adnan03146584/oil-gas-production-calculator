using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using ProductionCalc.Core;
using ProductionCalc.Core.Auth;    // session & auth
using ProductionCalc.Core.Models;
using ProductionCalc.Core.Services;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media.Animation;
using System.Text.Json;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Drawing;

namespace ProductionCalc.Desktop
{
    public partial class MainWindow : Window
    {
        private readonly ReportService _reportService = new ReportService();
        private List<Report> _myReports = new List<Report>();

        private GasOrificeCalculator.GasOrificeResult? _lastGasResult;

        private class LiquidResult
        {
            public double GrossVol { get; set; }
            public double BLPD { get; set; }
            public double BOPD { get; set; }
            public double BWPD { get; set; }
            public double MeterDifference { get; set; }
        }
        private LiquidResult? _lastLiquidResult;

        private string _selectedChoke = "";
        private string _whp = "";
        private string _wht = "";
        private string _flp = "";
        private string _casing = "";
        private string _arrival = "";
        private string _bsw = "";
        private string? _lastExportedPath;


        private ObservableCollection<ReportEntry> _savedEntries = new ObservableCollection<ReportEntry>();

        private class ReportEntry
        {
            public string Timestamp { get; set; } = "";
            public string ChokeSize { get; set; } = "";
            public string ChokeWHP { get; set; } = "";
            public string ChokeWHT { get; set; } = "";
            public string ChokeFLP { get; set; } = "";
            public string ChokeCasing { get; set; } = "";
            public string ChokeArrival { get; set; } = "";
            public string ChokeBSW { get; set; } = "";
            public string ReservoirPressure { get; set; } = "";
            public string OilAPI { get; set; } = "";
            public string GasDP { get; set; } = "";
            public string GasDownstream { get; set; } = "";
            public string GasTemp { get; set; } = "";
            public string GasSG { get; set; } = "";
            public string GasMeterD { get; set; } = "";
            public string GasOrificeD { get; set; } = "";
            public string GasCO2 { get; set; } = "";
            public string GasN2 { get; set; } = "";
            public string GasH2S { get; set; } = "";
            public string GasUpstream { get; set; } = "";
            public string GasMW { get; set; } = "";
            public string GasZBase { get; set; } = "";
            public string GasZFlowing { get; set; } = "";
            public string GasFlowVol { get; set; } = "";
            public string GasQ { get; set; } = "";
            public string LiqMeterStart { get; set; } = "";
            public string LiqMeterEnd { get; set; } = "";
            public string LiqBSW { get; set; } = "";
            public string LiqShrinkage { get; set; } = "";
            public string LiqTimeFactor { get; set; } = "";
            public string LiqGrossVol { get; set; } = "";
            public string LiqBLPD { get; set; } = "";
            public string LiqBOPD { get; set; } = "";
            public string LiqBWPD { get; set; } = "";
            public string LiqMeterDifference { get; set; } = "";
        }

        public MainWindow()
        {
            InitializeComponent();

            SavedEntriesGrid.ItemsSource = _savedEntries;

            if (!SessionManager.IsLoggedIn || SessionManager.CurrentUser == null)
            {
                MessageBox.Show("Session expired or not found. Please log in again.",
                                "Session Error",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                var login = new LoginWindow();
                login.Show();
                this.Close();
                return;
            }

            txtUserDisplay.Text = $"{SessionManager.CurrentUser.Username} ({SessionManager.CurrentUser.Role})";

            try
            {
                var fade = new DoubleAnimation
                {
                    From = 0,
                    To = 1,
                    Duration = TimeSpan.FromSeconds(0.9),
                    EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                };

                txtUserDisplay.BeginAnimation(OpacityProperty, fade);
            }
            catch { }

            LoadOperatorReports();
        }

        private void SessionPanel_Loaded(object sender, RoutedEventArgs e)
        {
            var panel = sender as FrameworkElement;
            if (panel == null) return;

            var slide = new ThicknessAnimation
            {
                From = new Thickness(0, -20, 0, 0),
                To = new Thickness(0, 0, 0, 0),
                Duration = TimeSpan.FromSeconds(0.6),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            var fade = new DoubleAnimation
            {
                From = 0,
                To = 1,
                Duration = TimeSpan.FromSeconds(0.8),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            panel.BeginAnimation(MarginProperty, slide);
            panel.BeginAnimation(OpacityProperty, fade);
        }

        private void LoadOperatorReports()
        {
            try
            {
                var currentUser = SessionManager.CurrentUser;
                if (currentUser == null) return;

                var all = _reportService.GetAllReports();

                _myReports = all
                    .Where(r => string.Equals(r.OperatorName, currentUser.Username, StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(r => r.DateSubmitted)
                    .ToList();

                dgMyReports.ItemsSource = null;
                dgMyReports.ItemsSource = _myReports;
            }
            catch (Exception ex)
            {
                ShowNotification($"Failed to load reports: {ex.Message}", true);
            }
        }

        private void BtnRefreshReports_Click(object sender, RoutedEventArgs e)
        {
            LoadOperatorReports();
            ShowNotification("Report list refreshed.");
        }

        private void BtnOpenReport_Click(object sender, RoutedEventArgs e)
        {
            if (dgMyReports.SelectedItem is Report selected)
            {
                string path = !string.IsNullOrWhiteSpace(selected.OperatorPath) ? selected.OperatorPath :
                              !string.IsNullOrWhiteSpace(selected.SupervisorPath) ? selected.SupervisorPath : "";

                if (string.IsNullOrWhiteSpace(path))
                {
                    ShowNotification("No file path available for this report.", true);
                    return;
                }

                if (!File.Exists(path))
                {
                    ShowNotification("Report file not found on system drive.", true);
                    return;
                }

                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    ShowNotification($"Failed to open report: {ex.Message}", true);
                }
            }
            else
            {
                ShowNotification("Please select a report to open.", true);
            }
        }

        private async void ShowNotification(string message, bool isError = false)
        {
            try
            {
                NotificationText.Text = message;

                var backgroundColor = isError
                    ? (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F8D7DA")
                    : (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#DFF0D8");

                NotificationBar.Background = new SolidColorBrush(backgroundColor);

                NotificationBar.Visibility = Visibility.Visible;

                var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(220));
                NotificationBar.BeginAnimation(OpacityProperty, fadeIn);

                await Task.Delay(2800);

                var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(300));
                fadeOut.Completed += (s, a) => NotificationBar.Visibility = Visibility.Collapsed;
                NotificationBar.BeginAnimation(OpacityProperty, fadeOut);
            }
            catch
            {
                // ignore
            }
        }

        // GAS / LIQUID calculations (kept as original; ensure GasOrificeCalculator exists)
        private void CalculateGas_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!double.TryParse(GasDP.Text, out double dp) ||
                    !double.TryParse(GasDownstream.Text, out double flp) ||
                    !double.TryParse(GasTemp.Text, out double temp) ||
                    !double.TryParse(GasSG.Text, out double sg) ||
                    !double.TryParse(GasMeterD.Text, out double meterD) ||
                    !double.TryParse(GasOrificeD.Text, out double orificeD) ||
                    !double.TryParse(GasCO2.Text, out double co2) ||
                    !double.TryParse(GasN2.Text, out double n2) ||
                    !double.TryParse(GasH2S.Text, out double h2s))
                {
                    MessageBox.Show("Please enter valid numeric values for all gas input fields.",
                        "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (dp <= 0)
                {
                    MessageBox.Show("Differential pressure must be positive.",
                        "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var calc = new GasOrificeCalculator
                {
                    DifferentialPressure_inH2O = dp,
                    DownstreamPressure_psig = flp,
                    FlowingTemperature_F = temp,
                    GasSpecificGravity = sg,
                    MeterRunDiameter_in = meterD,
                    OrificeDiameter_in = orificeD,
                    CO2_molPct = co2,
                    N2_molPct = n2,
                    H2S_ppm = h2s,
                };

                _lastGasResult = calc.Calculate();

                GasResult.Text =
                    $"Upstream P: {_lastGasResult.UpstreamPressure_psia:F3} psia\n" +
                    $"MW: {_lastGasResult.MW:F3}\n" +
                    $"Z_base: {_lastGasResult.Z_base:F6}  Z_flowing: {_lastGasResult.Z_flowing:F6}\n" +
                    $"β: {_lastGasResult.Beta:F3}  Cd: {_lastGasResult.DischargeCoefficient:F3}  Y: {_lastGasResult.ExpansionFactor_Y:F3}\n\n" +
                    $"Flowing Vol: {_lastGasResult.FlowingVol_ft3hr:F2} ft³/hr\n" +
                    $"Q (SCFH): {_lastGasResult.Qb_scfh:F2} scfh\n" +
                    $"Q (MMSCFD): {_lastGasResult.Qb_mmscfd:F6} MMSCFD";

                GasPreview.Text =
                    $"Q: {_lastGasResult.Qb_mmscfd:F6} MMSCFD | Upstream P: {_lastGasResult.UpstreamPressure_psia:F3} psia";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Gas calculation failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearGas_Click(object sender, RoutedEventArgs e)
        {
            GasDP.Clear();
            GasDownstream.Clear();
            GasTemp.Clear();
            GasSG.Text = "0.784";
            GasMeterD.Text = "4.000";
            GasOrificeD.Text = "1.750";
            GasCO2.Text = "0";
            GasN2.Text = "0";
            GasH2S.Text = "0";
            GasResult.Text = "";
            GasPreview.Text = "";
            _lastGasResult = null;
        }

        private void CalculateLiquid_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!double.TryParse(LiqMeterStart.Text, out double meterStart) ||
                    !double.TryParse(LiqMeterEnd.Text, out double meterEnd) ||
                    !double.TryParse(LiqBSW.Text, out double bswPct) ||
                    !double.TryParse(LiqShrinkage.Text, out double shrinkage) ||
                    !double.TryParse(LiqTimeFactor.Text, out double timeFactor))
                {
                    MessageBox.Show("Please enter valid numeric values for all liquid input fields.",
                        "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (meterEnd <= meterStart)
                {
                    MessageBox.Show("End Meter must be greater than Start Meter.",
                        "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                double grossVol = (meterEnd - meterStart) * timeFactor * shrinkage;
                double blpd = grossVol;
                double bopd = blpd * ((100 - bswPct) / 100.0);
                double bwpd = blpd * (bswPct / 100.0);
                double meterDiff = meterEnd - meterStart;

                _lastLiquidResult = new LiquidResult
                {
                    GrossVol = grossVol,
                    BLPD = blpd,
                    BOPD = bopd,
                    BWPD = bwpd,
                    MeterDifference = meterDiff
                };

                LiquidOutput.Text =
                    $"Gross Volume: {grossVol:F2}\nBLPD: {blpd:F2}\nBOPD: {bopd:F2}\nBWPD: {bwpd:F2}\nMeter Difference: {meterDiff:F2}";

                LiquidPreview.Text =
                    $"BOPD: {bopd:F2}, BWPD: {bwpd:F2}, ΔMeter: {meterDiff:F2}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Liquid calculation failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearLiquid_Click(object sender, RoutedEventArgs e)
        {
            LiqMeterStart.Clear();
            LiqMeterEnd.Clear();
            LiqBSW.Text = "0";
            LiqShrinkage.Text = "1.0";
            LiqTimeFactor.Text = "24";
            LiquidOutput.Text = "";
            LiquidPreview.Text = "";
            _lastLiquidResult = null;
        }

        private void ChokeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string name = (ChokeSelector?.Text ?? "").Trim();
                if (string.IsNullOrEmpty(name))
                    return;

                _selectedChoke = name;

                var presets = new Dictionary<string, (string whp, string wht, string flp, string casing, string arrival, string bsw)>
                {
                    {"10/64", ("760","110","120","30","95","0")},
                    {"12/64", ("765","112","122","31","96","0")},
                    {"14/64", ("770","113","123","31","97","0")},
                    {"16/64", ("775","114","124","32","97","0")},
                    {"18/64", ("780","115","125","32","98","0")},
                    {"20/64", ("785","116","126","33","98","0")},
                    {"24/64", ("790","116","126","33","98","0")},
                    {"28/64", ("795","117","127","34","99","0")},
                    {"32/64", ("800","118","128","34","99","0")},
                    {"36/64", ("805","119","129","35","100","0")},
                    {"40/64", ("810","120","130","35","100","0")},
                    {"44/64", ("815","121","131","36","101","0")},
                    {"48/64", ("820","122","132","36","101","0")},
                    {"52/64", ("825","123","133","37","102","0")},
                    {"56/64", ("830","124","134","37","102","0")},
                    {"60/64", ("835","125","135","38","103","0")},
                    {"64/64", ("840","126","136","38","103","0")}
                };

                if (presets.TryGetValue(name, out var p))
                {
                    txtChokeWHP.Text = p.whp;
                    txtChokeWHT.Text = p.wht;
                    txtChokeFLP.Text = p.flp;
                    txtChokeCasing.Text = p.casing;
                    txtChokeArrival.Text = p.arrival;
                    txtChokeBSW.Text = p.bsw;

                    _whp = p.whp;
                    _wht = p.wht;
                    _flp = p.flp;
                    _casing = p.casing;
                    _arrival = p.arrival;
                    _bsw = p.bsw;
                }
                else
                {
                    txtChokeWHP.Text = "";
                    txtChokeWHT.Text = "";
                    txtChokeFLP.Text = "";
                    txtChokeCasing.Text = "";
                    txtChokeArrival.Text = "";
                    txtChokeBSW.Text = "";

                    _whp = _wht = _flp = _casing = _arrival = _bsw = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load choke preset: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ApplyChokeToInputs_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _selectedChoke = ChokeSelector?.Text ?? _selectedChoke;
                _whp = txtChokeWHP.Text ?? _whp;
                _wht = txtChokeWHT.Text ?? _wht;
                _flp = txtChokeFLP.Text ?? _flp;
                _casing = txtChokeCasing.Text ?? _casing;
                _arrival = txtChokeArrival.Text ?? _arrival;
                _bsw = txtChokeBSW.Text ?? _bsw;

                MessageBox.Show("Choke constants saved independently (they will appear in exports but have NOT overwritten the Gas/Liquid inputs).",
                    "Choke Saved", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save choke constants: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveHourly_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var entry = new ReportEntry
                {
                    Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    ChokeSize = _selectedChoke ?? "",
                    ChokeWHP = _whp ?? "",
                    ChokeWHT = _wht ?? "",
                    ChokeFLP = _flp ?? "",
                    ChokeCasing = _casing ?? "",
                    ChokeArrival = _arrival ?? "",
                    ChokeBSW = _bsw ?? "",
                    ReservoirPressure = txtReservoirPressure?.Text ?? "",
                    OilAPI = txtOilAPI?.Text ?? "",
                    GasDP = GasDP?.Text ?? "",
                    GasDownstream = GasDownstream?.Text ?? "",
                    GasTemp = GasTemp?.Text ?? "",
                    GasSG = GasSG?.Text ?? "",
                    GasMeterD = GasMeterD?.Text ?? "",
                    GasOrificeD = GasOrificeD?.Text ?? "",
                    GasCO2 = GasCO2?.Text ?? "",
                    GasN2 = GasN2?.Text ?? "",
                    GasH2S = GasH2S?.Text ?? "",
                    GasUpstream = _lastGasResult != null ? _lastGasResult.UpstreamPressure_psia.ToString("F3") : "",
                    GasMW = _lastGasResult != null ? _lastGasResult.MW.ToString("F3") : "",
                    GasZBase = _lastGasResult != null ? _lastGasResult.Z_base.ToString("F6") : "",
                    GasZFlowing = _lastGasResult != null ? _lastGasResult.Z_flowing.ToString("F6") : "",
                    GasFlowVol = _lastGasResult != null ? _lastGasResult.FlowingVol_ft3hr.ToString("F2") : "",
                    GasQ = _lastGasResult != null ? _lastGasResult.Qb_mmscfd.ToString("F6") : "",
                    LiqMeterStart = LiqMeterStart?.Text ?? "",
                    LiqMeterEnd = LiqMeterEnd?.Text ?? "",
                    LiqBSW = LiqBSW?.Text ?? "",
                    LiqShrinkage = LiqShrinkage?.Text ?? "",
                    LiqTimeFactor = LiqTimeFactor?.Text ?? "",
                    LiqGrossVol = _lastLiquidResult != null ? _lastLiquidResult.GrossVol.ToString("F2") : "",
                    LiqBLPD = _lastLiquidResult != null ? _lastLiquidResult.BLPD.ToString("F2") : "",
                    LiqBOPD = _lastLiquidResult != null ? _lastLiquidResult.BOPD.ToString("F2") : "",
                    LiqBWPD = _lastLiquidResult != null ? _lastLiquidResult.BWPD.ToString("F2") : "",
                    LiqMeterDifference = _lastLiquidResult != null ? _lastLiquidResult.MeterDifference.ToString("F2") : ""
                };

                _savedEntries.Add(entry);

                MessageBox.Show($"Hourly reading saved ({entry.Timestamp}).\nSaved rows: {_savedEntries.Count}",
                    "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save hourly record: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearSaved_Click(object sender, RoutedEventArgs e)
        {
            if (_savedEntries.Count == 0)
            {
                MessageBox.Show("No saved records to clear.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var res = MessageBox.Show($"Clear all {_savedEntries.Count} saved hourly records?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (res != MessageBoxResult.Yes) return;

            _savedEntries.Clear();
        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"ProductionDaily_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };
                if (dlg.ShowDialog() != true) return;

                string remarks = "";
                try
                {
                    remarks = new TextRange(RemarksBox.Document.ContentStart, RemarksBox.Document.ContentEnd).Text.Trim();
                }
                catch { remarks = ""; }

                List<ReportEntry> toExport;
                if (_savedEntries.Any())
                    toExport = _savedEntries.ToList();
                else
                {
                    var snapshot = new ReportEntry
                    {
                        Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        ChokeSize = _selectedChoke ?? "",
                        ChokeWHP = _whp ?? "",
                        ChokeWHT = _wht ?? "",
                        ChokeFLP = _flp ?? "",
                        ChokeCasing = _casing ?? "",
                        ChokeArrival = _arrival ?? "",
                        ChokeBSW = _bsw ?? "",
                        ReservoirPressure = txtReservoirPressure?.Text ?? "",
                        OilAPI = txtOilAPI?.Text ?? "",
                        GasDP = GasDP?.Text ?? "",
                        GasDownstream = GasDownstream?.Text ?? "",
                        GasTemp = GasTemp?.Text ?? "",
                        GasSG = GasSG?.Text ?? "",
                        GasMeterD = GasMeterD?.Text ?? "",
                        GasOrificeD = GasOrificeD?.Text ?? "",
                        GasCO2 = GasCO2?.Text ?? "",
                        GasN2 = GasN2?.Text ?? "",
                        GasH2S = GasH2S?.Text ?? "",
                        GasUpstream = _lastGasResult != null ? _lastGasResult.UpstreamPressure_psia.ToString("F3") : "",
                        GasMW = _lastGasResult != null ? _lastGasResult.MW.ToString("F3") : "",
                        GasZBase = _lastGasResult != null ? _lastGasResult.Z_base.ToString("F6") : "",
                        GasZFlowing = _lastGasResult != null ? _lastGasResult.Z_flowing.ToString("F6") : "",
                        GasFlowVol = _lastGasResult != null ? _lastGasResult.FlowingVol_ft3hr.ToString("F2") : "",
                        GasQ = _lastGasResult != null ? _lastGasResult.Qb_mmscfd.ToString("F6") : "",
                        LiqMeterStart = LiqMeterStart?.Text ?? "",
                        LiqMeterEnd = LiqMeterEnd?.Text ?? "",
                        LiqBSW = LiqBSW?.Text ?? "",
                        LiqShrinkage = LiqShrinkage?.Text ?? "",
                        LiqTimeFactor = LiqTimeFactor?.Text ?? "",
                        LiqGrossVol = _lastLiquidResult != null ? _lastLiquidResult.GrossVol.ToString("F2") : "",
                        LiqBLPD = _lastLiquidResult != null ? _lastLiquidResult.BLPD.ToString("F2") : "",
                        LiqBOPD = _lastLiquidResult != null ? _lastLiquidResult.BOPD.ToString("F2") : "",
                        LiqBWPD = _lastLiquidResult != null ? _lastLiquidResult.BWPD.ToString("F2") : "",
                        LiqMeterDifference = _lastLiquidResult != null ? _lastLiquidResult.MeterDifference.ToString("F2") : ""
                    };

                    toExport = new List<ReportEntry> { snapshot };
                }

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("DailyReport");

                    // optional logo (assets/logo.png)
                    try
                    {
                        string logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "logo.png");
                        if (File.Exists(logoPath))
                        {
                            var pic = ws.Drawings.AddPicture("CompanyLogo", new FileInfo(logoPath));
                            pic.SetPosition(0, 0, 0, 0);
                            pic.SetSize(90, 90);
                        }
                    }
                    catch { }

                    var headers = new List<string>
                    {
                        "Timestamp",
                        "Choke Size","WHP (psia)","WHT (°F)","FLP (psig)","Casing (psig)","Arrival Manifold (psig)","Wellhead BS&W (%)",
                        "Reservoir Pressure (psi)","Oil API",
                        "ΔP (inH₂O)","Downstream Pressure (psig)","Flowing Temp (°F)","Gas Specific Gravity","Meter Run Diameter (in)","Orifice Diameter (in)","CO₂ (%)","N₂ (%)","H₂S (ppm)",
                        "Upstream Pressure (psia)","MW","Z_base","Z_flowing","Flowing Vol (ft³/hr)","Q (MMSCFD)",
                        "Start Meter (V1)","End Meter (V2)","BS&W (%)","Shrinkage Factor","Time Factor (hrs)",
                        "Gross Volume","BLPD","BOPD","BWPD","Meter Difference",
                        "Exported At"
                    };

                    for (int c = 0; c < headers.Count; c++)
                    {
                        ws.Cells[1, c + 1].Value = headers[c];
                        ws.Cells[1, c + 1].Style.Font.Bold = true;
                        ws.Cells[1, c + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[1, c + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        ws.Cells[1, c + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[1, c + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    int row = 2;
                    foreach (var eRow in toExport)
                    {
                        int col = 1;
                        ws.Cells[row, col++].Value = eRow.Timestamp;
                        ws.Cells[row, col++].Value = eRow.ChokeSize;
                        ws.Cells[row, col++].Value = eRow.ChokeWHP;
                        ws.Cells[row, col++].Value = eRow.ChokeWHT;
                        ws.Cells[row, col++].Value = eRow.ChokeFLP;
                        ws.Cells[row, col++].Value = eRow.ChokeCasing;
                        ws.Cells[row, col++].Value = eRow.ChokeArrival;
                        ws.Cells[row, col++].Value = eRow.ChokeBSW;
                        ws.Cells[row, col++].Value = eRow.ReservoirPressure;
                        ws.Cells[row, col++].Value = eRow.OilAPI;
                        ws.Cells[row, col++].Value = eRow.GasDP;
                        ws.Cells[row, col++].Value = eRow.GasDownstream;
                        ws.Cells[row, col++].Value = eRow.GasTemp;
                        ws.Cells[row, col++].Value = eRow.GasSG;
                        ws.Cells[row, col++].Value = eRow.GasMeterD;
                        ws.Cells[row, col++].Value = eRow.GasOrificeD;
                        ws.Cells[row, col++].Value = eRow.GasCO2;
                        ws.Cells[row, col++].Value = eRow.GasN2;
                        ws.Cells[row, col++].Value = eRow.GasH2S;
                        ws.Cells[row, col++].Value = eRow.GasUpstream;
                        ws.Cells[row, col++].Value = eRow.GasMW;
                        ws.Cells[row, col++].Value = eRow.GasZBase;
                        ws.Cells[row, col++].Value = eRow.GasZFlowing;
                        ws.Cells[row, col++].Value = eRow.GasFlowVol;
                        ws.Cells[row, col++].Value = eRow.GasQ;
                        ws.Cells[row, col++].Value = eRow.LiqMeterStart;
                        ws.Cells[row, col++].Value = eRow.LiqMeterEnd;
                        ws.Cells[row, col++].Value = eRow.LiqBSW;
                        ws.Cells[row, col++].Value = eRow.LiqShrinkage;
                        ws.Cells[row, col++].Value = eRow.LiqTimeFactor;
                        ws.Cells[row, col++].Value = eRow.LiqGrossVol;
                        ws.Cells[row, col++].Value = eRow.LiqBLPD;
                        ws.Cells[row, col++].Value = eRow.LiqBOPD;
                        ws.Cells[row, col++].Value = eRow.LiqBWPD;
                        ws.Cells[row, col++].Value = eRow.LiqMeterDifference;
                        ws.Cells[row, col++].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        row++;
                    }

                    int lastDataRow = row - 1;
                    int lastCol = headers.Count;

                    var dataRange = ws.Cells[1, 1, lastDataRow, lastCol];
                    dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    int remarksRow = lastDataRow + 2;
                    ws.Cells[remarksRow, 1, remarksRow, lastCol].Merge = true;
                    ws.Cells[remarksRow, 1].Value = "Remarks:\n" + (string.IsNullOrWhiteSpace(remarks) ? "(none)" : remarks);
                    ws.Cells[remarksRow, 1].Style.WrapText = true;
                    ws.Row(remarksRow).Height = 120;
                    ws.Cells[remarksRow, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                    ws.Cells[1, 1, remarksRow, lastCol].AutoFitColumns();

                    File.WriteAllBytes(dlg.FileName, package.GetAsByteArray());
                }
                _lastExportedPath = dlg.FileName;
                ShowNotification("Exported to Excel successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void SubmitToSupervisor_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Check if there is an exported report
                if (string.IsNullOrWhiteSpace(_lastExportedPath) || !File.Exists(_lastExportedPath))
                {
                    ShowNotification("Please export a report before submitting.", true);
                    return;
                }

                // Create PendingApproval folder
                string pendingDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "ProductionCalc",
                    "PendingApproval");

                Directory.CreateDirectory(pendingDir);

                // Copy exported Excel file to PendingApproval
                string pendingPath = Path.Combine(pendingDir, Path.GetFileName(_lastExportedPath));
                File.Copy(_lastExportedPath, pendingPath, true);

                // Get remarks
                string remarks = "";
                try
                {
                    remarks = new TextRange(RemarksBox.Document.ContentStart,
                                            RemarksBox.Document.ContentEnd).Text.Trim();
                }
                catch { }

                // Create report object
                var report = new Report
                {
                    ReportTitle = Path.GetFileNameWithoutExtension(_lastExportedPath),
                    OperatorPath = _lastExportedPath,
                    SupervisorPath = pendingPath,
                    OperatorName = SessionManager.CurrentUser?.Username ?? "Unknown",
                    SupervisorComment = "",
                    Status = "Pending",
                    DateSubmitted = DateTime.Now,
                    Remarks = remarks
                };

                // Save report to database/service
                _reportService.SaveReport(report);

                // Reload operator reports in UI
                LoadOperatorReports();

                ShowNotification("Report successfully submitted to Supervisor.");
            }
            catch (Exception ex)
            {
                ShowNotification($"Submission failed: {ex.Message}", true);
            }
        }



        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            var confirm = MessageBox.Show("Are you sure you want to log out?", "Confirm Logout", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;

            SessionManager.EndSession();
            var login = new LoginWindow();
            login.Show();
            this.Close();
        }
    }
}
