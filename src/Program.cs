using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Microsoft.CSharp.RuntimeBinder;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

[assembly: AssemblyTitle("SigaExcel")]
[assembly: AssemblyProduct("SigaExcel")]
[assembly: AssemblyDescription("Versao atualizada do rezende.exe para automacao Excel")]
[assembly: AssemblyCompany("Victor Cardoso")]
[assembly: AssemblyCopyright("Atualizacao por Victor Cardoso - vhsbc92@gmail.com")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace SigaExcel
{
    internal static class Program
    {
        private const string AppVersion = "v1";
        private const string SupportContact = "Victor Cardoso - vhsbc92@gmail.com";

        private const int XlContinuous = 1;
        private const int XlThin = 2;
        private const int XlLandscape = 2;
        private const int XlPortrait = 1;
        private const int XlCalculationAutomatic = -4105;
        private const int XlCalculationManual = -4135;

        private static dynamic _excel;
        private static dynamic _workbook;
        private static Options _options;
        private static ExcelPerformanceState _performanceState;
        private static ProgressOverlay _progressOverlay;

        private static int Main(string[] args)
        {
            if (args.Length == 0 || args[0] == "--help" || args[0] == "/?")
            {
                ShowInfo(BuildUsageText(), "SigaExcel");
                return args.Length == 0 ? 1 : 0;
            }

            if (args[0] == "--test-error")
            {
                ShowError("Erro de teste do SigaExcel.");
                return 9;
            }

            Options options;
            try
            {
                options = Options.Parse(args);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message + "\r\n\r\n" + BuildUsageText());
                return 1;
            }

            _options = options;
            var scriptPath = options.ScriptPath;
            if (!File.Exists(scriptPath))
            {
                ShowError("Arquivo de script nao encontrado:\r\n" + scriptPath);
                return 2;
            }

            try
            {
                var commands = new List<ScriptCommand>(ScriptParser.ParseFile(scriptPath));

                if (!options.DryRun && options.FastMode)
                    StartProgressOverlay(scriptPath, commands.Count);

                for (var index = 0; index < commands.Count; index++)
                {
                    var command = commands[index];
                    if (!options.DryRun && options.FastMode)
                        UpdateProgressOverlay(index + 1, command);

                    if (options.DryRun)
                        PrintDryRun(command);
                    else
                        ExecuteWithRetry(command);
                }

                if (!options.DryRun && options.ModernTheme)
                    ApplyModernTheme();

                if (!options.DryRun && !string.IsNullOrWhiteSpace(options.OutputPath))
                    ActiveWorkbook().SaveAs(options.OutputPath);

                if (!options.DryRun && options.FastMode)
                    EndFastMode(showExcel: true);

                StopProgressOverlay("Concluido");

                return 0;
            }
            catch (Exception ex)
            {
                if (!options.DryRun && options.FastMode)
                    EndFastMode(showExcel: true);

                StopProgressOverlay("Erro");

                ShowError(ex.Message);
                return 3;
            }
        }

        private static string BuildUsageText()
        {
            return
                "SigaExcel " + AppVersion + " - interpretador compativel de scripts Excel\r\n" +
                "Versao atualizada do rezende.exe por Victor Cardoso.\r\n" +
                "Contato: vhsbc92@gmail.com\r\n" +
                "Atualizado devido as mudancas da Microsoft que afetaram a automacao Excel/OLE.\r\n" +
                "Ferramenta em evolucao para compatibilidade, robustez e novos recursos.\r\n\r\n" +
                "Uso:\r\n" +
                "  sigaexcel.exe C:\\sigaexcel\\scripts\\20260617.txt\r\n" +
                "  sigaexcel.exe --dry-run C:\\sigaexcel\\scripts\\20260617.txt\r\n" +
                "  sigaexcel.exe --fast C:\\sigaexcel\\scripts\\20260617.txt\r\n" +
                "  sigaexcel.exe --modern C:\\sigaexcel\\scripts\\20260617.txt\r\n" +
                "  sigaexcel.exe --fast --modern C:\\sigaexcel\\scripts\\20260617.txt\r\n" +
                "  sigaexcel.exe --modern --output C:\\sigaexcel\\planilhas\\modern.xlsx C:\\sigaexcel\\scripts\\20260617.txt\r\n" +
                "  sigaexcel.exe --test-error";
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(
                message + "\r\n\r\nQualquer problema, entre em contato: " + SupportContact,
                "SigaExcel - erro",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        private static void ShowInfo(string message, string title)
        {
            MessageBox.Show(
                message,
                title,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static void StartProgressOverlay(string scriptPath, int totalCommands)
        {
            _progressOverlay = new ProgressOverlay();
            _progressOverlay.Start(scriptPath, totalCommands);
        }

        private static void UpdateProgressOverlay(int current, ScriptCommand command)
        {
            if (_progressOverlay == null)
                return;

            _progressOverlay.Update(current, command.LineNumber, command.Name);
        }

        private static void StopProgressOverlay(string status)
        {
            if (_progressOverlay == null)
                return;

            _progressOverlay.Stop(status);
            _progressOverlay = null;
        }

        private static void PrintDryRun(ScriptCommand command)
        {
            Console.WriteLine(command.LineNumber.ToString(CultureInfo.InvariantCulture) + ": " + command.Name + " (" + command.Parameters.Count + " parametros)");
            if (string.Equals(command.Name, "Escreve", StringComparison.OrdinalIgnoreCase) && command.Parameters.Count == 13)
            {
                Console.WriteLine("   row=" + command.Parameters[0] +
                                  " col=" + command.Parameters[1] +
                                  " borda=" + command.Parameters[10] +
                                  " corBorda=" + command.Parameters[11] +
                                  " corFundo=" + command.Parameters[12]);
            }
        }

        private static void Execute(ScriptCommand command)
        {
            try
            {
                switch (command.Name.ToUpperInvariant())
                {
                    case "ABREEXCEL":
                        AbreExcel(command);
                        break;
                    case "ABREARQUIVO":
                        AbreArquivo(command);
                        break;
                    case "ADICIONASHEET":
                        AdicionaSheet(command);
                        break;
                    case "ESCREVE":
                        Escreve(command, formula: false);
                        break;
                    case "ESCREVEFORMULA":
                        Escreve(command, formula: true);
                        break;
                    case "MERGECELULAS":
                        MergeCelulas(command);
                        break;
                    case "FORMATANUMEROCELULA":
                        FormataNumeroCelula(command);
                        break;
                    case "AUTOFORMATA":
                        AutoFormata(command);
                        break;
                    case "CONFIGURAPAGINA":
                        ConfiguraPagina(command);
                        break;
                    case "RENOMEIASHEET":
                        RenomeiaSheet(command);
                        break;
                    case "GRAVAARQUIVO":
                        GravaArquivo(command);
                        break;
                    case "FECHAEXCEL":
                        FechaExcel();
                        break;
                    case "IMPRIME":
                        ActiveWorkbook().PrintOut();
                        break;
                    case "VISUALIZAIMPRESSAO":
                        ActiveWorkbook().PrintPreview();
                        break;
                    default:
                        throw command.Error("Comando nao suportado nesta versao: " + command.Name);
                }
            }
            catch (Exception ex)
            {
                throw command.Error(ex.Message, ex);
            }
        }

        private static void ExecuteWithRetry(ScriptCommand command)
        {
            const int maxAttempts = 12;

            for (var attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    Execute(command);
                    return;
                }
                catch (Exception ex)
                {
                    if (attempt == maxAttempts || !IsExcelBusy(ex))
                        throw;

                    Thread.Sleep(250 * attempt);
                }
            }
        }

        private static bool IsExcelBusy(Exception ex)
        {
            while (ex != null)
            {
                var com = ex as COMException;
                if (com != null && com.HResult == unchecked((int)0x8001010A))
                    return true;

                if (ex.Message != null && ex.Message.IndexOf("application is busy", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;

                ex = ex.InnerException;
            }

            return false;
        }

        private static void AbreExcel(ScriptCommand command)
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw command.Error("Excel.Application nao esta registrado neste Windows.");

            _excel = Activator.CreateInstance(excelType);

            var visible = true;
            if (command.Parameters.Count >= 1 && !string.IsNullOrWhiteSpace(command.Parameters[0]))
                visible = IsYes(command.Parameters[0]);

            if (_options != null && _options.FastMode)
                visible = false;

            _excel.Visible = visible;

            if (_options != null && _options.FastMode)
                BeginFastMode();
        }

        private static void BeginFastMode()
        {
            if (_excel == null || _performanceState != null)
                return;

            _performanceState = new ExcelPerformanceState();

            try { _performanceState.ScreenUpdating = _excel.ScreenUpdating; } catch { }
            try { _performanceState.EnableEvents = _excel.EnableEvents; } catch { }
            try { _performanceState.DisplayAlerts = _excel.DisplayAlerts; } catch { }
            try { _performanceState.Calculation = _excel.Calculation; } catch { }
            try { _performanceState.Visible = _excel.Visible; } catch { }

            try { _excel.ScreenUpdating = false; } catch { }
            try { _excel.EnableEvents = false; } catch { }
            try { _excel.DisplayAlerts = false; } catch { }
            try { _excel.Calculation = XlCalculationManual; } catch { }
            try { _excel.Visible = false; } catch { }
        }

        private static void EndFastMode(bool showExcel)
        {
            if (_excel == null || _performanceState == null)
                return;

            try { _excel.ScreenUpdating = _performanceState.ScreenUpdating; } catch { }
            try { _excel.EnableEvents = _performanceState.EnableEvents; } catch { }
            try { _excel.DisplayAlerts = _performanceState.DisplayAlerts; } catch { }
            try { _excel.Calculation = _performanceState.Calculation ?? XlCalculationAutomatic; } catch { }
            try { _excel.Visible = showExcel || _performanceState.Visible; } catch { }

            _performanceState = null;
        }

        private static void AbreArquivo(ScriptCommand command)
        {
            RequireExcel(command);
            command.RequireParamCountAtLeast(1);

            var path = command.Parameters[0];
            _workbook = _excel.Workbooks.Open(path);
        }

        private static void AdicionaSheet(ScriptCommand command)
        {
            RequireExcel(command);
            command.RequireParamCountAtLeast(1);

            if (ActiveWorkbookOrNull() == null)
            {
                _workbook = CreateWorkbookWithOneSheet();
                return;
            }

            _excel.Worksheets.Add();
        }

        private static dynamic CreateWorkbookWithOneSheet()
        {
            object previous = null;

            try
            {
                previous = _excel.SheetsInNewWorkbook;
                _excel.SheetsInNewWorkbook = 1;
                return _excel.Workbooks.Add();
            }
            finally
            {
                if (previous != null)
                    _excel.SheetsInNewWorkbook = previous;
            }
        }

        private static void Escreve(ScriptCommand command, bool formula)
        {
            RequireExcel(command);

            if (command.Parameters.Count != 3 && command.Parameters.Count != 11 && command.Parameters.Count != 13)
                throw command.Error("Quantidade invalida de parametros para " + command.Name + ": " + command.Parameters.Count);

            var row = ParseInt(command.Parameters[0], "Linha", command);
            var col = ParseInt(command.Parameters[1], "Coluna", command);
            var content = command.Parameters[2];

            if (command.Parameters.Count == 3)
            {
                var cellOnly = ActiveSheet().Cells[row, col];
                if (formula)
                    cellOnly.FormulaLocal = content;
                else
                    cellOnly.Value = content;
                return;
            }

            var sheet = ParseInt(command.Parameters[3], "Sheet", command);
            var fontName = command.Parameters[4];
            var fontSize = ParseDouble(command.Parameters[5], "Tamanho", command);
            var bold = IsYes(command.Parameters[6]);
            var italic = IsYes(command.Parameters[7]);
            var underline = IsYes(command.Parameters[8]);
            var fontColor = ParseInt(command.Parameters[9], "Cor da Letra", command);
            var hasBorder = IsYes(command.Parameters[10]);

            var ws = Sheet(sheet);
            var cell = ws.Cells[row, col];

            if (formula)
                cell.FormulaLocal = content;
            else
                cell.Value = content;

            cell.Font.Name = string.IsNullOrEmpty(fontName) ? "Arial" : fontName;
            cell.Font.Size = fontSize;
            cell.Font.Bold = bold;
            cell.Font.Italic = italic;
            cell.Font.Underline = underline;
            cell.Font.Color = fontColor;

            if (hasBorder)
            {
                if (command.Parameters.Count >= 12 && !string.IsNullOrWhiteSpace(command.Parameters[11]))
                {
                    var borderColorIndex = ParseColorIndex(command.Parameters[11], "Cor da Borda", command);
                    cell.BorderAround(XlContinuous, XlThin, borderColorIndex, Type.Missing);
                    cell.Borders.ColorIndex = borderColorIndex;
                }
                else
                {
                    cell.BorderAround();
                }
            }

            if (command.Parameters.Count >= 13 && !string.IsNullOrWhiteSpace(command.Parameters[12]))
            {
                var fillColorIndex = ParseColorIndex(command.Parameters[12], "Cor de Fundo", command);
                cell.Interior.ColorIndex = fillColorIndex;
            }
        }

        private static void MergeCelulas(ScriptCommand command)
        {
            RequireExcel(command);
            command.RequireParamCountAtLeast(5);

            var row1 = ParseInt(command.Parameters[0], "Linha Inicial", command);
            var col1 = ParseInt(command.Parameters[1], "Coluna Inicial", command);
            var row2 = ParseInt(command.Parameters[2], "Linha Final", command);
            var col2 = ParseInt(command.Parameters[3], "Coluna Final", command);
            var sheet = ParseInt(command.Parameters[4], "Sheet", command);

            var ws = Sheet(sheet);
            var range = ws.Range[ws.Cells[row1, col1], ws.Cells[row2, col2]];
            range.Merge();
        }

        private static void FormataNumeroCelula(ScriptCommand command)
        {
            RequireExcel(command);
            command.RequireParamCountAtLeast(5);

            var row = ParseInt(command.Parameters[0], "Linha", command);
            var col = ParseInt(command.Parameters[1], "Coluna", command);
            var integerPlaces = ParseInt(command.Parameters[2], "Numero de Casas", command);
            var decimals = ParseInt(command.Parameters[3], "Numero de Decimais", command);
            var sheet = ParseInt(command.Parameters[4], "Sheet", command);

            Sheet(sheet).Cells[row, col].NumberFormat = BuildNumberFormat(integerPlaces, decimals);
        }

        private static void AutoFormata(ScriptCommand command)
        {
            RequireExcel(command);
            command.RequireParamCountAtLeast(3);

            var col1 = ParseInt(command.Parameters[0], "Coluna Inicial", command);
            var col2 = ParseInt(command.Parameters[1], "Coluna Final", command);
            var sheet = ParseInt(command.Parameters[2], "Sheet", command);

            var ws = Sheet(sheet);
            ws.Range[ws.Cells[1, col1], ws.Cells[1, col2]].EntireColumn.AutoFit();
        }

        private static void ConfiguraPagina(ScriptCommand command)
        {
            RequireExcel(command);
            command.RequireParamCountAtLeast(2);

            var orientation = command.Parameters[0].Trim().ToUpperInvariant();
            var sheet = ParseInt(command.Parameters[1], "Sheet", command);
            Sheet(sheet).PageSetup.Orientation = orientation == "P" ? XlLandscape : XlPortrait;
        }

        private static void RenomeiaSheet(ScriptCommand command)
        {
            RequireExcel(command);
            command.RequireParamCountAtLeast(2);

            var sheet = ParseInt(command.Parameters[0], "Sheet", command);
            Sheet(sheet).Name = command.Parameters[1];
        }

        private static void GravaArquivo(ScriptCommand command)
        {
            RequireExcel(command);
            command.RequireParamCountAtLeast(1);

            ActiveWorkbook().SaveAs(command.Parameters[0]);
        }

        private static void ApplyModernTheme()
        {
            var wb = ActiveWorkbook();
            var count = wb.Sheets.Count;

            for (var i = 1; i <= count; i++)
            {
                ApplyModernThemeToSheet(wb.Sheets[i]);
            }
        }

        private static void ApplyModernThemeToSheet(dynamic ws)
        {
            var used = ws.UsedRange;
            used.Font.Name = "Segoe UI";
            used.Font.Size = 9;
            used.VerticalAlignment = -4108; // xlCenter
            used.WrapText = true;

            // Base limpa, com grade leve.
            used.Interior.Color = Rgb(255, 255, 255);
            used.Borders.LineStyle = XlContinuous;
            used.Borders.Weight = XlThin;
            used.Borders.Color = Rgb(214, 220, 229);

            ws.Columns["A:J"].ColumnWidth = 14;
            ws.Columns["A:A"].ColumnWidth = 18;
            ws.Columns["B:E"].ColumnWidth = 16;
            ws.Columns["F:F"].ColumnWidth = 8;
            ws.Columns["G:J"].ColumnWidth = 14;

            ws.Rows["1:30"].RowHeight = 21;
            ws.Rows["1:1"].RowHeight = 28;
            ws.Rows["6:6"].RowHeight = 26;
            ws.Rows["18:18"].RowHeight = 24;
            ws.Rows["23:23"].RowHeight = 24;
            ws.Rows["27:30"].RowHeight = 26;

            StyleBand(ws.Range["A1:J1"], Rgb(23, 49, 79), Rgb(255, 255, 255), 12, true);
            StyleBand(ws.Range["A2:J5"], Rgb(241, 245, 249), Rgb(15, 23, 42), 9, false);
            StyleBand(ws.Range["A6:J6"], Rgb(14, 116, 144), Rgb(255, 255, 255), 10, true);
            StyleBand(ws.Range["F9:J9"], Rgb(30, 64, 175), Rgb(255, 255, 255), 9, true);
            StyleBand(ws.Range["A18:J18"], Rgb(30, 64, 175), Rgb(255, 255, 255), 9, true);
            StyleBand(ws.Range["A23:J23"], Rgb(30, 64, 175), Rgb(255, 255, 255), 9, true);
            StyleBand(ws.Range["A27:J27"], Rgb(15, 23, 42), Rgb(255, 255, 255), 9, true);

            StyleDataArea(ws.Range["A7:J17"]);
            StyleDataArea(ws.Range["A19:J22"]);
            StyleDataArea(ws.Range["A24:J26"]);
            StyleDataArea(ws.Range["A28:J30"]);

            ws.Range["A1:J30"].HorizontalAlignment = -4108; // xlCenter
            ws.Range["B:B"].HorizontalAlignment = -4131; // xlLeft
            ws.Range["A1:J30"].Borders.Color = Rgb(203, 213, 225);

            ws.PageSetup.Orientation = XlLandscape;
            ws.PageSetup.Zoom = false;
            ws.PageSetup.FitToPagesWide = 1;
            ws.PageSetup.FitToPagesTall = 1;
        }

        private static void StyleBand(dynamic range, int background, int foreground, int size, bool bold)
        {
            range.Interior.Color = background;
            range.Font.Color = foreground;
            range.Font.Bold = bold;
            range.Font.Size = size;
            range.Borders.Color = Rgb(148, 163, 184);
        }

        private static void StyleDataArea(dynamic range)
        {
            range.Interior.Color = Rgb(255, 255, 255);
            range.Font.Color = Rgb(15, 23, 42);
            range.Borders.Color = Rgb(226, 232, 240);
        }

        private static int Rgb(int red, int green, int blue)
        {
            return red | (green << 8) | (blue << 16);
        }

        private static void FechaExcel()
        {
            if (_excel != null)
            {
                _excel.Quit();
                ReleaseCom(_excel);
                _excel = null;
                _workbook = null;
            }
        }

        private static dynamic Sheet(int sheet)
        {
            return ActiveWorkbook().Sheets[sheet];
        }

        private static dynamic ActiveSheet()
        {
            return ActiveWorkbook().ActiveSheet;
        }

        private static dynamic ActiveWorkbook()
        {
            RequireExcel(null);
            return _workbook ?? _excel.ActiveWorkbook;
        }

        private static dynamic ActiveWorkbookOrNull()
        {
            try
            {
                if (_workbook != null)
                    return _workbook;
                return _excel.ActiveWorkbook;
            }
            catch (COMException)
            {
                return null;
            }
            catch (RuntimeBinderException)
            {
                return null;
            }
        }

        private static void RequireExcel(ScriptCommand command)
        {
            if (_excel == null)
            {
                var message = "Excel ainda nao foi aberto. Execute AbreExcel antes.";
                if (command != null) throw command.Error(message);
                throw new InvalidOperationException(message);
            }
        }

        private static bool IsYes(string value)
        {
            return string.Equals(value.Trim(), "S", StringComparison.OrdinalIgnoreCase);
        }

        private static int ParseInt(string value, string fieldName, ScriptCommand command)
        {
            int result;
            if (!int.TryParse(value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out result))
                throw command.Error("Parametro numerico invalido em " + fieldName + ": '" + value + "'");
            return result;
        }

        private static int ParseColorIndex(string value, string fieldName, ScriptCommand command)
        {
            var result = ParseInt(value, fieldName, command);
            if (result == -4142 || result == -4105 || (result >= 1 && result <= 56))
                return result;

            throw command.Error(fieldName + " fora do intervalo de ColorIndex suportado: " + value);
        }

        private static double ParseDouble(string value, string fieldName, ScriptCommand command)
        {
            double result;
            if (double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out result))
                return result;

            if (double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out result))
                return result;

            throw command.Error("Parametro numerico invalido em " + fieldName + ": '" + value + "'");
        }

        private static string BuildNumberFormat(int integerPlaces, int decimals)
        {
            var left = integerPlaces > 0 ? new string('#', integerPlaces) : "#";
            if (decimals <= 0)
                return left;
            return left + "." + new string('0', decimals);
        }

        private static void ReleaseCom(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
                Marshal.FinalReleaseComObject(obj);
        }
    }

    internal sealed class ScriptCommand
    {
        public ScriptCommand(string file, int lineNumber, string raw, string name, IReadOnlyList<string> parameters)
        {
            File = file;
            LineNumber = lineNumber;
            Raw = raw;
            Name = name;
            Parameters = parameters;
        }

        public string File { get; private set; }
        public int LineNumber { get; private set; }
        public string Raw { get; private set; }
        public string Name { get; private set; }
        public IReadOnlyList<string> Parameters { get; private set; }

        public void RequireParamCountAtLeast(int count)
        {
            if (Parameters.Count < count)
                throw Error("Parametros insuficientes. Esperado ao menos " + count + ", recebido " + Parameters.Count + ".");
        }

        public Exception Error(string message, Exception inner = null)
        {
            return new InvalidOperationException(
                "Erro ao executar script.\r\n" +
                "Arquivo: " + File + "\r\n" +
                "Linha: " + LineNumber + "\r\n" +
                "Comando: " + Name + "\r\n" +
                "Mensagem: " + message + "\r\n" +
                "Texto: " + Raw,
                inner);
        }
    }

    internal sealed class Options
    {
        public bool DryRun { get; private set; }
        public bool ModernTheme { get; private set; }
        public bool FastMode { get; private set; }
        public string OutputPath { get; private set; }
        public string ScriptPath { get; private set; }

        public static Options Parse(string[] args)
        {
            var options = new Options();

            for (var i = 0; i < args.Length; i++)
            {
                var arg = args[i];

                if (string.Equals(arg, "--dry-run", StringComparison.OrdinalIgnoreCase))
                {
                    options.DryRun = true;
                    continue;
                }

                if (string.Equals(arg, "--modern", StringComparison.OrdinalIgnoreCase))
                {
                    options.ModernTheme = true;
                    continue;
                }

                if (string.Equals(arg, "--fast", StringComparison.OrdinalIgnoreCase))
                {
                    options.FastMode = true;
                    continue;
                }

                if (string.Equals(arg, "--output", StringComparison.OrdinalIgnoreCase))
                {
                    if (i + 1 >= args.Length)
                        throw new InvalidOperationException("Informe o caminho apos --output.");

                    options.OutputPath = args[++i];
                    continue;
                }

                if (arg.StartsWith("--", StringComparison.Ordinal))
                    throw new InvalidOperationException("Opcao desconhecida: " + arg);

                if (options.ScriptPath != null)
                    throw new InvalidOperationException("Informe apenas um arquivo de script.");

                options.ScriptPath = arg;
            }

            if (options.ScriptPath == null)
                throw new InvalidOperationException("Arquivo de script nao informado.");

            return options;
        }
    }

    internal sealed class ExcelPerformanceState
    {
        public bool ScreenUpdating { get; set; }
        public bool EnableEvents { get; set; }
        public bool DisplayAlerts { get; set; }
        public int? Calculation { get; set; }
        public bool Visible { get; set; }
    }

    internal sealed class ProgressOverlay
    {
        private Thread _thread;
        private ProgressForm _form;
        private readonly ManualResetEventSlim _ready = new ManualResetEventSlim(false);

        public void Start(string scriptPath, int totalCommands)
        {
            _thread = new Thread(() =>
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                _form = new ProgressForm(scriptPath, totalCommands);
                _ready.Set();
                Application.Run(_form);
            });

            _thread.SetApartmentState(ApartmentState.STA);
            _thread.IsBackground = true;
            _thread.Start();
            _ready.Wait(3000);
        }

        public void Update(int current, int lineNumber, string commandName)
        {
            var form = _form;
            if (form == null || form.IsDisposed)
                return;

            try
            {
                form.BeginInvoke((Action)(() => form.SetProgress(current, lineNumber, commandName)));
            }
            catch
            {
            }
        }

        public void Stop(string status)
        {
            var form = _form;
            if (form == null || form.IsDisposed)
                return;

            try
            {
                form.Invoke((Action)(() =>
                {
                    form.SetStatus(status);
                    form.Close();
                }));
            }
            catch
            {
            }
        }
    }

    internal sealed class ProgressForm : Form
    {
        private readonly int _totalCommands;
        private readonly Label _title;
        private readonly Label _script;
        private readonly Label _status;
        private readonly ProgressBar _progress;

        public ProgressForm(string scriptPath, int totalCommands)
        {
            _totalCommands = Math.Max(totalCommands, 1);

            Text = "SigaExcel";
            Width = 520;
            Height = 178;
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = true;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = System.Drawing.Color.White;

            _title = new Label
            {
                Left = 18,
                Top = 14,
                Width = 470,
                Height = 24,
                Text = "SigaExcel esta gerando o relatorio",
                Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold)
            };

            _script = new Label
            {
                Left = 18,
                Top = 42,
                Width = 470,
                Height = 20,
                Text = "Script: " + Path.GetFileName(scriptPath),
                Font = new System.Drawing.Font("Segoe UI", 9)
            };

            _status = new Label
            {
                Left = 18,
                Top = 68,
                Width = 470,
                Height = 20,
                Text = "Preparando o Excel em modo rapido...",
                Font = new System.Drawing.Font("Segoe UI", 9)
            };

            _progress = new ProgressBar
            {
                Left = 18,
                Top = 96,
                Width = 470,
                Height = 18,
                Minimum = 0,
                Maximum = _totalCommands,
                Value = 0,
                Style = ProgressBarStyle.Continuous
            };

            var footer = new Label
            {
                Left = 18,
                Top = 122,
                Width = 470,
                Height = 20,
                Text = "Suporte: Victor Cardoso - vhsbc92@gmail.com",
                Font = new System.Drawing.Font("Segoe UI", 8),
                ForeColor = System.Drawing.Color.FromArgb(71, 85, 105)
            };

            Controls.Add(_title);
            Controls.Add(_script);
            Controls.Add(_status);
            Controls.Add(_progress);
            Controls.Add(footer);
        }

        public void SetProgress(int current, int lineNumber, string commandName)
        {
            var safeCurrent = Math.Min(Math.Max(current, 0), _totalCommands);
            _progress.Value = safeCurrent;
            _status.Text = "Executando linha " + lineNumber.ToString(CultureInfo.InvariantCulture) +
                           " | comando " + commandName +
                           " | progresso " + safeCurrent.ToString(CultureInfo.InvariantCulture) +
                           "/" + _totalCommands.ToString(CultureInfo.InvariantCulture);
        }

        public void SetStatus(string status)
        {
            _status.Text = status;
        }
    }

    internal static class ScriptParser
    {
        public static IEnumerable<ScriptCommand> ParseFile(string path)
        {
            var text = File.ReadAllText(path, Encoding.Default);
            var commands = SplitCommands(text);
            var logicalLine = 0;

            foreach (var rawCommand in commands)
            {
                logicalLine++;
                var raw = rawCommand.TrimEnd('\r', '\n');
                if (string.IsNullOrWhiteSpace(raw))
                    continue;

                var delimiterIndex = FirstDelimiterIndex(raw);
                if (delimiterIndex < 0)
                    continue;

                var delimiter = raw[delimiterIndex];
                var name = raw.Substring(0, delimiterIndex).Trim();
                var rest = raw.Substring(delimiterIndex + 1);
                var parts = new List<string>(rest.Split(delimiter));

                while (parts.Count > 0 && parts[parts.Count - 1] == string.Empty)
                    parts.RemoveAt(parts.Count - 1);

                yield return new ScriptCommand(path, logicalLine, raw, name, parts);
            }
        }

        private static IEnumerable<string> SplitCommands(string text)
        {
            if (text.IndexOf('\r') < 0)
            {
                foreach (var part in text.Split('\n'))
                    yield return part;
                yield break;
            }

            text = text.Replace("\r\n", "\r");
            var parts = text.Split('\r');
            foreach (var part in parts)
                yield return part;
        }

        private static int FirstDelimiterIndex(string raw)
        {
            var pipe = raw.IndexOf('|');
            var semicolon = raw.IndexOf(';');

            if (pipe < 0) return semicolon;
            if (semicolon < 0) return pipe;
            return Math.Min(pipe, semicolon);
        }
    }
}
