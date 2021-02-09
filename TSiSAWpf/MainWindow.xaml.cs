using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace TSiSAWpf
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Excel.Application Excel;
        private Workbook WorkBook;
        private Worksheet OperationsSheet;
        private Worksheet InvoicesSheet;
        private Worksheet OSVSheet;
        private Worksheet BalanceSheet;

        private string ExcelFilePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "balance.xlsx");
        private Process ExternalExcelProcess;
        public MainWindow()
        {
            InitializeComponent();
            InitExcelFile();
        }

        private void InitExcelFile()
        {
            //CloseExistingExcelProcesses();
            Excel = new Excel.Application();
            Excel.DisplayAlerts = false;
            Excel.ScreenUpdating = false;

            WorkBook = Excel.Workbooks.Open(ExcelFilePath);

            OperationsSheet = GetOrCreateWorksheet("Operations");
            InvoicesSheet = GetOrCreateWorksheet("Invoices");
            BalanceSheet = GetOrCreateWorksheet("Balance");
            OSVSheet = GetOrCreateWorksheet("OSV");

            WorkBook.Save();
            ExternalExcelProcess = Process.Start(ExcelFilePath);
        }


        private Worksheet GetOrCreateWorksheet(string wsName)
        {
            Worksheet ws = WorkBook.Worksheets.Cast<Worksheet>().FirstOrDefault(w => w.Name == wsName);
            if (ws == null)
            {
                int existedWsCount = WorkBook.Worksheets.Count;
                ws = WorkBook.Sheets.Add(WorkBook.Sheets[existedWsCount], Type.Missing, Type.Missing, Type.Missing);
                ws.Name = wsName;
            }
            return ws;
        }

        private void CloseExistingExcelProcesses()
        {
            Excel.Application excel = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Workbooks wbs = excel.Workbooks;
            foreach (Workbook wb in wbs)
            {
                Console.WriteLine(wb.Name); // print the name of excel files that are open
                wb.Save();
                wb.Close();
            }
            excel.Quit();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            Range range = OperationsSheet.UsedRange;
            List<Operation> operations = new List<Operation>();
            for (int i = 2; i < range.Rows.Count; i++)
            {
                Range row = OperationsSheet.UsedRange.Rows[i];
                var operation = new Operation();
                operation.Number = Convert.ToInt32(row.Cells[1, 1].Value2);
                operation.Date = DateTime.ParseExact(row.Cells[1, 2].Value.ToString(), "dd.MM.yyyy h:mm:ss", CultureInfo.InvariantCulture);
                operation.Documents = row.Cells[1, 3].Value2.ToString().Split('\n');
                operation.Subjects = ((string[])row.Cells[1, 4].Value2.ToString().Split('\n')).Skip(1).ToArray();
                operation.Debets = (string[])row.Cells[1, 5].Value2.ToString().Split('\n');
                operation.Credits = (string[])row.Cells[1, 6].Value2.ToString().Split('\n');
                string[] sumsStringArr = (string[])row.Cells[1, 7].Value2.ToString().Split('\n');
                operation.Sums = Array.ConvertAll(sumsStringArr, x => {
                    double y;
                    double.TryParse(x.Replace(" ", ""), out y);
                    return y;
                });
                operations.Add(operation);
            }
            SaveOperationInfoToJsonFile(operations);
        }

        private void SaveOperationInfoToJsonFile(List<Operation> operations)
        {
            List<OperationInvoiceRecord> infoRecords = new List<OperationInvoiceRecord>();
            foreach (var op in operations)
            {
                for (var i = 0; i < op.Subjects.Length; i++)
                {
                    try
                    {
                        var amount = op.Sums[i];

                        var debet = op.Debets[i];
                        var debetShortName = debet.Substring(0, 2);

                        var credit = op.Credits[i];
                        var creditShortName = credit.Substring(0, 2);

                        infoRecords.Add(new OperationInvoiceRecord()
                        {
                            FullName = debet,
                            ShortName = debetShortName,
                            Date = op.Date,
                            Amount = amount,
                            OperationInfoType = OperationInfoType.Debet,
                        });
                        infoRecords.Add(new OperationInvoiceRecord()
                        {
                            FullName = credit,
                            ShortName = creditShortName,
                            Date = op.Date,
                            Amount = amount,
                            OperationInfoType = OperationInfoType.Credit,
                        });
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }

                }
            }


            var jsonText = JsonConvert.SerializeObject(infoRecords, Formatting.Indented);
            File.WriteAllText("operations_info.json", jsonText);
            logListBox.Items.Add($"Операции сохранены в хранилище");
        }

        private List<T> GetStorageInfo<T>(string fileName)
        {
            try
            {
                string jsonText = File.ReadAllText(fileName);
                return JsonConvert.DeserializeObject<List<T>>(jsonText);
            }
            catch (Exception ex)
            {
                logListBox.Items.Add($"В хранилище нет файла {fileName}");
                return null;
            }            
        }

        private void CreateInvoices_Button_Click(object sender, RoutedEventArgs e)
        {

            List<OperationInvoiceRecord> infos = GetStorageInfo<OperationInvoiceRecord>("operations_info.json");
            if (infos == null)
            {
                return;
            }

            InvoicesSheet.Select(Type.Missing);
            ApplyGeneralStyle();

            int currentRowIndex = 1;

            var distinctInvoiceNamesGroups = infos.GroupBy(g => g.FullName);
            List<Invoice> invoices = new List<Invoice>();
            foreach (var group in distinctInvoiceNamesGroups)
            {

                var invoiceNameRow = InvoicesSheet.Rows.Range[
                    InvoicesSheet.Cells[currentRowIndex, 1],
                    InvoicesSheet.Cells[currentRowIndex, 2]] as Range;

                invoiceNameRow.Merge(); // merge 1st and 2nd cells in a row
                ApplyRowStyle(currentRowIndex, 2, XlRgbColor.rgbAquamarine, true);
                invoiceNameRow.Value = group.Key;

                currentRowIndex++;

                InvoicesSheet.Rows.Cells[currentRowIndex, 1].Value = "Дебет";
                InvoicesSheet.Rows.Cells[currentRowIndex, 2].Value = "Кредит";
                ApplyRowStyle(currentRowIndex, 2, XlRgbColor.rgbFloralWhite, false);
                
                currentRowIndex++;
                var debetAmounts = group.Where(x => x.OperationInfoType == OperationInfoType.Debet)
                    .Select(x => x.Amount).ToList();

                var creditAmounts = group.Where(x => x.OperationInfoType == OperationInfoType.Credit)
                    .Select(x => x.Amount).ToList();


                foreach (var cr in debetAmounts)
                {
                    InvoicesSheet.Rows.Cells[currentRowIndex, 1].Value = cr;
                    ApplyRowStyle(currentRowIndex, 2, XlRgbColor.rgbFloralWhite, false);
                    currentRowIndex++;
                }
                foreach (var db in creditAmounts)
                {
                    InvoicesSheet.Rows.Cells[currentRowIndex, 2].Value = db;
                    ApplyRowStyle(currentRowIndex, 2, XlRgbColor.rgbFloralWhite, false);
                    currentRowIndex++;
                }

                var invoiceShortName = group.First().ShortName;
                var debetInvoiceType = InvoiceTypeManager.GetInvoiceType(invoiceShortName);
                var sumDebet = debetAmounts.Sum();
                var sumCredit = creditAmounts.Sum();

                var invoice = new Invoice();
                invoice.ShortName = invoiceShortName;
                invoice.Description = InvoiceTypeManager.GetInvoiceDescription(invoiceShortName);
                invoice.DebetSum = sumDebet;
                invoice.CreditSum = sumCredit;
                invoice.Type = debetInvoiceType;


                InvoicesSheet.Rows.Cells[currentRowIndex, 1].Value = sumDebet;
                InvoicesSheet.Rows.Cells[currentRowIndex, 2].Value = sumCredit;
                InvoicesSheet.Rows.Cells[currentRowIndex, 3].Value = "Оборот";
                ApplyRowStyle(currentRowIndex, 3, XlRgbColor.rgbCornsilk, true);
                currentRowIndex++;

                double saldoDebet = 0, saldoCredit = 0;
                int saldoColumn;
                switch (debetInvoiceType)
                {
                    case InvoiceType.Active:
                        saldoDebet = sumDebet - sumCredit;
                        saldoCredit = 0;
                        saldoColumn = 1;
                        break;
                    case InvoiceType.Passive:
                        saldoDebet = 0;
                        saldoCredit = sumCredit - sumDebet;
                        saldoColumn = 2;
                        break;
                    case InvoiceType.ActivePassive:
                        if (sumDebet > sumCredit)
                        {
                            saldoColumn = 1;
                            saldoDebet = sumDebet - sumCredit;
                            saldoCredit = 0;
                        }
                        else
                        {
                            saldoColumn = 2;
                            saldoCredit = sumCredit - sumDebet;
                            saldoDebet = 0;
                        }
                        break;
                    default:
                        saldoColumn = 0;
                        break;
                }
                invoice.SaldoEndCredit = saldoCredit;
                invoice.SaldoEndDebet = saldoDebet;
                var saldo = Math.Max(saldoDebet, saldoCredit);
                InvoicesSheet.Rows.Cells[currentRowIndex, saldoColumn].Value = saldo;
                InvoicesSheet.Rows.Cells[currentRowIndex, 3].Value = "Сальдо";

                ApplyRowStyle(currentRowIndex, 3, XlRgbColor.rgbChartreuse, true);

                invoices.Add(invoice);
                currentRowIndex += 2;

            }

            var jsonText = JsonConvert.SerializeObject(invoices, Formatting.Indented);
            File.WriteAllText("invoices_info.json", jsonText);
            logListBox.Items.Add($"Счета записаны в хранилище");

            WorkBook.Save();            
            logListBox.Items.Add($"Счета созданы");
        }

        private void ApplyGeneralStyle()
        {
            InvoicesSheet.Columns[1].ColumnWidth = 20;
            InvoicesSheet.Columns[2].ColumnWidth = 20;
            InvoicesSheet.Columns[1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            InvoicesSheet.Columns[2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            InvoicesSheet.Columns[3].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            InvoicesSheet.UsedRange.Font.Color = XlRgbColor.rgbBlack;
        }

        private void ApplyRowStyle(int rowIndex, int columnsCount, XlRgbColor bgColor, bool needBold)
        {
            Range range = InvoicesSheet.Rows.Range[InvoicesSheet.Rows.Cells[rowIndex, 1],
                                                 InvoicesSheet.Rows.Cells[rowIndex, columnsCount]];
            range.Font.Bold = needBold;
            range.Interior.Color = bgColor;
            range.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            KillLocalExcelProcess();
        }

        private void KillLocalExcelProcess()
        {
            // Garbage collecting
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // Clean up references to all COM objects
            // As per above, you're just using a Workbook and Excel Application instance, so release them:  
            //Get PID
            int pid = -1;
            HandleRef hwnd = new HandleRef(Excel, (IntPtr)Excel.Hwnd);
            GetWindowThreadProcessId(hwnd, out pid);
            Excel.Quit();
            Marshal.FinalReleaseComObject(WorkBook);
            Marshal.FinalReleaseComObject(Excel);

            Process[] AllProcesses = Process.GetProcessesByName("EXCEL");
            foreach (Process process in AllProcesses)
            {
                if (process.Id == pid)
                {
                    process.Kill();
                }
            }
            AllProcesses = null;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowThreadProcessId(HandleRef handle, out int processId);
        private void ReloadExcel_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                KillLocalExcelProcess();
                ExternalExcelProcess.Kill();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            InitExcelFile();
        }

        private void CreateOSV_Button_Click(object sender, RoutedEventArgs e)
        {
            
            List<Invoice> invoices = GetStorageInfo<Invoice>("invoices_info.json");
            if (invoices == null)
            {
                return;
            }
            CreateOSVHeader();
            var currentRowIndex = 2;
            var groups = invoices.GroupBy(g => g.ShortName);
            foreach (var group in groups)
            {
                var invoisFirst = group.First();
                OSVSheet.Cells[currentRowIndex, 1].Value = group.Key;
                OSVSheet.Cells[currentRowIndex, 2].Value = invoisFirst.Description;
                OSVSheet.Cells[currentRowIndex, 3].Value = 0;
                OSVSheet.Cells[currentRowIndex, 4].Value = 0;
                OSVSheet.Cells[currentRowIndex, 5].Value = group.Sum(g => g.DebetSum);
                OSVSheet.Cells[currentRowIndex, 6].Value = group.Sum(g => g.CreditSum);
                OSVSheet.Cells[currentRowIndex, 7].Value = group.Sum(g => g.SaldoEndDebet);
                OSVSheet.Cells[currentRowIndex, 8].Value = group.Sum(g => g.SaldoEndCredit);
                currentRowIndex++;
            }
            currentRowIndex++;
            OSVSheet.Cells[currentRowIndex, 1].Value = "Итого";
            OSVSheet.Cells[currentRowIndex, 5].Value = invoices.Sum(x => x.DebetSum);
            OSVSheet.Cells[currentRowIndex, 6].Value = invoices.Sum(x => x.CreditSum);
            OSVSheet.Cells[currentRowIndex, 7].Value = invoices.Sum(x => x.SaldoEndDebet);
            OSVSheet.Cells[currentRowIndex, 8].Value = invoices.Sum(x => x.SaldoEndCredit);

            ApplyOSVUsedRangeStyle();
            WorkBook.Save();
            logListBox.Items.Add($"ОСВ создано");
        }
        private void CreateOSVHeader()
        {
            OSVSheet.Select(Type.Missing);
            ApplyOSVSheetStyle();
            OSVSheet.Cells[1, 1].Value = "Номер счета";
            OSVSheet.Cells[1, 2].Value = "Расшифровка";
            OSVSheet.Cells[1, 3].Value = "Сальдно начальное дебетовое";
            OSVSheet.Cells[1, 4].Value = "Сальдно начальное кредитовое";
            OSVSheet.Cells[1, 5].Value = "Оборот по дебету";
            OSVSheet.Cells[1, 6].Value = "Оборот по кредиту";
            OSVSheet.Cells[1, 7].Value = "Сальдно конечное дебетовое";
            OSVSheet.Cells[1, 8].Value = "Сальдно конечное кредитовое";
        }
        private void ApplyOSVSheetStyle()
        {
            var header = OSVSheet.Rows.Range[
                    OSVSheet.Cells[1, 1],
                    OSVSheet.Cells[1, 8]] as Range;

            header.RowHeight = 50;
            header.WrapText = true;
            header.Font.Bold = true;
            header.Interior.Color = XlRgbColor.rgbBeige;
            header.Borders.LineStyle = XlLineStyle.xlContinuous;
            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        }

        private void ApplyOSVUsedRangeStyle()
        {
            var range = OSVSheet.UsedRange;
            range.Columns.AutoFit();
            range.Borders.LineStyle = XlLineStyle.xlContinuous;
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        private void ApplyBalanceUsedRangeStyle()
        {
            var range = BalanceSheet.UsedRange;
            range.Columns.AutoFit();
            range.Borders.LineStyle = XlLineStyle.xlContinuous;
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        private void CreateBalance_Button_Click(object sender, RoutedEventArgs e)
        {
            
            List<Invoice> invoices = GetStorageInfo<Invoice>("invoices_info.json");
            if (invoices == null)
            {
                return;
            }
            CreateBalanceHeader();
            var groups = invoices.GroupBy(g => g.ShortName);
            double apSaldoSum = 0;            
            int activeIndex = 2, passiveIndex = 2, apIndex = 2;
            foreach (var group in groups)
            {
                var firstInvoice = group.First();
                switch (firstInvoice.Type)
                {
                    case InvoiceType.Active:
                        BalanceSheet.Cells[activeIndex, 1].Value = firstInvoice.ShortName;
                        BalanceSheet.Cells[activeIndex, 2].Value = firstInvoice.Description;
                        BalanceSheet.Cells[activeIndex, 3].Value = group.Sum(x => x.SaldoEndDebet);
                        activeIndex++;
                        break;
                    case InvoiceType.Passive:
                        BalanceSheet.Cells[passiveIndex, 6].Value = firstInvoice.ShortName;
                        BalanceSheet.Cells[passiveIndex, 7].Value = firstInvoice.Description;
                        BalanceSheet.Cells[passiveIndex, 8].Value = group.Sum(x => x.SaldoEndCredit);
                        passiveIndex++;
                        break;
                    case InvoiceType.ActivePassive:
                        BalanceSheet.Cells[apIndex, 11].Value = firstInvoice.ShortName;
                        BalanceSheet.Cells[apIndex, 12].Value = firstInvoice.Description;
                        var saldo = Math.Max(group.Sum(x => x.SaldoEndDebet), group.Sum(x => x.SaldoEndCredit));
                        apSaldoSum += saldo;
                        BalanceSheet.Cells[apIndex, 13].Value = saldo;
                        apIndex++;
                        break;
                    default:
                        break;
                }                
            }


            var indexMax = new[] { activeIndex, passiveIndex, apIndex }.Max();
            indexMax++;

            BalanceSheet.Cells[indexMax, 1].Value = "Итого";
            BalanceSheet.Cells[indexMax, 3].Value = invoices.Sum(x => x.SaldoEndDebet);
            BalanceSheet.Cells[indexMax, 8].Value = invoices.Sum(x => x.SaldoEndCredit);
            BalanceSheet.Cells[indexMax, 13].Value = apSaldoSum;

            ApplyBalanceUsedRangeStyle();
            WorkBook.Save();
            logListBox.Items.Add($"Баланс создан");
        }

        private void CreateBalanceHeader()
        {
            BalanceSheet.Select(Type.Missing);
            ApplyBalanceSheetStyle();
            BalanceSheet.Cells[1, 1].Value = "Актив";
            BalanceSheet.Cells[1, 2].Value = "Статья";
            BalanceSheet.Cells[1, 3].Value = "Сумма";

            BalanceSheet.Cells[1, 6].Value = "Пассив";
            BalanceSheet.Cells[1, 7].Value = "Статья";
            BalanceSheet.Cells[1, 8].Value = "Сумма";

            BalanceSheet.Cells[1, 11].Value = "Актив-Пассив";
            BalanceSheet.Cells[1, 12].Value = "Статья";
            BalanceSheet.Cells[1, 13].Value = "Сумма";
        }

        private void ApplyBalanceSheetStyle()
        {                        
            var header = BalanceSheet.Rows.Range[
                    BalanceSheet.Cells[1, 1],
                    BalanceSheet.Cells[1, 13]] as Range;

            header.RowHeight = 35;
            header.WrapText = true;
            header.Font.Bold = true;
            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            header.Interior.Color = XlRgbColor.rgbBeige;
            header.Borders.LineStyle = XlLineStyle.xlContinuous;

        }
    }
}
