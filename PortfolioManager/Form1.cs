using AMTD_Test_API;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using ZedGraph;
using static PortfolioManager.Form1;
using static AMTD_Test_API.AmeritradeBrokerAPI;
using System;
using System.Runtime.InteropServices;

namespace PortfolioManager
{
    public partial class Form1 : Form, Handler
    {
        public AmeritradeBrokerAPI oBroker;
        private Excel.Workbook _workbook;
        private Excel.Application _xlApp;
        private Excel._Worksheet _mainWorksheet;
        private Excel._Worksheet _pnlSheet;

        private string BrokerUserName;
        private string BrokerPassword;
        private string SourceID;
        private string BrokerUserNameKey = "BrokerUserName";
        private string BrokerPasswordKey = "BrokerPassword";
        private string SourceIDKey = "SourceID";
        private string FilePathKey = "FilePath";
        string path;

        private Dictionary<string, List<AmeritradeBrokerAPI.Option>> options;
        private List<AmeritradeBrokerAPI.L1quotes> quotes;
        private List<CashBalances> oCashBalances;
        private List<Positions> oPositions;

        public Form1()
        {
            oBroker = new AmeritradeBrokerAPI(this);
            InitializeComponent();
            info("Starting application...");
            usrName.Text = Settings.GetProtected(BrokerUserNameKey);
            password.Text = Settings.GetProtected(BrokerPasswordKey);
            sourceIdTextBox.Text = Settings.GetProtected(SourceIDKey);
            path = Settings.GetProtected(FilePathKey);
            fileLabel.Text = path;

        }

        void ThisWorkbook_BeforeClose(ref bool Cancel)
        {
            this.Invoke((MethodInvoker)delegate
            {
                cancelButton.PerformClick();
            });
        }

        public void info(string v)
        {
            error(v, null);
        }

        public void error(string text, Exception ex)
        {
            string timestamp = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss - ");
            string threadId = "Thread " + Thread.CurrentThread.ManagedThreadId + ": ";

            if (InvokeRequired)
            {
                text = timestamp + threadId + text;
                this.Invoke(new Action<string, Exception>(error), new object[] { text, ex });
                return;
            }
            else if (!text.Contains("Thread"))
            {
                text = timestamp + threadId + text;
            }

            if (ex != null)
            {
                text += "\r\n" + ex.Message;
                text += "\r\n" + ex.StackTrace;
            }

            output.AppendText(text + "\r\n");
        }

        private void login_Click(object sender, EventArgs e)
        {
            login();
        }

        private void login()
        {
            BrokerUserName = usrName.Text;
            BrokerPassword = password.Text;
            SourceID = sourceIdTextBox.Text;


            Settings.SetProtected(BrokerPasswordKey, BrokerPassword);
            Settings.SetProtected(BrokerUserNameKey, BrokerUserName);
            Settings.SetProtected(SourceIDKey, SourceID);
            Settings.SetProtected(FilePathKey, path);


            if (GetWorkBook() == null)
            {
                error("Unable to open workbook. Please select a correct Excel template file.", null);
                return;
            }


            if (SourceID.Length > 0)
            {
                if (BrokerUserName.Length > 0 && BrokerPassword.Length > 0)
                {
                    if (oBroker.TD_loginStatus == false)
                    {
                        info("Logging in...");
                        if (oBroker.TD_brokerLogin(BrokerUserName, BrokerPassword, SourceID, "1"))
                        {
                            oBroker.TD_GetStreamerInfo(BrokerUserName, BrokerPassword, SourceID, "1");
                            oBroker.TD_KeepAlive(BrokerUserName, BrokerPassword, SourceID, "1");

                            info("Opening Excel...");
                            GetMainWorkSheet().Select();
                            GetExcel().Visible = true;

                            Refresh();
                        }
                        else
                        {
                            info("Login FAILED.");
                        }
                    }
                    else
                    {
                        info("Already logged in.");
                    }
                }
                else
                {

                    info("Please enter your broker username and password, then try again.");
                }

            }
            else
            {
                info("Please enter your broker provided Source ID, then try again.");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Refresh();
        }

        public override void Refresh()
        {
            if (oBroker.TD_loginStatus == true)
            {

                try
                {
                    RefreshStocks(SourceID, BrokerUserName, BrokerPassword);
                }
                catch (Exception ex)
                {
                    error("Unable to refresh stock quotes", ex);
                }

                try
                {
                    RefreshOptions(SourceID, BrokerUserName, BrokerPassword);

                }
                catch (Exception ex)
                {
                    error("Unable to refresh options", ex);
                }

                try
                {
                    //RefreshPositions(SourceID, BrokerUserName, BrokerPassword);
                }
                catch (Exception ex)
                {
                    error("Unable to refresh positions", ex);
                }

            }
            else
            {
                info("Please login first.");
            }
        }


        private void RefreshPositions(string sourceID, string brokerUserName, string brokerPassword)
        {

            lock (excelLock)
            {
                info("Locked Excel.");
                info("Refreshing positions...");
                GetMainWorkSheet().Range["D15:D50"].Font.Color = ColorTranslator.ToOle(Color.Red);
                GetPnLWorksheetSheet().Range["G2:G13"].Font.Color = ColorTranslator.ToOle(Color.Red);

                GetPositions getPositions = new GetPositions(sourceID, brokerPassword, brokerUserName, oBroker, this);
                Thread oThread = new Thread(new ThreadStart(getPositions.ObtainPositions));
                oThread.Start();
            }
            info("Unlocked Excel.");

        }



        private void RefreshStocks(string SourceID, string BrokerUserName, string BrokerPassword)
        {
            lock (excelLock)
            {
                info("Locked Excel.");
                info("Refreshing stock quotes...");
                GetMainWorkSheet().Range["N2:P8"].Font.Color = ColorTranslator.ToOle(Color.Red);

                for (int row = 2; row < 9; row++)
                {
                    string symbol = (string)(GetMainWorkSheet().Cells[row, 3] as Excel.Range).Value;
                    if (!string.IsNullOrEmpty(symbol))
                    {
                        GetStockQuote getStockQuote = new GetStockQuote(symbol, SourceID, BrokerPassword, BrokerUserName, oBroker, this);
                        Thread oThread = new Thread(new ThreadStart(getStockQuote.GetQuotes));
                        oThread.Start();
                    }
                }

            }
            info("Unlocked Excel.");

        }




        private Excel.Application GetExcel()
        {
            if (_xlApp == null)
            {
                _xlApp = new Excel.Application();
            }
            return _xlApp;
        }

        private Excel.Workbook GetWorkBook()
        {
            if (_workbook == null)
            {                
                if (!string.IsNullOrEmpty(path))
                {
                    _workbook = GetExcel().Workbooks.Open(path);
                    _workbook.BeforeClose += ThisWorkbook_BeforeClose;
                }                     
            }
            return _workbook;
        }

        private Excel._Worksheet GetMainWorkSheet()
        {
            if (_mainWorksheet == null)
            {
                _mainWorksheet = (Excel._Worksheet)GetWorkBook().Sheets["Main"];
            }
            return _mainWorksheet;
        }

        private Excel._Worksheet GetPnLWorksheetSheet()
        {
            if (_pnlSheet == null)
            {
                _pnlSheet = (Excel._Worksheet)GetWorkBook().Sheets["PNL"];
            }
            return _pnlSheet;
        }


        private void RefreshOptions(string SourceID, string BrokerUserName, string BrokerPassword)
        {

            lock (excelLock)
            {
                info("Locked Excel.");
                info("Refreshing option quotes...");
                GetMainWorkSheet().Range["C2:C8"].Font.Color = ColorTranslator.ToOle(Color.Red);

                for (int row = 2; row < 9; row++)
                {
                    string symbol = (string)(GetMainWorkSheet().Cells[row, 3] as Excel.Range).Value;
                    if (!string.IsNullOrEmpty(symbol))
                    {
                        GetoptionChain getoptionChain = new GetoptionChain(symbol, SourceID, BrokerPassword, BrokerUserName,
                            oBroker, this);
                        Thread oThread = new Thread(new ThreadStart(getoptionChain.GetOptionChain));
                        oThread.Start();
                    }
                }
            }
            info("Unlocked Excel.");
        }

        internal void processEvent(DateTime time, AmeritradeBrokerAPI.ATradeArgument args)
        {
            throw new NotImplementedException();
        }

        private void Cleanup()
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            if (_mainWorksheet != null)
            {
                Marshal.ReleaseComObject(_mainWorksheet);
            }


            //close and release
            //  xlWorkbook.Save();
            if (_workbook != null)
            {
                Marshal.ReleaseComObject(_workbook);
            }


            //quit and release
            if (_xlApp != null)
            {
                _xlApp.Quit();
                Marshal.ReleaseComObject(_xlApp);
            }
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Cleanup();
        }

        private object excelLock = new object();


        void Handler.HandleOptionChain(List<AmeritradeBrokerAPI.Option> options)
        {
            if (options.Count > 0)
            {
                lock (excelLock)
                {
                    info("Locked Excel.");
                    try
                    {
                        string symbol = options[0].UnderlyingSymbol;
                        Excel._Worksheet symbolSheet = (Excel._Worksheet)GetWorkBook().Sheets[symbol];
                        Excel.Range xlRange = symbolSheet.UsedRange;
                        xlRange.ClearContents();
                        object[,] data = new object[options.Count, 13];
                        int row = 0;
                        foreach (AmeritradeBrokerAPI.Option option in options)
                        {
                            data[row, 0] = option.OptionSymbol;
                            data[row, 1] = option.GetType().Name;
                            data[row, 3] = DateTime.ParseExact(option.ExpirationDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            data[row, 5] = option.Strike;
                            data[row, 9] = option.Bid;
                            data[row, 10] = option.Ask;
                            data[row, 11] = option.ExpirationType;
                            data[row, 12] = option.Delta;
                            row++;
                        }
                        xlRange = GetExcel().Range[symbolSheet.Cells[1, 1], symbolSheet.Cells[data.GetLength(0), data.GetLength(1)]];
                        xlRange.Value = data;
                        (GetMainWorkSheet().Cells[25, "C"] as Excel.Range).Value = DateTime.Today;
                        (GetMainWorkSheet().Cells[GetSymbolRow(symbol), "C"] as Excel.Range).Font.Color = ColorTranslator.ToOle(Color.Black);
                    }
                    catch (Exception ex)
                    {
                        error("Unable to hendle option chain", ex);
                    }

                }
                info("Unlocked Excel.");

            }
        }

        private int GetSymbolRow(string symbol)
        {
            for (int row = 2; row < 9; row++)
            {
                var range = (GetMainWorkSheet().Cells[row, 3] as Excel.Range);
                string cell = (string)range.Value;
                if (!string.IsNullOrEmpty(cell) && cell.Equals(symbol))
                {
                    return row;
                }
            }
            return 0;
        }

        public void HandleStockQuote(AmeritradeBrokerAPI.L1quotes quote)
        {
            lock (excelLock)
            {
                info("Locked Excel.");
                try
                {
                    object[] data = new object[3];
                    data[0] = quote.last;
                    data[1] = quote.change;
                    data[2] = quote.close;
                    int row = GetSymbolRow(quote.stock);
                    Excel.Range xlRange = GetExcel().Range[GetMainWorkSheet().Cells[row, 14], GetMainWorkSheet().Cells[row, 16]];
                    xlRange.Value = data;
                    (GetMainWorkSheet().Cells[row, 14] as Excel.Range).Font.Color = ColorTranslator.ToOle(Color.Black);
                    (GetMainWorkSheet().Cells[row, 15] as Excel.Range).Font.Color = ColorTranslator.ToOle(Color.Black);
                    (GetMainWorkSheet().Cells[row, 16] as Excel.Range).Font.Color = ColorTranslator.ToOle(Color.Black);
                }
                catch (Exception ex)
                {
                    error("Unable to hendle stock quote", ex);
                }

            }
            info("Unlocked Excel.");
        }


        public void HandlePositions(List<CashBalances> oCashBalances, List<Positions> oPositions)
        {
            lock (excelLock)
            {
                info("Locked Excel.");
                try
                {
                    if (oCashBalances.Any())
                    {
                        GetMainWorkSheet().Cells[27, "D"] = oCashBalances[0].AvailableFundsForTrading;
                        GetMainWorkSheet().Cells[27, "D"].Font.Color = ColorTranslator.ToOle(Color.Black);
                        GetMainWorkSheet().Cells[28, "D"] = oCashBalances[0].ChangeInCashBalance;
                        GetMainWorkSheet().Cells[28, "D"].Font.Color = ColorTranslator.ToOle(Color.Black);
                        GetMainWorkSheet().Cells[29, "D"] = oCashBalances[0].CurrentCashBalance;
                        GetMainWorkSheet().Cells[29, "D"].Font.Color = ColorTranslator.ToOle(Color.Black);
                    }

                    if (oPositions.Any())
                    {
                        for (int row = 2; row < 50; row++)
                        {
                            string optoinSymbol = (GetPnLWorksheetSheet().Cells[row, 1] as Excel.Range).Value;
                            if (!string.IsNullOrEmpty(optoinSymbol) && !optoinSymbol.Equals("TOTAL"))
                            {
                                Positions position = oPositions.Where(p => p.StockSymbol == optoinSymbol).FirstOrDefault();
                                GetPnLWorksheetSheet().Cells[row, 3] = position.AveragePric;
                                GetPnLWorksheetSheet().Cells[row, 7] = position.ClosePrice;

                                GetPnLWorksheetSheet().Cells[row, 7].Font.Color = ColorTranslator.ToOle(Color.Black);
                                GetMainWorkSheet().Range["D15:D21"].Font.Color = ColorTranslator.ToOle(Color.Black);

                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    error("Unable to hendle balances/positions", ex);
                }

            }
            info("Unlocked Excel.");
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                path = file.FileName;
                fileLabel.Text = path;
            }
        }

        private void fileLabel_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
           

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(usrName.Text) && !string.IsNullOrEmpty(password.Text) && !string.IsNullOrEmpty(sourceIdTextBox.Text) && !string.IsNullOrEmpty(path))
            {
                 login();
            }

        }
    }


    public class GetPositions
    {

        private string source;
        private string password;
        private string username;
        public AmeritradeBrokerAPI oBroker;
        public Handler handler;

        public GetPositions(string source, string password, string username, AmeritradeBrokerAPI oBroker,
           Handler handler)
        {
            this.source = source;
            this.password = password;
            this.username = username;
            this.oBroker = oBroker;
            this.handler = handler;
        }


        public void ObtainPositions()
        {
            handler.info("In GetPositions.ObtainPositions: checkin login status...");
            if (oBroker.TD_loginStatus)
            {
                List<CashBalances> oCashBalances = new List<CashBalances>();
                List<Positions> oPositions = new List<Positions>();
                handler.info("In GetPositions.ObtainPositions: calling TD_getAcctBalancesAndPositions...");
                oBroker.TD_getAcctBalancesAndPositions(username, password, source, "1", ref oCashBalances, ref oPositions);
                handler.HandlePositions(oCashBalances, oPositions);
            }

        }
    }

    public class GetStockQuote
    {
        private string symbol;
        private string source;
        private string password;
        private string username;
        public AmeritradeBrokerAPI oBroker;
        private Handler handler;

        public GetStockQuote(string symbol, string source, string password, string username, AmeritradeBrokerAPI oBroker,
           Handler handler)
        {
            this.symbol = symbol;
            this.source = source;
            this.password = password;
            this.username = username;
            this.oBroker = oBroker;
            this.handler = handler;
        }


        public void GetQuotes()
        {
            if (oBroker.TD_loginStatus)
            {
                L1quotes quote = oBroker.TD_getsnapShot(symbol, source, username, password);
                handler.HandleStockQuote(quote);
            }
        }
    }

    public interface Handler
    {
        void HandleOptionChain(List<AmeritradeBrokerAPI.Option> options);
        void HandleStockQuote(AmeritradeBrokerAPI.L1quotes quote);
        void HandlePositions(List<CashBalances> oCashBalances, List<Positions> oPositions);

        void info(string v);
        void error(string text, Exception ex);
    }


    public class GetoptionChain
    {
        private string symbol;
        private string source;
        private string password;
        private string username;
        public AmeritradeBrokerAPI oBroker;
        public Handler handler;

        public GetoptionChain(string symbol, string source, string password, string username, AmeritradeBrokerAPI oBroker,
           Handler handler)
        {
            this.symbol = symbol;
            this.source = source;
            this.password = password;
            this.username = username;
            this.oBroker = oBroker;
            this.handler = handler;
        }


        public void GetOptionChain()
        {
            if (oBroker.TD_loginStatus)
            {
                List<AmeritradeBrokerAPI.Option> options = new List<AmeritradeBrokerAPI.Option>();
                oBroker.TD_getOptionChain(symbol, username, password, source, "1", ref options, true);
                handler.HandleOptionChain(options);
            }

        }
    }
}

