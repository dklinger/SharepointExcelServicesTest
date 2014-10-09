using System;
using System.IO;
using System.Windows;
using DummyWpfApp.SP2ExcelService;

namespace DummyWpfApp
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnDoSomething_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                #region old

                //var xlService = new SPExcelService.ExcelServiceSoapClient();
                //SPExcelService.Status[] outStatus;

                //var workbookPath = "http://vsgc-wss/personal/dklinger/Personal%20Documents/TestExcel.xlsx";

                //var sheetName = "Tabelle1";

                //#region DK-PW

                //xlService.ClientCredentials.Windows.ClientCredential = new NetworkCredential("username", "password");

                //#endregion DK-PW

                //var apiVer = xlService.GetApiVersion(out outStatus);

                //txtDebug.Text = apiVer;

                //var sessionId = xlService.OpenWorkbook(workbookPath, "de-DE", "de-DE", out outStatus);

                //xlService.SetRangeA1(sessionId, sheetName, "ZelleEins", new ArrayOfAnyType()
                //{
                //    5
                //});

                //xlService.CalculateWorkbook(sessionId, CalculateType.CalculateFull);

                //var zelleDrei = xlService.GetCell(sessionId, sheetName, 1, 3, true, out outStatus);

                //MessageBox.Show(zelleDrei.ToString());

                #endregion old

                //Service Definition
                var xlService = new SP2ExcelService.ExcelService();
                SP2ExcelService.Status[] outStatus;

                xlService.Credentials = System.Net.CredentialCache.DefaultCredentials;

                //Workbook Definition
                var workbookPath = "http://vsgc-wss/technics/Satelliten%20Team/Mappe1.xlsx";

                var sheetName = "Tabelle1";

                //Open Workbook
                var sessionId = xlService.OpenWorkbook(workbookPath, "de-DE", "de-DE", out outStatus);

                //Write to Cell
                xlService.SetCellA1(sessionId, sheetName, "ZelleEins", 5);

                var zelleDrei1 = xlService.GetCell(sessionId, sheetName, 0, 2, true, out outStatus);

                MessageBox.Show(zelleDrei1.ToString());
                //Recalculate Workbook
                xlService.CalculateWorkbook(sessionId, SP2ExcelService.CalculateType.CalculateFull);

                //Read Result
                var zelleDrei = xlService.GetCell(sessionId, sheetName, 0, 2, true, out outStatus);

                MessageBox.Show(zelleDrei.ToString());

                var workbook = xlService.GetWorkbook(sessionId, WorkbookType.FullSnapshot, out outStatus);

                outStatus = xlService.CloseWorkbook(sessionId);

                var writeStream = new FileStream("c:\\temp\\testxls.xlsx", FileMode.Create);

                var binaryWriter = new BinaryWriter(writeStream);
                binaryWriter.Write(workbook);
                binaryWriter.Close();
            }
            catch(Exception ex)
            {
                txtDebug.Text += "\r\n\r\n" + ex.Message + "\r\n\r\n" + ex.StackTrace;
            }
        }
    }
}