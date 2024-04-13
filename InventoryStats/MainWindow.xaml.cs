/* Title:           Calculate Run Rates for Parts
 * Date:            10-11-18
 * Author:          Terry Holmes
 * 
 * Description:     This is the window for calculating stats */

using System;
using System.Collections.Generic;
using System.Linq;
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
using DateSearchDLL;
using InventoryDLL;
using InventoryStatsDLL;
using NewEventLogDLL;
using NewEmployeeDLL;
using NewPartNumbersDLL;
using Microsoft.Win32;
using excel = Microsoft.Office.Interop.Excel;

namespace InventoryStats
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        InventoryClass TheInventoryClass = new InventoryClass();
        InventoryStatsClass TheInventoryStatsClass = new InventoryStatsClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        PartNumberClass ThePartNumberClass = new PartNumberClass();

        //setting up the data
        FindSpectrumInventoryIssueStatsDataSet TheFindSpectrumInventoryIssueStatsDataSet = new FindSpectrumInventoryIssueStatsDataSet();
        FindPartsWarehousesDataSet TheFindPartsWarehouseDataSet = new FindPartsWarehousesDataSet();
        FindPartByPartNumberDataSet TheFindPartByPartNumberDataSet = new FindPartByPartNumberDataSet();
        CalculatedInventoryStats TheCalculatedInventoryStats = new CalculatedInventoryStats();
        CalculatedInventoryStats TheFinalCalculatedInventoryStats = new CalculatedInventoryStats();
        ImportPartsDataSet TheImportPartsDataSet = new ImportPartsDataSet();
        FindWarehouseInventoryPartDataSet TheFindWarehouseInventoryPartDataSet = new FindWarehouseInventoryPartDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void mitCreateHelpDeskTicket_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.LaunchHelpDeskTickets();
        }

        private void mitHelpSite_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.LaunchHelpSite();
        }

        private void mitCloseApplication_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }

        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;

            try
            {
                TheFindPartsWarehouseDataSet = TheEmployeeClass.FindPartsWarehouses();
                cboSelectWarehouse.Items.Add("Select Warehouse");
                intNumberOfRecords = TheFindPartsWarehouseDataSet.FindPartsWarehouses.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    cboSelectWarehouse.Items.Add(TheFindPartsWarehouseDataSet.FindPartsWarehouses[intCounter].FirstName);
                }

                cboSelectWarehouse.SelectedIndex = 0;
            }
            catch(Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Inventory Stats // Main Window // Window Loaded " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void cboSelectWarehouse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intSelectedIndex;
            int intWarehouseID;
            DateTime datStartDate = DateTime.Now;
            DateTime datEndDate = DateTime.Now;
            DateTime datTodaysDate = DateTime.Now;
            int intCounter;
            int intNumberOfRecords;
            int intRecordsReturned;
            int intPartID;
            int intLoop;
            string strPartNumber;
            double douAverage;
            double douSTDev;
            int intCount;
            bool blnDoNotCopy;

            PleaseWait PleaseWait = new PleaseWait();
            PleaseWait.Show();

            try
            {
                intSelectedIndex = cboSelectWarehouse.SelectedIndex - 1;

                if (intSelectedIndex > -1)
                {
                    TheCalculatedInventoryStats.inventorystats.Rows.Clear();

                    intWarehouseID = TheFindPartsWarehouseDataSet.FindPartsWarehouses[intSelectedIndex].EmployeeID;

                    datTodaysDate = TheDateSearchClass.RemoveTime(datTodaysDate);

                    intNumberOfRecords = TheImportPartsDataSet.parts.Rows.Count - 1;


                    for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        datStartDate = TheDateSearchClass.SubtractingDays(datTodaysDate, 180);
                        datEndDate = TheDateSearchClass.AddingDays(datStartDate, 30);
                        intLoop = 1;
                        douAverage = 0;
                        intCount = 0;

                        intPartID = TheImportPartsDataSet.parts[intCounter].PartID;

                        if (intPartID != -1)
                        {
                            TheFindWarehouseInventoryPartDataSet = TheInventoryClass.FindWarehouseInventoryPart(intPartID, intWarehouseID);

                            TheImportPartsDataSet.parts[intCounter].OnHandQty = TheFindWarehouseInventoryPartDataSet.FindWarehouseInventoryPart[0].Quantity;

                            while (datStartDate < datTodaysDate)
                            {
                               TheFindSpectrumInventoryIssueStatsDataSet = TheInventoryStatsClass.FindSpectrumInventoryIssueStats(intPartID, intWarehouseID, datStartDate, datEndDate);

                                intRecordsReturned = TheFindSpectrumInventoryIssueStatsDataSet.FindSpectrumInventoryIssueStats.Rows.Count;

                                if (intRecordsReturned == 0)
                                {
                                    intCount += 0;
                                }
                                else
                                {
                                    intCount += TheFindSpectrumInventoryIssueStatsDataSet.FindSpectrumInventoryIssueStats[0].TotalIssued;
                                }


                                datStartDate = datEndDate;
                                datEndDate = TheDateSearchClass.AddingDays(datEndDate, 30);
                                intLoop++;
                            }

                            douAverage = (Convert.ToDouble(intCount)) / 6;

                            TheImportPartsDataSet.parts[intCounter].RunRate = Convert.ToDecimal(douAverage);
                        }

                    }

                    dgrResults.ItemsSource = TheImportPartsDataSet.parts;
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Inventory Stats // Main Window // Combo Box Changed " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

            PleaseWait.Close();
        }

        private void mitExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "OpenOrders";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;
                intRowNumberOfRecords = TheImportPartsDataSet.parts.Rows.Count;
                intColumnNumberOfRecords = TheImportPartsDataSet.parts.Columns.Count;

                for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                {
                    worksheet.Cells[cellRowIndex, cellColumnIndex] = TheImportPartsDataSet.parts.Columns[intColumnCounter].ColumnName;

                    cellColumnIndex++;
                }

                cellRowIndex++;
                cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                {
                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheImportPartsDataSet.parts.Rows[intRowCounter][intColumnCounter].ToString();

                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                saveDialog.ShowDialog();

                workbook.SaveAs(saveDialog.FileName);
                MessageBox.Show("Export Successful");

            }
            catch (System.Exception ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay ERP // Find Employee Productivity Footage // Export to Excel " + ex.Message);

                MessageBox.Show(ex.ToString());
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }

        private void mitImportExcel_Click(object sender, RoutedEventArgs e)
        {
            excel.Application xlDropOrder;
            excel.Workbook xlDropBook;
            excel.Worksheet xlDropSheet;
            excel.Range range;

            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            string strPartNumber;
            string strItemNumber;
            int intTransactionID;
            string strPartDescription;
            int intPartID;
            bool blnItemNotFound;
            int intRecordsReturned;

            try
            {
                TheImportPartsDataSet.parts.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 2; intCounter <= intNumberOfRecords; intCounter++)
                {

                    strPartNumber = Convert.ToString((range.Cells[intCounter, 6] as excel.Range).Value2);
                    intTransactionID = Convert.ToInt32((range.Cells[intCounter, 1] as excel.Range).Value2);
                    strItemNumber = Convert.ToString((range.Cells[intCounter, 2] as excel.Range).Value2);
                    strPartDescription = Convert.ToString((range.Cells[intCounter, 3] as excel.Range).Value2);

                    TheFindPartByPartNumberDataSet = ThePartNumberClass.FindPartByPartNumber(strPartNumber);

                    intRecordsReturned = TheFindPartByPartNumberDataSet.FindPartByPartNumber.Rows.Count;

                    if (intRecordsReturned > 0)
                    {
                        intPartID = TheFindPartByPartNumberDataSet.FindPartByPartNumber[0].PartID;
                        
                    }
                    else
                    {
                        intPartID = -1;
                    }

                    ImportPartsDataSet.partsRow NewPartRow = TheImportPartsDataSet.parts.NewpartsRow();

                    NewPartRow.ItemDescription = strPartDescription;
                    NewPartRow.ItemNumber = strItemNumber;
                    NewPartRow.OnHandQty = 0;
                    NewPartRow.PartID = intPartID;
                    NewPartRow.PartNumber = strPartNumber;
                    NewPartRow.RunRate = 0;
                    NewPartRow.TransactionID = intTransactionID;

                    TheImportPartsDataSet.parts.Rows.Add(NewPartRow);
                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheImportPartsDataSet.parts;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay ERP // Import Vehicles // Import Excel Menu Item " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
