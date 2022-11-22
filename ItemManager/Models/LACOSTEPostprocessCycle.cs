using System;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using DataSolutions.ApplicationFramework;
using DataSolutions.Logging.Logger;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;
using NPOI;
using NPOI.HSSF.Util;
using System.Xml;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;

namespace ItemManager.Models
{
    public class LACOSTEPostprocessCycle /*: DataSolutionsServiceBase*/
    {
        public const string DOCUMENT_CODE = "ItemTable";
        public const string FILE_EXTENSION_IN = "xlsx";
        public const string FILE_EXTENSION_OUT = "xlsx";  // Extract to the excel file

        private readonly string _inputFolder;

        private readonly string _outputFolder;
        private readonly string _workingFolder;

        private readonly string _failedFolder;
        private readonly string _failedReportedFolder;
        private readonly string _failedSentFolder;
        private readonly string _archiveFilePath;
        private readonly bool _archiveEnabled;
        private readonly string _connectionString;

        private readonly char _tab = ((char)09);    // Tab 


        private FileInfo _currentFile;

        public LACOSTEPostprocessCycle(IHostingEnvironment env)
            //: base(logger, Guid.NewGuid())
        {
            string BuyerShortCode = "LCO";
            string DocumentCode = DOCUMENT_CODE;


            _inputFolder = Path.Combine(env.ContentRootPath, "TEST", "ftp", "Upload");
            _workingFolder = Path.Combine(env.ContentRootPath, "TEST", "ftp", "Upload", "tmp");
            _outputFolder = Path.Combine(env.ContentRootPath, "TEST", "ftp", "inproc997");
            _failedFolder = Path.Combine(env.ContentRootPath, "TEST", "ftp", "xfailed");
            _failedReportedFolder = Path.Combine(env.ContentRootPath, "TEST", "ftp", "xfailed", "Reported");
            _failedSentFolder = Path.Combine(env.ContentRootPath, "TEST", "ftp", "xfailed", "reported","Sent");
            _archiveFilePath = Path.Combine(env.ContentRootPath, "TEST", "ftp");
            string defaultContextName = "portal20PS";
            _connectionString = "Server=localhost;Database=TLO20PSUAT;Trusted_Connection=True;";

        }

        public void DoWork()
        {
            var inputFiles = (new DirectoryInfo(_inputFolder)).GetFiles($@"*.{FILE_EXTENSION_IN}").ToList();

            foreach (var purchaseOrderFile in inputFiles)
            { // in each excel file

                var fileCorrelationId = Guid.NewGuid();
                try
                {
                    _currentFile = purchaseOrderFile;
                    _currentFile.MoveTo(Path.Combine(_workingFolder, _currentFile.Name));

                    int failureCount = 0;
                    string errorMessage = "";

                    FileStream file = new FileStream(_currentFile.FullName, FileMode.Open, FileAccess.Read);
                    List<string> ColumnNames = new List<string>();

                    using (var package = new ExcelPackage(file))
                    {
                        var currentSheet = package.Workbook.Worksheets;
                        var workSheet = currentSheet.First();
                        var noOfCol = workSheet.Dimension.End.Column;
                        var noOfRow = workSheet.Dimension.End.Row;
                        foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                        {
                            
                            for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
                            {
                                ColumnNames.Add(worksheet.Cells[1, i].Value.ToString()); // 1 = First Row, i = Column Number
                            }
                        }
                    }



                    //XSSFWorkbook workbook;
                    //using (FileStream file1 = new FileStream(_currentFile.FullName, FileMode.Open, FileAccess.Read))
                    //{
                    //    workbook = new XSSFWorkbook(file);
                    //}

                    //header
                    //string SupplierCol = workbook.GetSheetAt(0).GetRow(0).GetCell(0).ToString().Trim();
                    //string TARIFCol = workbook.GetSheetAt(0).GetRow(0).GetCell(1).ToString().Trim();
                    //string RefCol = workbook.GetSheetAt(0).GetRow(0).GetCell(2).ToString().Trim();
                    //string EANCol = workbook.GetSheetAt(0).GetRow(0).GetCell(3).ToString().Trim();
                    //string UPCCol = workbook.GetSheetAt(0).GetRow(0).GetCell(4).ToString().Trim();
                    //string UnitPriceCol = workbook.GetSheetAt(0).GetRow(0).GetCell(5).ToString().Trim();

                    string SupplierCol = ColumnNames.ElementAt(0);
                    string TARIFCol = ColumnNames.ElementAt(1);
                    string RefCol = ColumnNames.ElementAt(2);
                    string EANCol = ColumnNames.ElementAt(3);
                    string UPCCol = ColumnNames.ElementAt(4);
                    string UnitPriceCol = ColumnNames.ElementAt(5);


                    if (SupplierCol.ToUpper() != "SUPPLIER" || TARIFCol.ToUpper() != "TARIF" || RefCol.ToUpper() != "REF COL"
                        || EANCol.ToUpper() != "EAN" || UPCCol.ToUpper() != "UPC" || UnitPriceCol.ToUpper() != "PRICES USD")
                    {
                        //Logger.Error(string.Format("Excel file header columns label are not correct : SUPPLIER,TARIF,REF COL,EAN,UPC,PRICES USD "), null, BuyerShortCode, DocumentCode, fileCorrelationId, _currentFile.FullName, _currentFile.Length);
                        failureCount = failureCount + 1;
                    }
                    else
                    {

                        List<ItemMasterTableList> masterItemList = loadItemMaterTableList(_currentFile.FullName);
                        var emptyValueList = from itemList in masterItemList
                                             where (itemList.Supplier == "" || (itemList.EAN == "" && itemList.UPC == "")
                                             || (itemList.EAN.Trim() == "#N/A" && itemList.UPC.Trim() == "#N/A") || itemList.UnitPrice == "")
                                             select new { Supplier = itemList.Supplier, Tarif = itemList.Tarif, RefCol = itemList.RefCol, EAN = itemList.EAN, UPC = itemList.UPC, UnitPrice = itemList.UnitPrice };


                        if (emptyValueList.Count() > 0) 
                        {
                            string[] outPutArray = new string[emptyValueList.Count()];
                            int itemIndex = 0;
                            foreach (var masterItem in emptyValueList)
                            {
                                outPutArray[itemIndex] = masterItem.Supplier + _tab + masterItem.Tarif + _tab + masterItem.RefCol
                                    + _tab + masterItem.EAN + _tab + masterItem.UPC + _tab + masterItem.UnitPrice;
                                itemIndex++;
                            }


                            string fileName = Path.GetFileNameWithoutExtension(_currentFile.FullName) + "_failed_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                            string worksheetName = "Sheet1";
                            string excelFileName = _failedReportedFolder + @"\" + fileName;
                            ConvertToExcelWithNPOI(excelFileName, worksheetName, outPutArray);
                            string dateFolder = createDateFolder(_failedSentFolder);
                            string sentFile = Path.Combine(dateFolder, Path.GetFileName(excelFileName));
                            File.Move(excelFileName, sentFile);

                        }

                    }

                    if (failureCount > 0)
                    {
                        string failedFile = Path.Combine(_failedFolder, Path.GetFileName(_workingFolder));
                        _currentFile.MoveTo(Path.Combine(_failedFolder, _currentFile.Name));
                    }
                    else if (failureCount == 0)
                    {
                        List<ItemMasterTableList> masterItemList = loadItemMaterTableList(_currentFile.FullName);
                        var masterValueList = masterItemList;

                        int totalRecord = 0;
                        int countInsert = 0;
                        int countUpdate = 0;

                        foreach (var masterItem in masterValueList)
                        {
                            SqlConnection connection = new SqlConnection(_connectionString);
                            connection.Open();
                            SqlCommand recordExist = new SqlCommand("SELECT count(1) FROM tblItemMaster Where BUYERLONGCODE='LACOSTE' AND SUPPLIERLONGCODE='" + masterItem.Supplier + "'" + "AND EAN ='" + masterItem.EAN + "'" + "AND UPC ='" + masterItem.UPC + "'" + "AND HARMONIZEDTARIFFSCHEDULE ='" + masterItem.Tarif + "'", connection);

                            Int32 countRecordExist = Convert.ToInt32(recordExist.ExecuteScalar());  //check contains key
                            if (countRecordExist > 0)
                            {
                                //Update Record
                                if (!UpdateDataInMasterTable(connection, masterItem.Supplier, masterItem.RefCol, masterItem.EAN, masterItem.UPC, masterItem.UnitPrice, masterItem.Tarif))
                                {
                                    string error1 = "Update excel record to db failed...SeqID: {0}, EAN: {1}";
                                }
                                countUpdate++;
                            }
                            else
                            {
                                //Insert Record
                                if (!InsertDataIntoMasterTable(connection, masterItem.Supplier, masterItem.RefCol, masterItem.EAN, masterItem.UPC, masterItem.UnitPrice, masterItem.Tarif))
                                {
                                    string error2 = "Insert excel record to db failed...SeqID: {0}, EAN: {1}";
                                }
                                countInsert++;
                            }

                            connection.Close(); //Remember close the connection

                        }
                        //Logger.Info(string.Format("Insert/Changed excel record to db ...Inserted: {0}, Updated: {1}, Total Record: {2}", countInsert, countUpdate, masterValueList.Count()), BuyerShortCode, DocumentCode, fileCorrelationId, _currentFile.FullName, _currentFile.Length);
                        //Console.WriteLine("Inserted:" + countInsert + " Updated:" + countUpdate + " Total Record:" + masterValueList.Count());

                        //string dateFolder = createDateFolder(_archiveFilePath);
                        //string archiveFile = Path.Combine(dateFolder, Path.GetFileName(_currentFile.FullName));
                        //File.Move(_currentFile.FullName, archiveFile);
                        //Logger.Info(string.Format("The source file was moved to archive folder...{0}", archiveFile), BuyerShortCode, DocumentCode, fileCorrelationId, Path.GetFileName(archiveFile), (new FileInfo(archiveFile).Length));
                    }
                }
                catch (Exception ex)
                {
                    string exceptionMessage = "Error while loading item master excel file, find exception for more details";
                    //Logger.Error(exceptionMessage
                    //            , ex, BuyerShortCode, DocumentCode, fileCorrelationId, _currentFile.FullName, _currentFile.Length);
                    _currentFile.MoveTo(Path.Combine(_failedFolder, _currentFile.Name));
                }
            }

            //Logger.Info("Leaving LACOSTEPostprocessCycle.DoWork..."
            //            , BuyerShortCode, DocumentCode, CorrelationId);

        }

        private class ItemMasterTableList
        {
            public int SeqID { get; set; }
            public string Sender { get; set; }
            public string Supplier { get; set; }
            public string Tarif { get; set; }
            public string RefCol { get; set; }
            public string EAN { get; set; }
            public string UPC { get; set; }
            public string UnitPrice { get; set; }
        }

        private List<ItemMasterTableList> loadItemMaterTableList(string inputFile)
        {
            List<ItemMasterTableList> MasterItemList = new List<ItemMasterTableList>();
            try
            {
                XSSFWorkbook workbook;
                //using (FileStream file = new FileStream(@"C:\TEST\FTP\FXR\upload\850\ASN0000002.xlsx", FileMode.Open, FileAccess.Read))
                using (FileStream file = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(file);
                }

                var sheet = workbook.GetSheetAt(0); // first sheet
                int SeqID = 0;
                for (var i = 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null) continue;

                    string colSender = "LACOSTE";


                    // Calvin - Do checking here
                    //string colSupplier = row.GetCell(0) == null ? "NULL" : row.GetCell(0).ToString();
                    string colSupplier = row.GetCell(0).ToString();
                    //string colTarif = row.GetCell(1) == null ? "" : row.GetCell(1).ToString().Trim();
                    string colTarif = row.GetCell(1).ToString();
                    string colRefCol = row.GetCell(2) == null ? "" : row.GetCell(2).ToString();
                    string colEAN = row.GetCell(3) == null ? "" : row.GetCell(3).ToString().Trim();
                    string colUPC = row.GetCell(4) == null ? "" : row.GetCell(4).ToString().Trim();
                    string colUnitPrice = row.GetCell(5) == null ? "" : row.GetCell(5).ToString();

                    if (IsNullOrWhiteSpace(colSupplier) == true || IsNullOrWhiteSpace(colTarif) == true)
                    {
                        failedWithMissingColumn("Missing Supplier or Tarif", i, inputFile);
                    }

                    if (colSupplier != "NULL" && colTarif != "NULL")
                    {
                        MasterItemList.Add(new ItemMasterTableList()
                        {
                            SeqID = i + 1,
                            Sender = colSender,
                            Supplier = colSupplier,
                            Tarif = colTarif,
                            RefCol = colRefCol,
                            EAN = colEAN,
                            UPC = colUPC,
                            UnitPrice = colUnitPrice
                        });
                    }
                }
                return MasterItemList;
            }
            catch (Exception e1)
            {
                //Logger.Error(string.Format("Load Item Master Table: exceptional error occured...{0}", e1.Message), BuyerShortCode, DocumentCode, CorrelationId);
                return MasterItemList;
            }
        }

        private bool InsertDataIntoMasterTable(SqlConnection connection, string Supplier, string RefCol, string EAN, string UPC, string UnitPrice, string Tarif)
        {
            try
            {
                //using (var connection = new SqlConnection(_connectionString))
                using (var command = connection.CreateCommand())
                {
                    command.CommandType = CommandType.Text;
                    command.CommandText = @"insert into tblItemMaster (
                                            BUYERLONGCODE,SUPPLIERLONGCODE, STATUS, ORGANIZATION, LABEL, UPC, EAN, 
                                            COLORCODE, HARMONIZEDTARIFFSCHEDULE, UNITPRICE, LASTMODIFIEDAT, LASTMODIFIEDBY) values (
                                            @BuyerCode, @SupplierLongCode, @Status, @BuyerName, @Label, @UPC, @EAN, 
                                            @ColorCode, @Tarif, @unitPrice, @ModifiedDate, @ModifiedBy)";
                    command.Parameters.AddWithValue("@BuyerCode", "LACOSTE");
                    command.Parameters.AddWithValue("@SupplierLongCode", Supplier);
                    command.Parameters.AddWithValue("@Status", "1");
                    command.Parameters.AddWithValue("@BuyerName", "LACOSTE");
                    command.Parameters.AddWithValue("@Label", "PRICES FW 18");
                    command.Parameters.AddWithValue("@UPC", (UPC == "#N/A") ? Convert.DBNull : UPC);
                    command.Parameters.AddWithValue("@EAN", (EAN == "#N/A") ? Convert.DBNull : EAN);
                    command.Parameters.AddWithValue("@ColorCode", RefCol);
                    command.Parameters.AddWithValue("@Tarif", Tarif);
                    command.Parameters.AddWithValue("@unitPrice", UnitPrice);
                    command.Parameters.AddWithValue("@ModifiedDate", DateTime.Now);
                    command.Parameters.AddWithValue("@ModifiedBy", "TLO");

                    //connection.Open();
                    var result = command.ExecuteScalar();
                    //connection.Close();
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        private bool UpdateDataInMasterTable(SqlConnection connection, string Supplier, string RefCol, string EAN, string UPC, string UnitPrice, string Tarif)
        {

            String LabelName = "PRICES FW 18";
            try
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandType = CommandType.Text;
                    command.CommandText = @"UPDATE tblItemMaster SET SUPPLIERLONGCODE=@SupplierLongCode,LABEL=@Label,UPC=@UPC,
                                            COLORCODE=@ColorCode,HARMONIZEDTARIFFSCHEDULE=@Tarif,UNITPRICE=@unitPrice,LASTMODIFIEDAT= @ModifiedDate
                                            WHERE BUYERLONGCODE='LACOSTE' AND SUPPLIERLONGCODE=@SupplierLongCode AND EAN=@EAN";

                    command.Parameters.AddWithValue("@SupplierLongCode", Supplier);
                    command.Parameters.AddWithValue("@Label", LabelName);
                    command.Parameters.AddWithValue("@UPC", (UPC == "#N/A") ? Convert.DBNull : UPC);
                    command.Parameters.AddWithValue("@EAN", (EAN == "#N/A") ? Convert.DBNull : EAN);
                    command.Parameters.AddWithValue("@ColorCode", RefCol);
                    command.Parameters.AddWithValue("@Tarif", Tarif);
                    command.Parameters.AddWithValue("@unitPrice", UnitPrice);
                    command.Parameters.AddWithValue("@ModifiedDate", DateTime.Now);
                    var result = command.ExecuteScalar();
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        private string createDateFolder(string archivePath)
        {
            string thisYear = DateTime.Now.ToString("yyyy");
            string thisMth = DateTime.Now.ToString("MM");
            string thisday = DateTime.Now.ToString("dd");

            archivePath = Path.Combine(archivePath, thisYear);
            if (Directory.Exists(archivePath) == false)
            {
                Directory.CreateDirectory(archivePath);
            }
            archivePath = Path.Combine(archivePath, thisMth);
            if (Directory.Exists(archivePath) == false)
            {
                Directory.CreateDirectory(archivePath);
            }
            archivePath = Path.Combine(archivePath, thisday);
            if (Directory.Exists(archivePath) == false)
            {
                Directory.CreateDirectory(archivePath);
            }

            return archivePath;
        }

        private bool ConvertToExcelWithNPOI(string excelFileName, string worksheetName, string[] csvLines)
        {
            if (csvLines == null || csvLines.Count() == 0)
            {
                return false;
            }

            try
            {
                int rowCount = 0;
                int colCount = 0;
                int qtyResult;

                IWorkbook workbook = new HSSFWorkbook();
                ISheet worksheet = workbook.CreateSheet(worksheetName);

                HSSFFont hFont = (HSSFFont)workbook.CreateFont();
                hFont.FontHeightInPoints = 11;
                hFont.FontName = "Calibri";
                HSSFCellStyle hStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                hStyle.SetFont(hFont);
                worksheet.SetColumnWidth(0, 15 * 256);
                IDataFormat dataFormatCustom = workbook.CreateDataFormat();

                //--------------------------------Title---------------------------------------------
                //worksheet.CreateRow(0);
                //worksheet.GetRow(0).CreateCell(0).SetCellValue("UPC");
                //worksheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
                //worksheet.GetRow(0).CreateCell(2).SetCellValue("Unit Type");

                //column width
                worksheet.SetColumnWidth(0, 7000);
                worksheet.SetColumnWidth(1, 5000);
                worksheet.SetColumnWidth(2, 5000);
                worksheet.SetColumnWidth(3, 5000);
                worksheet.SetColumnWidth(4, 5000);
                worksheet.SetColumnWidth(5, 3000);

                //header style
                ICellStyle headerStyle = workbook.CreateCellStyle();
                headerStyle.FillForegroundColor = IndexedColors.Yellow.Index;
                headerStyle.FillPattern = FillPattern.SolidForeground;

                var headerFont = workbook.CreateFont();
                headerFont.FontHeightInPoints = 11;
                headerFont.FontName = "Calibri";
                headerFont.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                headerStyle.SetFont(headerFont);

                IRow rowHeader = worksheet.CreateRow(rowCount);
                ICell cellSupplier = rowHeader.CreateCell(0);
                ICell cellTarif = rowHeader.CreateCell(1);
                ICell cellRefCol = rowHeader.CreateCell(2);
                ICell cellEAN = rowHeader.CreateCell(3);
                ICell cellUPC = rowHeader.CreateCell(4);
                ICell cellUnitPrice = rowHeader.CreateCell(5);

                cellSupplier.SetCellValue("SUPPLIER");
                cellTarif.SetCellValue("TARIF");
                cellRefCol.SetCellValue("REF COL");
                cellEAN.SetCellValue("EAN");
                cellUPC.SetCellValue("UPC");
                cellUnitPrice.SetCellValue("PRICES USD");

                cellSupplier.CellStyle = headerStyle;
                cellTarif.CellStyle = headerStyle;
                cellRefCol.CellStyle = headerStyle;
                cellEAN.CellStyle = headerStyle;
                cellUPC.CellStyle = headerStyle;
                cellUnitPrice.CellStyle = headerStyle;

                rowCount = rowCount + 1;

                foreach (var line in csvLines)
                {
                    IRow row = worksheet.CreateRow(rowCount);
                    colCount = 0;
                    foreach (var col in line.Split(_tab))
                    {
                        HSSFCell cell = (HSSFCell)row.CreateCell(colCount);

                        //checked Qty column is number, if the ceil value is number,
                        //change and set data type to integer, otherwise keep original value
                        if (colCount == 1)
                        {
                            if (int.TryParse(col, out qtyResult))
                            {
                                cell.SetCellValue(Convert.ToInt64(col));
                            }
                            else
                            {
                                cell.SetCellValue(col);
                            }
                        }
                        else
                        {
                            cell.SetCellValue(col);
                        }
                        cell.CellStyle = hStyle;
                        cell.CellStyle.DataFormat = dataFormatCustom.GetFormat("General");

                        int myInt;
                        Regex r = new Regex(@"\d{2,4}/\d{2,4}/\d{2,4}$");
                        if (r.Match(col).Success)
                        {
                            DateTime formatedDate = DateTime.ParseExact(col, "MM/dd/yy", CultureInfo.InvariantCulture);
                            cell.CellStyle.DataFormat = dataFormatCustom.GetFormat("MM/dd/yy");
                        }
                        else if (int.TryParse(col, out myInt))
                        {
                            cell.CellStyle.Alignment = HorizontalAlignment.Left;
                            cell.CellStyle.DataFormat = dataFormatCustom.GetFormat("0");
                        }
                        colCount++;
                    }
                    rowCount++;
                }

                using (FileStream fileWriter = File.Create(excelFileName))
                {   // Write Excel file
                    workbook.Write(fileWriter);
                    fileWriter.Close();
                }

                worksheet = null;
                workbook = null;
                return true;  // success

            }
            catch (Exception ex)
            {
                //Logger.Error("Exception caught during DoWork:", ex, BuyerShortCode, DocumentCode, CorrelationId, excelFileName, 0);
                return false;
            }
        }

        public static bool IsNullOrWhiteSpace(String value)
        {
            if (value == null) return true;

            for (int i = 0; i < value.Length; i++)
            {
                if (!Char.IsWhiteSpace(value[i])) return false;
            }

            return true;
        }

        // Modified by Calvin (2020/1/13) - Helper function to log the missing field and quit the program
        private string failedWithMissingColumn(string msg, int rowNumber, string inputFile) {
            msg = msg + " at row " + rowNumber;
            //Logger.Error(msg, BuyerShortCode, CorrelationId);
            Console.WriteLine(msg);

            string failedFile = Path.Combine(_failedFolder, Path.GetFileName(inputFile));
            File.Move(inputFile, failedFile);

            System.Environment.Exit(-1);
            return "NULL";
        }
    }
}
