using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Windows.Forms.VisualStyles;
using Common;
using DAL;

using LSExtensionWindowLib;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
//using Oracle.DataAccess.Client;
using LSSERVICEPROVIDERLib;
using Telerik.WinControls.UI;
using Telerik.WinControls.UI.Export;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;
using Point = System.Drawing.Point;
using XlHAlign = Microsoft.Office.Core.XlHAlign;
using Oracle.DataAccess.Client;


namespace revenewSameryByClinet
{

    [ComVisible(true)]
    [ProgId("revenewSameryByClinet.revenewSameryByClinetcls")]
    public partial class revenewSameryByClinet : UserControl, IExtensionWindow
    {
        private int rowToStart = 15;
        int firstDynamicCol = 5;
        public bool DEBUG;
        private IDataLayer dal;
        public Application ExcelApp;//= new Excel.Application();
        public Excel._Worksheet WorkSheet;
        static string DATEFORMAT = "dd/MM/yyyy";
        Dictionary<Client, string> clientReport;



        List<string> _listOfTestTempletEx = new List<string>();
        private DataTable _dataTable = new DataTable();
        private List<DataClinetAdress> _dataTableArr;// = new List<DataClinetAdress>();
        public List<Client> clientArr = new List<Client>();
        public List<Address> AdressArr = new List<Address>();
        private List<ClientObj> _ClientObj;
        private List<ClientContract> _ClientContract;


        private string _connectionString;
        private OracleConnection oa;
        private OracleCommand cmd;
        private OracleCommand cmd1;


        PhraseHeader phraseHeader;
        PhraseEntry firstOrDefault;
        string savePath = null;

        #region Ctor

        public revenewSameryByClinet()
        {
            InitializeComponent();
            //    ExcelApp.Workbooks.Add();
            //   WorkSheet = ExcelApp.ActiveSheet;
            this.BackColor = Color.FromName("Control");

        }

        #endregion


        #region private members


        private IExtensionWindowSite2 _ntlsSite;
        private INautilusDBConnection _ntlsCon;
        private INautilusProcessXML _processXml;

        #endregion


        #region Implementation of IExtensionWindow

        public bool CloseQuery()
        {
            return true;
        }

        public void Internationalise()
        {
        }

        public void SetSite(object site)
        {
            _ntlsSite = (IExtensionWindowSite2)site;
            _ntlsSite.SetWindowInternalName("revenewSameryByClinet");
            _ntlsSite.SetWindowRegistryName("revenewSameryByClinet");
            _ntlsSite.SetWindowTitle("revenewSameryByClinet");
        }
        public void PreDisplay()
        {
            if (DEBUG)
            {
                //OracleConnection oa  =new OracleConnection("DATA SOURCE=MICNAUT;PASSWORD=lims;USER ID=LIMS");
                oa = new OracleConnection("DATA SOURCE=MICNAUT;PASSWORD=lims_sys;USER ID=lims_sys");
                oa.Open();


                dal = new MockDataLayer();
            }
            else
            {
                Utils.CreateConstring(_ntlsCon);
                oa = (GetConnection(_ntlsCon));
                //oa.Open();

                dal = new DataLayer();

            }

            dal.Connect();


            var customers = dal.GetClients();
            //label14.Text = "max-" + (customers.Max(I => I.ClientId) + 1);
            radDropDownListCustomer.DisplayMember = "name";
            radDropDownListCustomer.DataSource = customers;

            var labs = dal.GetLabs();
            ddlLabs.DisplayMember = "LabHebrewName";
            ddlLabs.DataSource = labs;

            SetDefultSetDates();

            phraseHeader = dal.GetPhraseByName("Location folders");
            firstOrDefault = phraseHeader.PhraseEntries.FirstOrDefault(pe => pe.PhraseDescription == "excel pic");
            savePath = phraseHeader.PhraseEntries.FirstOrDefault(pe => pe.PhraseDescription == "revenew excel").PhraseName;
        }

        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {
            OracleConnection connection = null;
            if (ntlsCon != null)
            {
                //initialize variables
                string rolecommand;
                //try catch block
                try
                {
                    _connectionString = ntlsCon.GetADOConnectionString();
                    var splited = _connectionString.Split(';');
                    _connectionString = "";
                    for (int i = 1; i < splited.Count(); i++)
                    {
                        _connectionString += splited[i] + ';';
                    }

                    //create connection
                    connection = new OracleConnection(_connectionString);

                    //open the connection
                    connection.Open();

                    //get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    //set role lims user
                    if (limsUserPassword == "")
                    {
                        //lims_user is not password protected 
                        rolecommand = "set role lims_user";
                    }
                    else
                    {
                        //lims_user is password protected
                        rolecommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    //set the oracle user for this connection
                    OracleCommand command = new OracleCommand(rolecommand, connection);

                    //try/catch block
                    try
                    {
                        //execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception f)
                    {
                        //throw the exeption
                        MessageBox.Show("Inconsistent role Security : " + f.Message);
                    }

                    //get session id
                    double sessionId = ntlsCon.GetSessionId();

                    //connect to the same session 
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    //Build the command 
                    command = new OracleCommand(sSql, connection);

                    //execute the command
                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    //throw the exeption
                    MessageBox.Show("Err At GetConnection: " + e.Message);
                }
            }
            return connection;
        }
        private void SetDefultSetDates()
        {
            var today = DateTime.Today;
            var month = new DateTime(today.Year, today.Month, 1);
            var first = month.AddMonths(-1);
            var last = month.AddDays(-1);
            radDateTimePickerFrom.Value = first;
            radDateTimePickerTo.Value = last;
        }

        public WindowButtonsType GetButtons()
        {
            return LSExtensionWindowLib.WindowButtonsType.windowButtonsNone;
        }

        public bool SaveData()
        {
            return false;
        }

        public void SetServiceProvider(object serviceProvider)
        {
            var sp = serviceProvider as NautilusServiceProvider;
            _processXml = Common.Utils.GetXmlProcessor(sp);
            _ntlsCon = Common.Utils.GetNtlsCon(sp);
        }

        public void SetParameters(string parameters)
        {

        }

        public void Setup()
        {

        }

        public WindowRefreshType DataChange()
        {
            return LSExtensionWindowLib.WindowRefreshType.windowRefreshNone;
        }

        public WindowRefreshType ViewRefresh()
        {
            return LSExtensionWindowLib.WindowRefreshType.windowRefreshNone;
        }

        public void refresh()
        {
        }

        public void SaveSettings(int hKey)
        {
        }

        public void RestoreSettings(int hKey)
        {
        }

        public void Close()
        {

        }


        #endregion

        private DateTime fromDate;
        private DateTime toDate;





        private void GetValue(ClientObj _ClientObj, Contract contract, int rowCount, DataRow countRow, DataRow price,
            DataRow forPaiment)
        {

            foreach (SdjObj sdg in _ClientObj.sdgs)
            {
                bool addrow = false;



                ConstForSdgRow(sdg, rowCount);
                foreach (TestObj Test in sdg.tests)
                {
                    var test = _dataTable.Rows[rowCount][Test.TestName];

                    _dataTable.Rows[rowCount][Test.TestName] = Test.CountTest;

                    addrow = true;
                    var value = Test.Price;
                    price[Test.TestName] = value != null ? value : 0;
                    countRow[Test.TestName] = Test.CountTest;
                    forPaiment[Test.TestName] = Test.Price * Test.CountTest;
                    DataTable da;

                }

                if (addrow)
                {
                    rowCount++;
                }
                else
                {
                    _dataTable.Rows.RemoveAt(_dataTable.Rows.Count - 1);
                }
            }
            forPaiment["סהכ לתשלום"] = 0;
            double a = 0;
            for (int i = 1; i < forPaiment.ItemArray.Count(); i++)
            {
                var b = forPaiment[i] != DBNull.Value ? Convert.ToDouble(forPaiment[i]) : 0;

                a = a + b;
            }
            forPaiment["סהכ לתשלום"] = a;
        }

        private List<Sdg> SdgByDate(List<Sdg> sdgList)
        {
            List<Sdg> sdgByDate;
            if (radCheckBoxAproveOnly.Checked)
            {
                sdgByDate = (from item in sdgList
                             where
                                 Convert.ToDateTime(item.AUTHORISED_ON).Date >= fromDate.Date &&
                                 Convert.ToDateTime(item.AUTHORISED_ON).Date < toDate.Date
                                 && item.Status == "A"
                             orderby item.CREATED_ON
                             select item).ToList();
            }
            else
            {
                sdgByDate = (from item in sdgList
                             where
                                 Convert.ToDateTime(item.CREATED_ON).Date >= fromDate.Date &&
                                 Convert.ToDateTime(item.CREATED_ON).Date <= toDate.Date
                             orderby item.CREATED_ON
                             select item).ToList();
            }
            return sdgByDate;
        }


        private void DeatailsForClient(Client clinet)
        {
            Address address = dal.GetAddresses("CLIENT", clinet.ClientId).FirstOrDefault(x => x.AddressType == "C");
            if (address != null)
            {
                radTextBoxFax.Text = address.Fax;
                radTextBoxPhone2.Text = address.Phone;
                radTextBoxAdress.Text = address.FullAddress;
                radTextBoxEmail.Text = address.Email;
            }
        }

        private void ConstForSdgRow(SdjObj sdg, int rowCount)
        {
            _dataTable.Rows.Add(_dataTable.NewRow());
            _dataTable.Rows[rowCount]["הזמנה"] = sdg.SdjName;
            if (sdg.DeliveryDate != null)
            {

                var dt = Convert.ToDateTime(sdg.DeliveryDate).ToString(DATEFORMAT);
                _dataTable.Rows[rowCount]["תאריך"] = dt;
            }
            _dataTable.Rows[rowCount]["מעבדה"] = sdg.LabName;


            _dataTable.Rows[rowCount]["הזמנת לקוח"] = sdg.ExternalReference;
        }


        private DataRow ConstFiled(out DataRow forPaiment, out DataRow price)//עריכת שורות הסה"כ
        {

            var countRow = _dataTable.NewRow();
            price = _dataTable.NewRow();
            forPaiment = _dataTable.NewRow();
            countRow["הזמנה"] = "סהכ";
            price["הזמנה"] = "מחיר לאחר הנחה";
            forPaiment["הזמנה"] = "לתשלום";
            return countRow;
        }


        private void SetDataTable(ClientObj _ClientObj)//עריכת כותרות העמודות בטבלה
        {
            _dataTable = new DataTable();
            _dataTable.Columns.Clear();

            foreach (var column in _ClientObj.Columns)
            {
                _dataTable.Columns.Add(column);
            }

        }


        public void SaveFile(string excelFilePath)
        {
            if (!string.IsNullOrEmpty(excelFilePath))
            {
                try
                {
                    WorkSheet.SaveAs(excelFilePath);
                    ExcelApp.Quit();
                    // MessageBox.Show("Excel file saved!");
                }
                catch (Exception ex)
                {
                    throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                        + ex.Message);
                }
            }
        }

        public void AddFiledToSpecificPlace(int cell, int row, string text)
        {
            WorkSheet.Cells[row, cell] = text;
        }

        public void AddLogo(string picPath)
        {
            AddPictureToFile(picPath, 56, 0, 300, 50);
        }
        public void AddPictureToFile(string picPath, float left, float top, float width, float high)
        {
            WorkSheet.Shapes.AddPicture(picPath, MsoTriState.msoFalse, MsoTriState.msoCTrue, left, top, width, high);
        }
        public void SetStyleToCell(string fromCell, string toCell, bool bold, string fontName, int size, bool italic, bool underline, XlHAlign centerText = XlHAlign.xlHAlignGeneral)
        {
            WorkSheet.Range[fromCell, toCell].Font.Bold = bold;
            WorkSheet.Range[fromCell, toCell].Font.Name = fontName;
            WorkSheet.Range[fromCell, toCell].Font.Size = size;
            WorkSheet.Range[fromCell, toCell].Font.Italic = italic;
            WorkSheet.Range[fromCell, toCell].Font.Underline = underline;
            //  if (centerText)
            //  {
            WorkSheet.Range[fromCell, toCell].HorizontalAlignment = centerText;
            //  }
        }
        private void AllBorders(Excel.Borders _borders)
        {
            _borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders.Color = Color.Black;
        }
        public void ShhetDefeniton()
        {
            var selectedRange = ExcelApp.ActiveSheet.Range["A1", "Z99"];
            selectedRange.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            selectedRange.Columns.AutoFit();
        }



        string[] formats = {DATEFORMAT, "M/d/yyyy h:mm:ss tt", "M/d/yyyy h:mm tt",
                       "MM/dd/yyyy hh:mm:ss", "M/d/yyyy h:mm:ss",
                       "M/d/yyyy hh:mm tt", "M/d/yyyy hh tt",
                       "M/d/yyyy h:mm", "M/d/yyyy h:mm",
                       "MM/dd/yyyy hh:mm", "M/dd/yyyy hh:mm"};



        private void radButton2_Click(object sender, EventArgs e)
        {
            try
            {
                var fromDate = radDateTimePickerFrom.Value.Date;
                var toDate = radDateTimePickerTo.Value.AddDays(1);

                if (fromDate > toDate)
                {
                    MessageBox.Show("טווח תאריכים לא הגיוני");
                    return;
                }

                var lab = (LabInfo)ddlLabs.SelectedItem.DataBoundItem;

                ExportPerList(lab, fromDate, toDate, !radCheckBoxAproveOnly.Checked, true);
                lblnum.Text = "";

                MessageBox.Show("הקובץ נשמר בתיקייה");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void revenewSameryByClinet_Resize(object sender, EventArgs e)
        {
            lblHeader.Location = new Point(Width / 2 - lblHeader.Width / 2, lblHeader.Location.Y);
            radPanel3.Location = new Point(Width / 2 - radPanel3.Width / 2, radPanel3.Location.Y);
        }

        private void radDropDownListCustomer_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {
            RefrshTable();
            if (radDropDownListCustomer.SelectedItem != null)
            {
                var clinet = (Client)radDropDownListCustomer.SelectedItem.DataBoundItem;
                if (clinet != null)
                {
                    DeatailsForClient(clinet);
                }
            }
        }

        private void radDateTimePickerTo_ValueChanged(object sender, EventArgs e)
        {
            RefrshTable();
        }

        private void radDateTimePickerFrom_ValueChanged(object sender, EventArgs e)
        {
            RefrshTable();
        }

        private void radCheckBoxAproveOnly_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            RefrshTable();
        }

        private void RefrshTable()
        {
            _dataTable = new DataTable();
            radGridView1.DataSource = _dataTable;
        }

        private void AllClient_CheckedChanged(object sender, EventArgs e)
        {
            if (AllClient.Checked)
            {
                radButton1.Text = "אקסל";

            }
            else
            {
                radButton1.Text = "סנן";

            }

            //pnlClientId.Visible = AllClient.Checked;

            radDropDownListCustomer.Enabled = !AllClient.Checked;

        }

        private void radPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }


        private void revenewSameryByClinet_Load(object sender, EventArgs e)
        {

        }





        //#####################################################################################################


        private void radButton1_Click(object sender, EventArgs e)
        {
            try
            {
                //  MessageBox.Show ( "גרסת נסיון" );
                fromDate = radDateTimePickerFrom.Value.Date;
                toDate = radDateTimePickerTo.Value.AddDays(1);

                if (fromDate > toDate)
                {
                    MessageBox.Show("טווח תאריכים לא הגיוני");
                    return;
                }

                //הורדתי
                //clientReport = new Dictionary<Client, string>();

                if (AllClient.Checked)
                {

                    var lab = (LabInfo)ddlLabs.SelectedItem.DataBoundItem;

                    ExportPerList(lab, fromDate, toDate, !radCheckBoxAproveOnly.Checked, false);
                    lblnum.Text = "";

                }

                else
                {
                    //RefrshTable();
                    var clinet = (Client)radDropDownListCustomer.SelectedItem.DataBoundItem;
                    var lab = (LabInfo)ddlLabs.SelectedItem.DataBoundItem;

                    var clinetID = (int)clinet.ClientId;

                    _ClientObj = new List<ClientObj>();
                    _ClientObj = GetSdgByClinetNotCancele2(fromDate, toDate, lab, !radCheckBoxAproveOnly.Checked, true);
                    if (_ClientObj.Count() != 0)
                    {
                        SetDataTable(_ClientObj[0]);
                        int rowCount = 0;
                        DataRow forPaiment;
                        DataRow price;
                        var countRow = ConstFiled(out forPaiment, out price);//הוספת כותרת לשורות סה"כ ותשלום
                        var contract = dal.GetLastContract(clinetID);//מקבל את החוזה שהיה בתוקף בזמן הקמת ההזמנה

                        GetValue(_ClientObj[0], contract, rowCount, countRow, price, forPaiment);//בניית טבלת טסטים
                        _dataTable.Rows.Add(countRow);
                        _dataTable.Rows.Add(price);

                        _dataTable.Rows.Add(forPaiment);
                        radGridView1.DataSource = _dataTable;
                        radButtonExcel.Enabled = true;
                    }
                    else
                    {
                        MessageBox.Show("לא קיימות דרישות בטווח תאריכים זה");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //הורדתי
                //generateReport();

            }
        }




        private void ExportPerList(LabInfo lab, DateTime fromDate, DateTime to, bool isCreatedOn, bool SingleCustomer)
        {
            int countValidSDG = 0;
            _ClientObj = new List<ClientObj>();
            _ClientObj = GetSdgByClinetNotCancele2(fromDate, to, lab, isCreatedOn, SingleCustomer);
            if (_ClientObj.Count() != 0)
            {
                lblnum.Text += "-" + _ClientObj.Count.ToString();
                Console.WriteLine("Total needs to ckeck: " + countValidSDG);

                if (_ClientObj.Count > 0)
                    ExportList2(false);
            }
            else
            {
                MessageBox.Show("לא קיימות דרישות בטווח תאריכים זה");
            }

        }



        private List<ClientObj> GetSdgByClinetNotCancele2(DateTime fromDate, DateTime to, LabInfo lab, bool isCreatedOn, bool SingleCustomer)
        {
            _ClientObj = new List<ClientObj>();
            Client SinglCustomer = null;

            //In sql developer  אפשר להריץ את השאילתא הזו רק מלימססיס
            //lims doesn't return any value

            string strSQL = "";
            string strSQL1 = "SELECT C.CLIENT_ID CLIENTID,C.CLIENT_CODE CLIENTCODE,C.NAME CLIENTNAME,ADR.PHONE PHONE,ADR.EMAIL EMAIL,ADR.FAX FAX,ADR.ADDRESS_LINE_2 FULLADRESS,D.SDG_ID SDJID,D.NAME SDJNAME,LU.U_LAB_HEBREW_NAME LABNAME,DU.U_LAB_INFO_ID LABID, " +
                        "D.EXTERNAL_REFERENCE EXTERNALREFERENCE,D.DELIVERY_DATE DELIVERYDATE,TTEX.U_TEST_TEMPLATE_EX_ID TESTID,TTEX.NAME TESTNAME,COUNT(TTEX.U_TEST_TEMPLATE_EX_ID) COUNTTEST " +
                        "FROM lims_sys.SDG D " +
                        "INNER JOIN lims_sys.SDG_USER DU ON D.SDG_ID=DU.SDG_ID " +
                        "INNER JOIN lims_sys.CLIENT C ON DU.U_SDG_CLIENT = C.CLIENT_ID " +
                        "INNER JOIN lims_sys.CLIENT_USER CU ON CU.CLIENT_ID = C.CLIENT_ID " +
                        "INNER JOIN lims_sys.SAMPLE S ON S.SDG_ID=D.SDG_ID " +
                        "INNER JOIN lims_sys.ALIQUOT A ON S.SAMPLE_ID=A.SAMPLE_ID " +
                        "INNER JOIN lims_sys.ALIQUOT_USER AU ON AU.ALIQUOT_ID=A.ALIQUOT_ID " +
                        "INNER JOIN lims_sys.U_TEST_TEMPLATE_EX_USER TTEXU ON TTEXU.U_TEST_TEMPLATE_EX_ID=AU.U_TEST_TEMPLATE_EXTENDED " +
                        "INNER JOIN lims_sys.U_TEST_TEMPLATE_EX TTEX ON TTEXU.U_TEST_TEMPLATE_EX_ID=TTEX.U_TEST_TEMPLATE_EX_ID " +
                        "INNER JOIN lims_sys.U_LABS_INFO_USER LU ON LU.U_LABS_INFO_ID =DU.U_LAB_INFO_ID " +
                        "LEFT JOIN lims_sys.ADDRESS ADR ON ( ADR.ADDRESS_ITEM_ID = C.CLIENT_ID  AND ADR.ADDRESS_TABLE_NAME = '{5}' AND ADR.ADDRESS_TYPE = '{6}' ) " +
                             "WHERE Du.U_LAB_INFO_ID = '{2}' AND S.STATUS != '{3}' AND A.STATUS != '{3}' AND Au.U_CHARGE = '{4}'  ";
            string strSQL2 = "AND C.CLIENT_ID ='{7}' ";
            string strSQL3 = "AND D.CREATED_ON >= to_date('{0}','dd/MM/yyyy') AND D.CREATED_ON <=to_date('{1}','dd/MM/yyyy') ";
            string strSQL4 = "AND D.STATUS != '{3}' AND D.AUTHORISED_ON >= to_date('{0}','dd/MM/yyyy') AND D.AUTHORISED_ON <=to_date('{1}','dd/MM/yyyy') ";
            string strSQL5 = "GROUP BY C.CLIENT_ID,C.CLIENT_CODE,C.NAME,ADR.PHONE,ADR.EMAIL,ADR.FAX,ADR.ADDRESS_LINE_2,D.SDG_ID,D.NAME,LU.U_LAB_HEBREW_NAME,DU.U_LAB_INFO_ID,D.EXTERNAL_REFERENCE,D.DELIVERY_DATE,TTEX.U_TEST_TEMPLATE_EX_ID,TTEX.NAME,TTEXU.U_PRICE ORDER BY C.NAME,D.NAME,TTEX.NAME ";

            try
            {
                if (SingleCustomer)
                {
                    strSQL1 += strSQL2;
                    SinglCustomer = (Client)radDropDownListCustomer.SelectedItem.DataBoundItem;
                }
                if (isCreatedOn)
                {
                    strSQL1 += strSQL3;
                }
                else
                {
                    strSQL1 += strSQL4;
                }
                strSQL1 += strSQL5;

                if (SingleCustomer)
                {
                    strSQL = string.Format(strSQL1, fromDate.ToString("dd/MM/yyyy"), to.ToString("dd/MM/yyyy"), lab.LabInfoId, "X", "T", "CLIENT", "C", SinglCustomer.ClientId);
                }
                else
                {
                    strSQL = string.Format(strSQL1, fromDate.ToString("dd/MM/yyyy"), to.ToString("dd/MM/yyyy"), lab.LabInfoId, "X", "T", "CLIENT", "C");
                }

                Logger.WriteLogFile(strSQL);
                cmd = new OracleCommand(strSQL, oa);

                OracleDataReader res = cmd.ExecuteReader();

                while (res.Read()) //insert the sdgs to the client list
                {

                    int cId = int.Parse(res["CLIENTID"].ToString());
                    ClientObj client = _ClientObj.Where(item => item.ClientId == cId).FirstOrDefault();
                    if (client == null)
                    {
                        ClientObj newclient = InsertNewClient(res); //add new client
                        _ClientObj.Add(newclient);

                    }
                    else
                    {
                        int sId = int.Parse(res["SDJID"].ToString());
                        SdjObj sdg = client.sdgs.Where(item => item.SdjId == sId).FirstOrDefault();
                        if (sdg == null)
                        {
                            SdjObj newsdg = InsertSdgForClient(res); //add new sgd to client
                            client.sdgs.Add(newsdg);
                        }
                        else
                        {
                            int tId = int.Parse(res["TESTID"].ToString());
                            TestObj test = sdg.tests.Where(item => item.TestId == tId).FirstOrDefault();
                            if (test == null)
                            {
                                TestObj newtest = InsertTestForSdg(res); //add new test to sdg
                                sdg.tests.Add(newtest);
                            }
                            else
                            {
                                MessageBox.Show("Error, No test found");
                            }
                        }
                    }

                }

                var clientIds = _ClientObj.Select(x => x.ClientId);

                if (_ClientObj.Count() == 0)
                {
                    return _ClientObj;
                }
                var jpined = String.Join(",", clientIds.ToArray());

                GetFinalPrice(jpined);

                List<string> constColumns = new List<string>() { "הזמנה", "תאריך", "מעבדה", "הזמנת לקוח" };
                Logger.WriteLogFile("סיום השליפה");
                Logger.WriteLogFile("clients found: " + _ClientObj.Count());

                foreach (ClientObj client in _ClientObj) //Inserting the column headings of the table in Excel
                {

                    List<string> AllTests4client = new List<string> { };
                    List<string> testNames4sdg = new List<string> { };

                    AllTests4client = AllTests4client.Concat(constColumns).ToList();

                    foreach (SdjObj sdg in client.sdgs) //Finding all the existing tests for the customer
                    {
                        testNames4sdg = sdg.tests.Select(test => test.TestName).ToList();
                        AllTests4client = AllTests4client.Concat(testNames4sdg).ToList();
                    }
                    Logger.WriteLogFile("Client: " + client.ClientName + " with: " + client.sdgs.Count() + " sdgs and " + (AllTests4client.Count() - 4) + " tests");
                    AllTests4client = AllTests4client.Distinct().ToList();

                    AllTests4client.Add("סהכ לתשלום");

                    client.Columns = AllTests4client;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

            return _ClientObj;
        }

        private ClientObj InsertNewClient(OracleDataReader res) //add new client
        {
            ClientObj newClient = new ClientObj()
            {
                ClientId = int.Parse(res["CLIENTID"].ToString()),
                ClientCode = res["CLIENTCODE"].ToString(),
                ClientName = res["CLIENTNAME"].ToString(),
                Adress = res["FULLADRESS"].ToString(),
                Email = res["EMAIL"].ToString(),
                Fax = res["FAX"].ToString(),
                Fhone = res["PHONE"].ToString()
            };

            SdjObj newsdg = InsertSdgForClient(res);
            newClient.sdgs.Add(newsdg);
            return newClient;
        }

        private SdjObj InsertSdgForClient(OracleDataReader res) //add new sdg for client
        {
            //   var strDate = "";
            DateTime? strDateD = null;
            try
            {



                if (!string.IsNullOrEmpty(res["DELIVERYDATE"].ToString()))
                {
                    //    strDate = DateTime.Parse(res["DELIVERYDATE"].ToString()).ToString("dd/MM/yyyy");
                    strDateD = DateTime.Parse(res["DELIVERYDATE"].ToString());//.ToString("dd/MM/yyyy");
                }
            }
            catch (Exception)
            {

                Logger.WriteLogFile("delivery date is not valid");
            }
            SdjObj newsdg = new SdjObj()
            {
                SdjId = int.Parse(res["SDJID"].ToString()),
                SdjName = res["SDJNAME"].ToString(),
                LabName = res["LABNAME"].ToString(),
                ExternalReference = res["EXTERNALREFERENCE"].ToString(),
                DeliveryDate = strDateD.HasValue ? strDateD.Value : default(DateTime)

            };

            TestObj newtest = InsertTestForSdg(res);
            newsdg.tests.Add(newtest);
            return newsdg;
        }

        private TestObj InsertTestForSdg(OracleDataReader res) //add new test for sdg
        {
            TestObj newtest = new TestObj()
            {
                TestId = int.Parse(res["TESTID"].ToString()),
                TestName = res["TESTNAME"].ToString(),
                //Price = int.Parse(res["PRICE"].ToString()),
                CountTest = int.Parse(res["COUNTTEST"].ToString())
            };
            return newtest;

        }


        //Receiving the correct price for the test the correct price for the test
        private void GetFinalPrice(string jpined)
        {
            try
            {


                _ClientContract = new List<ClientContract>();

                string strSQL = string.Format(
                        "SELECT A.U_CLIENT_ID CLIENTID,TTEX.U_TEST_TEMPLATE_EX_ID TESTID,NVL(CD.U_FINAL_PRICE,TTEX.U_PRICE) PRICE " +
                        "FROM lims_sys.U_CONTRACT_USER A " +
                        "INNER JOIN lims_sys.U_CONTRACT_DATA_USER CD ON CD.U_CONTRACT_ID  = A.U_CONTRACT_ID " +
                        "INNER JOIN lims_sys.U_TEST_TEMPLATE_EX_USER TTEX ON TTEX.U_TEST_TEMPLATE_EX_ID=CD.U_TEST_TEMPLATE_EX_ID " +
                        "and A.U_CLIENT_ID in ( " + jpined + ")" +
                        " AND A.U_CONFIRM_DATE =(SELECT MAX(B.U_CONFIRM_DATE)FROM lims_sys.U_CONTRACT_USER B WHERE B.U_CLIENT_ID=A.U_CLIENT_ID)");

                cmd1 = new OracleCommand(strSQL, oa);
                OracleDataReader resPrice = cmd1.ExecuteReader();
                int x = 0;
                while (resPrice.Read())
                {
                    x++;
                    if (x > 54)
                    {

                    }
                    int cId = int.Parse(resPrice["CLIENTID"].ToString());
                    int tId = int.Parse(resPrice["TESTID"].ToString());
                    double price = double.Parse(resPrice["PRICE"].ToString());

                    ClientObj newClient = _ClientObj.Where(i => i.ClientId == cId).FirstOrDefault();
                    if (newClient != null)
                    {
                        foreach (SdjObj sdg in newClient.sdgs)
                        {
                            TestObj newTest = sdg.tests.Where(i => i.TestId == tId).FirstOrDefault();
                            if (newTest != null)
                            {
                                newTest.Price = price;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Error, clielt not found");
                    }

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        //Creating a new excel file
        private void ExportList2(bool showMessage = true)
        {
            Logger.WriteLogFile("Starts editing the excel file");

            PhraseHeader phraseHeader = dal.GetPhraseByName("Location folders");
            PhraseEntry firstOrDefault = phraseHeader.PhraseEntries.FirstOrDefault(pe => pe.PhraseDescription == "excel pic");
            string savePath = phraseHeader.PhraseEntries.FirstOrDefault(pe => pe.PhraseDescription == "revenew excel").PhraseName;

            if (!Directory.Exists(savePath))
            {
                MessageBox.Show("לא יופק קובץ \n אנא הקם תיקייה" + savePath, "הפקת דוח", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string save = "";
            if (firstOrDefault != null)
            {
                save = firstOrDefault.PhraseName;
            }
            var from = radDateTimePickerFrom.Value.Date;
            var to = radDateTimePickerTo.Value.Date;
            ExportListOfDtataTableToExcel2(_ClientObj, 10, 0, true, from, to, save);
            string path = savePath + "\\" + ddlLabs.Text + "_" + DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss") + ".xlsx";// ".xls";

            SaveFile(path);
            Close();
            ExcelApp = null;
            //     Process.Start(path);
            if (showMessage)
                MessageBox.Show("הקובץ נשמר בתיקייה");
            Logger.WriteLogFile("Editing completed");
            lblClients.Text = path;

        }


        Excel.Range dateRange;
        public void ExportListOfDtataTableToExcel2(List<ClientObj> listtable, int rowToStart, int cellStart, bool headerBold, DateTime from, DateTime to, string path)
        {


            var ii = 0;
            ExcelApp = new Excel.Application();
            ExcelApp.DefaultSheetDirection = (int)XlDirection.xlToLeft;
            //        var workbook = ExcelApp.Workbooks.Add();

            var workbook = ExcelApp.Workbooks.Add(Missing.Value);
            WorkSheet = ExcelApp.ActiveSheet;



            WorkSheet.Cells.NumberFormat = "General";
            //var list = listtable.OrderBy(da => da.ClientName);//ממוין בשליפה

            foreach (ClientObj Client in listtable)//pass all the client
            {
                Logger.WriteLogFile("strat " + Client.ClientName + " excel sheet");
                lblClients.Text = "Working on " + Client.ClientName;
                try
                {
                    if (ii > 0)//Ashi Chenge here 2 to 1 8/3/21 don't know why
                    {
                        WorkSheet = (Excel.Worksheet)workbook.Worksheets.Add();
                        WorkSheet.DisplayRightToLeft = true;
                    }
                    else
                    {
                        WorkSheet = (Excel.Worksheet)workbook.Worksheets.Item[ii + 1];
                        WorkSheet.DisplayRightToLeft = true;

                    }

                    var sheetName = Client.ClientName.MakeSafeFilename('_');
                    if (sheetName.Length > 30)
                    {
                        sheetName = sheetName.Substring(0, 29);
                    }
                    WorkSheet.Name = string.Format(sheetName, ii + 1);


                    ii++;
                    if (Client.sdgs.Count == 0 || Client.Columns.Count == 0)
                        throw new Exception("ExportToExcel: Null or empty input table!\n");
                    for (int i = 0; i < Client.Columns.Count; i++)
                    {
                        if (headerBold)
                        {
                            WorkSheet.Cells[rowToStart, (i + 1) + cellStart] = Client.Columns[i]; //1
                            WorkSheet.Cells[rowToStart, (i + 1) + cellStart].Font.Bold = true;
                            WorkSheet.Cells[rowToStart, (i + 1) + cellStart].Font.Underline = true;

                        }
                        else
                        {
                            WorkSheet.Cells[rowToStart, (i + 1) + cellStart] = Client.Columns[i]; //1
                        }
                    }
                    SetOriention2(Client, rowToStart);


                    int x = rowToStart - 1;
                    foreach (SdjObj sdg in Client.sdgs)
                    {
                        Logger.WriteLogFile(sdg.SdjName);
                        WorkSheet.Cells[(x + 2), 1] = sdg.SdjName;


                        dateRange = (Excel.Range)WorkSheet.Cells[(x + 2), 2];
                        dateRange.EntireColumn.NumberFormat = "dd/MM/yyyy";
                        dateRange.Value = sdg.DeliveryDate;

                        WorkSheet.Cells[(x + 2), 3] = sdg.LabName;
                        WorkSheet.Cells[(x + 2), 4] = sdg.ExternalReference;

                        foreach (TestObj test in sdg.tests)
                        {
                            WorkSheet.Cells[(x + 2), Client.Columns.IndexOf(test.TestName) + 1] = test.CountTest;
                            WorkSheet.Cells[(rowToStart + 2 + Client.sdgs.Count()), Client.Columns.IndexOf(test.TestName) + 1] = test.Price;

                        }
                        x++;
                    }

                    WorkSheet.Cells[(x + 2), 1] = "סהכ";
                    WorkSheet.Cells[(x + 3), 1] = "מחיר לאחר הנחה";
                    WorkSheet.Cells[(x + 4), 1] = "לתשלום";


                    SetFormula2(rowToStart, Client);

                    AddFiledToSpecificPlace(2, 5, ": שם לקוח");
                    SetStyleToCell("B5", "B5", true, "Arial", 11, false, false);

                    AddFiledToSpecificPlace(3, 5, Client.ClientName);
                    SetStyleToCell("C5", "C5", false, "Arial", 11, false, false, (XlHAlign)Excel.XlHAlign.xlHAlignCenter);

                    AddFiledToSpecificPlace(4, 5, Client.ClientCode);
                    SetStyleToCell("D5", "D5", false, "Arial", 11, false, false, (XlHAlign)Excel.XlHAlign.xlHAlignCenter);


                    AddFiledToSpecificPlace(2, 6, ": הזמנות בין תאריכים");
                    SetStyleToCell("B6", "B6", true, "Arial", 11, false, false);
                    AddFiledToSpecificPlace(3, 6,
                                            from.ToString("dd/MM/yyyy") + " - " + to.ToString("dd/MM/yyyy"));
                    SetStyleToCell("C6", "C6", false, "Arial", 11, false, false, (XlHAlign)Excel.XlHAlign.xlHAlignCenter);

                    //להוסיף אח"כ
                    if (Client.Adress != null)
                    {
                        SetAddressDetails2(Client, true, Client.Remark);
                    }
                    string save = path;

                    if (File.Exists(save))
                    {
                        AddLogo(save);

                    }
                    ShhetDefeniton();
                    if (Client.ClientId.ToString().Equals("10541"))
                    {
                        throw new Exception("my exception");
                    }
                    //הורדתי
                    //clientReport[dataClinetAdress.Client] = "OK";
                }
                catch (Exception ex)
                {
                    Logger.WriteLogFile(ex);
                    throw new Exception("ExportToExcel: \n" + ex.Message);
                }
            }
            lblClients.Text = "Finish";
        }

        private void SetOriention2(ClientObj tbl, int rowToStart)
        {
            if (tbl.Columns.Count != firstDynamicCol)
            {
                Excel.Range formatRange;
                var s1 = GetExcelColumnName(firstDynamicCol);
                var s2 = GetExcelColumnName(tbl.Columns.Count - 1);
                formatRange = WorkSheet.Range[s1 + "" + rowToStart, s2 + "" + rowToStart];
                formatRange.Orientation = 90;
                formatRange.Font.Underline = false;
                var s3 = GetExcelColumnName(tbl.Columns.Count);
                var lastCell = s3 + "" + (rowToStart + tbl.sdgs.Count + 3);
                var tableRange = WorkSheet.Range["A" + "" + rowToStart, lastCell];
                tableRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                AllBorders(tableRange.Borders);

                WorkSheet.Range[lastCell, lastCell].Font.Bold = true;
            }
            ////Print settings
            var r = tbl.sdgs.Count + 3 + rowToStart;
            var c = GetExcelColumnName(tbl.Columns.Count + 3);
            string lastCellToPrint = string.Format("{0}{1}", c, r);
            WorkSheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            WorkSheet.PageSetup.PrintArea = "A1:" + lastCellToPrint;
            //////////////////////

            WorkSheet.Range["A" + (rowToStart + tbl.sdgs.Count + 3 - 2), "A" + (rowToStart + tbl.sdgs.Count + 3)].Font.Bold = true;
            WorkSheet.Range["A" + (rowToStart + tbl.sdgs.Count + 3 - 2), "A" + (rowToStart + tbl.sdgs.Count + 3)].HorizontalAlignment = XlHAlign.xlHAlignRight;

        }

        private void SetFormula2(int rowToStart, ClientObj tbl)
        {

            var lastRow = tbl.sdgs.Count + 3 - 1;
            var lastRowBottomTable = lastRow + rowToStart;
            //set formula (price * sum)
            int i;
            for (i = firstDynamicCol; i < tbl.Columns.Count; i++)
            {

                string colSign = GetExcelColumnName(i);

                //formula for total  test
                string f1 = string.Format("=SUM({0}{1}:{2}{3})", colSign, rowToStart + 1, colSign, lastRowBottomTable - 2);// "=SUM(F16:F19)";               
                WorkSheet.Cells[lastRowBottomTable - 1, i].Formula = f1;

                //formula for total  payment per test
                string f = "=SUM(" + colSign + "" + (lastRowBottomTable - 1) + "" + "*" + colSign + "" + lastRowBottomTable + "" + ")";
                WorkSheet.Cells[tbl.sdgs.Count + 3 + rowToStart, i].Formula = f;
            }



            //formula for total  payment
            string sumTotalFormula = string.Format("=SUM({0}{1}:{2}{3})", GetExcelColumnName(firstDynamicCol), lastRowBottomTable + 1, GetExcelColumnName(tbl.Columns.Count - 1), lastRowBottomTable + 1);
            WorkSheet.Cells[lastRowBottomTable + 1, i].Formula = sumTotalFormula;

        }
        private void SetAddressDetails2(ClientObj client, bool remarks = false, string remark = "")
        {
            var colToDetails = client.Columns.Count + 4;
            var colToAssignDetails = colToDetails + 1;
            var colToDetailsSign = GetExcelColumnName(colToDetails);
            var colToAssignDetailsSign = GetExcelColumnName(colToAssignDetails);
            var row = 1;
            string boldCell = colToDetailsSign + "" + row;
            AddFiledToSpecificPlace(colToDetails, row, "         : אימייל");
            SetStyleToCell(boldCell, boldCell, true, "Arial", 11, false, false);
            AddFiledToSpecificPlace(colToAssignDetails, row, client.Email);
            SetStyleToCell(colToAssignDetailsSign + row, colToAssignDetailsSign + row, false, "Arial", 11, false, false,
                (XlHAlign)Excel.XlHAlign.xlHAlignCenter);

            row++;
            boldCell = colToDetailsSign + "" + row;

            AddFiledToSpecificPlace(colToDetails, row, ": פקס");
            SetStyleToCell(boldCell, boldCell, true, "Arial", 11, false, false);
            AddFiledToSpecificPlace(colToAssignDetails, row, client.Fax);
            SetStyleToCell(colToAssignDetailsSign + row, colToAssignDetailsSign + row, false, "Arial", 11, false, false,
                (XlHAlign)Excel.XlHAlign.xlHAlignCenter);

            row++;
            boldCell = colToDetailsSign + "" + row;


            AddFiledToSpecificPlace(colToDetails, row, ": טלפון");
            SetStyleToCell(boldCell, boldCell, true, "Arial", 11, false, false);
            AddFiledToSpecificPlace(colToAssignDetails, row, client.Fhone);
            SetStyleToCell(colToAssignDetailsSign + row, colToAssignDetailsSign + row, false, "Arial", 11, false, false,
                (XlHAlign)Excel.XlHAlign.xlHAlignCenter);

            row++;
            boldCell = colToDetailsSign + "" + row;


            AddFiledToSpecificPlace(colToDetails, row, ": כתובת");
            SetStyleToCell(boldCell, boldCell, true, "Arial", 11, false, false);
            AddFiledToSpecificPlace(colToAssignDetails, row, client.Adress);
            SetStyleToCell(colToAssignDetailsSign + row, colToAssignDetailsSign + row, false, "Arial", 11, false, false,
                (XlHAlign)Excel.XlHAlign.xlHAlignCenter);

            if (remarks)
            {

                row++;
                boldCell = colToDetailsSign + "" + row;

                AddFiledToSpecificPlace(colToDetails, row, ": הערות");
                SetStyleToCell(boldCell, boldCell, true, "Arial", 11, false, false);
                AddFiledToSpecificPlace(colToAssignDetails, row, remark);
                SetStyleToCell(colToAssignDetailsSign + row, colToAssignDetailsSign + row, false, "Arial", 11, false,
                    false, (XlHAlign)Excel.XlHAlign.xlHAlignCenter);
            }
        }

        private void radPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

    }
}


