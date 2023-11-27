using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using System.IO;
using ExcelDataReader;
using System.Diagnostics;
using System.Net;

namespace GTMFiller
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Fields
        private const string columnWeek = "C";
        private string columnDescription = "HC";

        private const string columnDate = "A";
        private const string PROGRAM_NAME_XPATH = "//*[@id=\"MiddleContents_ddlProgram\"]";
        private const string PROJECT_NAME_XPATH = "//*[@id=\"MiddleContents_ddlProjectName\"]";
        private const string YEAR_XPATH = "//*[@id=\"MiddleContents_ddlYears\"]";
        private const string WEEK_XPATH = "//*[@id=\"MiddleContents_ddlWeek\"]";
        private const string DATE_XPATH = "//*[@id=\"MiddleContents_ddlCurrentWeek1\"]";
        private const string WBS_XPATH = "//*[@id=\"MiddleContents_ddlWBS1\"]";
        private const string HOURS_XPATH = "//*[@id=\"MiddleContents_ddlHourSpent\"]";
        private const string DESCRIPTION_XPATH = "//*[@id=\"MiddleContents_TxtDescription\"]";
        private const string SAVEBUTTON_XPATH = "//*[@id=\"MiddleContents_LnkSave\"]";
        private const string SUBMIT_XPATH = "//*[@id=\"MiddleContents_LnkSubmitTimeSheet\"]";
        private const string LogOut_XPATH = "//*[@id=\"dtlNav_ctl03_HypLogout\"]";
        private const int MillisecondsTimeout = 500;
        private const int INTROP_OFFSET = 1;


        public string ProgramName { get; set; } = "Smart Wires";
        public string UserName { get; set; }
        public string GtmPassword { get; set; }
        public string ExcelNamei { get; set; }
        public string ExcelWeeki { get; set; }
        public string ExcelDatei { get; set; }
        public string task;
        public bool Login = false;




        private int startRow = 6;
        private int endRow = 46;
        public string Attendence;

        public List<string> UsersCredentialsList = new List<string>();
        public List<List<string>> SingleUserCredentialsList = new List<List<string>>();
        public string[] credentials = new string[3];

        Thread threadSeleniumAutomation;
        MonthlyRecord employeeRecord;
        ManualResetEvent manualResetEvent = new ManualResetEvent(false);
        DailyRecordTemp r1;
        public List<string> PupulateUsersCredentialsList()
        {

            UsersCredentialsList.Add("Abdul Raqeeb-" + "abdul.raqeeb-" + "gtm");
            UsersCredentialsList.Add("Halif Sohail-" + "halif.sohail-" + "gtm");
            UsersCredentialsList.Add("Ibrar Sakhi-" + "ibrar.sakhi-" + "gtm");
            UsersCredentialsList.Add("Ahmad Malik-" + "muhammad.ahmad-" + "gtm");
            UsersCredentialsList.Add("Nawaz Akhtar-" + "nakhtar-" + "nakhtar");
            UsersCredentialsList.Add("Zohaib Nayyar-" + "zohaib.nayyar-" + "gtm");
            UsersCredentialsList.Add("Muhammad Usman-" + "usmanm-" + "gtm");
            UsersCredentialsList.Add("Habib Ullah-" + "habibu-" + "gtm");
            UsersCredentialsList.Add("Hassan Rahman-" + "mhassan-" + "gtm");
            UsersCredentialsList.Add("Muhammad Haroon-" + "muhammad.haroon-" + "gtm");
            UsersCredentialsList.Add("Haseeb Ahmad-" + "haseeb.ahmed-" + "gtm");
            UsersCredentialsList.Add("Zain Islam-" + "zain.islam-" + "gtm");
            UsersCredentialsList.Add("Adeel Kamran-" + "adeel.kamran-" + "gtm");
            UsersCredentialsList.Add("Muhammad Abubakar-" + "abu.bakar-" + "gtm");
            UsersCredentialsList.Add("Farhan Ali-" + "farhan.ali-" + "gtm");
            UsersCredentialsList.Add("Asad Faiz-" + "asad.faiz-" + "gtm");
            UsersCredentialsList.Add("Abdul Manan-" + "abdul.manan-" + "gtm");
            UsersCredentialsList.Add("Shehroz Munir-" + "shehroz.munir-" + "gtm");
            UsersCredentialsList.Add("Shehbaz Hussain-" + "shehbaz.hussain-" + "gtm");
            UsersCredentialsList.Add("Osama Khalid-" + "osama.khalid-" + "gtm");
            UsersCredentialsList.Add("Hamza Saleem-" + "hamza.saleem-" + "gtm");




            return UsersCredentialsList;

        }
        #endregion
        public bool firstTim = true;
        public string resultUrl;
        public MainWindow()
        {
            InitializeComponent();
            PupulateUsersCredentialsList();

            {
                string sharingUrl = "https://powersoft19-my.sharepoint.com/:x:/r/personal/mhassan_powersoft19_com/_layouts/15/guestaccess.aspx?email=abdul.raqeeb%40Powersoft19.com&e=4%3ApgB7J4&at=9&share=ESJ5e224vuhOhBc6-BXFrhsBChy6YwbwkcvH-miZC_enGw";
                string base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(sharingUrl));
                string encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');

                resultUrl = string.Format("https://api.onedrive.com/v1.0/shares/{0}/root/content", encodedUrl);
            }
        }
        public void AddDatatoComboBox()
        {
            try
            {
                string FilePath = tbGtmFilePath.Text;
                using (var stream = File.Open(FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        for (int i = 0; i < result.Tables.Count; i++)
                        {
                            if (!comboBox_Months.Items.Contains(result.Tables[i].TableName.ToString()))
                            {
                                comboBox_Months.Items.Add(result.Tables[i].TableName.ToString());

                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void BtnAddRecord_Click(object sender, RoutedEventArgs e)
        {
            #region temp dsiplay
            ImportDataFromExcelTemp();
            List<DataGridData> datatemp = new List<DataGridData>(); ;
            for (int i = 0; i < userstemplist.Count; i++)
            {
                datatemp = new List<DataGridData>();

                for (int h = 0; h < userstemplist[i].weekly_record.DailyRecordList.Count; h++)
                {
                    DataGridData temp = new DataGridData();
                    //temp.SingleRecords = worker[i].WeeklyRecordsList.DailyRecordsList[h];
                    temp.Description = userstemplist[i].weekly_record.DailyRecordList[h].Description;
                    temp.RecordDate = userstemplist[i].weekly_record.DailyRecordList[h].RecordDate;
                    temp.Week = userstemplist[i].weekly_record.DailyRecordList[h].Week;

                    temp.Project = userstemplist[i].weekly_record.DailyRecordList[h].Project;
                    temp.Hours = userstemplist[i].weekly_record.DailyRecordList[h].Hours;
                    temp.Name = userstemplist[i].Name;
                    temp.UserName = userstemplist[i].UserName;
                    temp.Password = userstemplist[i].Password;
                    datatemp.Add(temp);
                }
            }
            dgExcelData.ItemsSource = datatemp.ToList();
            #endregion
        }

        private void btnStartLoggingOnGtm_Click(object sender, RoutedEventArgs e)
        {

            ProgramName = tbProgramName.Text;

            threadSeleniumAutomation = new Thread(UploadEmployeesDataToGTM);

            WindowState = WindowState.Minimized;
            threadSeleniumAutomation.Start();
        }

        public void Resume()
        {
            manualResetEvent.Set();
        }

        public void Pause()
        {
            manualResetEvent.Reset();
        }

        WeeklyRecordTemp weekrecordtemp = new WeeklyRecordTemp();
        public List<UserTemp> userstemplist = new List<UserTemp>();

        #region Public Methods
        #region ImportFromExcelTemp
        public void ImportDataFromExcelTemp()
        {
            userstemplist = new List<UserTemp>();
            if (tbGtmFilePath.Text == "")
            {
                MessageBox.Show("Please add file Path");
                return;
            }
            if (comboBox_Months.SelectedItem == null || comboBox_Months.SelectedItem.ToString() == "")
            {
                MessageBox.Show("Please Select Month");
                return;
            }
            if (cbWeeklyOrMonthly.SelectedItem == null || cbWeeklyOrMonthly.SelectedItem.ToString() == "")
            {
                MessageBox.Show("Please Select a Criterea");
                return;
            }
            else if (cbWeeklyOrMonthly.SelectedItem.ToString().Contains("Weekly"))
            {
                if (comboBox_week.SelectedItem == null || comboBox_week.SelectedItem.ToString() == "")
                {
                    MessageBox.Show("Please Select Week");
                    return;

                }
            }

            if (combo_employeename.SelectedItem == null || combo_employeename.SelectedItem.ToString() == "")
            {
                MessageBox.Show("Please Select User Name");
                return;
            }

            DateTime recordDate;

            string MonthYear = comboBox_Months.SelectedItem.ToString();

            WeeklyRecordTemp weekrecordtemp = new WeeklyRecordTemp();
            //List<UserTemp> userstemplist = new List<UserTemp>();
            using (var stream = File.Open(tbGtmFilePath.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    foreach (string SingleUserCredentials in UsersCredentialsList)
                    {
                        UserTemp usertemp = new UserTemp();
                        credentials = SingleUserCredentials.Split('-');
                        usertemp.Name = credentials[0];
                        usertemp.UserName = credentials[1];
                        usertemp.Password = credentials[2];
                        usertemp.weekly_record = new WeeklyRecordTemp();
                        usertemp.weekly_record.DailyRecordList = new List<DailyRecordTemp>();
                        usertemp.weekly_record_list = new List<WeeklyRecordTemp>();
                        if (usertemp.Name == combo_employeename.SelectedItem.ToString())
                        {
                            for (int i = 1; i < result.Tables[MonthYear].Rows.Count; i++)
                            {
                                string excelName = result.Tables[MonthYear].Rows[i][3].ToString();
                                string excelWeek = result.Tables[MonthYear].Rows[i][2].ToString();

                                if (usertemp.Name == excelName)
                                {
                                    task = result.Tables[MonthYear].Rows[i][7].ToString();
                                    #region Monthly Addition
                                    if (cbWeeklyOrMonthly.SelectedItem.ToString().Contains("This month"))
                                    {
                                        DailyRecordTemp dailyrecordtemp = new DailyRecordTemp();
                                        Attendence = result.Tables[MonthYear].Rows[i][4].ToString();
                                        if (string.IsNullOrEmpty(Attendence))
                                        {
                                            MessageBox.Show($"Please enter Attendance Status for {(DateTime)result.Tables[MonthYear].Rows[i][0]} and try again");
                                            userstemplist = new List<UserTemp>();
                                            return;
                                        }
                                        if (Attendence.ToUpper() != "OFF" && Attendence.ToUpper() != "PUBLIC HOLIDAY")
                                        {
                                            dailyrecordtemp = new DailyRecordTemp();
                                            dailyrecordtemp.RecordDate = (DateTime)result.Tables[MonthYear].Rows[i][0];
                                            dailyrecordtemp.Project = result.Tables[MonthYear].Rows[i][6].ToString();
                                            if (string.IsNullOrEmpty(dailyrecordtemp.Project))
                                            {
                                                MessageBox.Show("Please enter Project and try again");
                                                userstemplist = new List<UserTemp>();
                                                return;
                                            }
                                            dailyrecordtemp.Hours = result.Tables[MonthYear].Rows[i][8].ToString();
                                            if (string.IsNullOrEmpty(dailyrecordtemp.Hours))
                                            {
                                                MessageBox.Show("Please enter Hours and try again");
                                                userstemplist = new List<UserTemp>();
                                                return;
                                            }
                                            dailyrecordtemp.Description = result.Tables[MonthYear].Rows[i][7].ToString();
                                            if (string.IsNullOrEmpty(dailyrecordtemp.Description))
                                            {
                                                MessageBox.Show("Please enter Task Description and try again");
                                                userstemplist = new List<UserTemp>();
                                                return;
                                            }
                                            dailyrecordtemp.Week = result.Tables[MonthYear].Rows[i][2].ToString();
                                            usertemp.weekly_record.DailyRecordList.Add(dailyrecordtemp);
                                            //usertemp.weekly_record.Week = result.Tables[MonthYear].Rows[i][2].ToString();
                                            userstemplist.Add(usertemp);
                                        }
                                    }
                                    #endregion
                                }
                            }
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            #endregion
            #region CleanUpExcel
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            #endregion
        }



        public string GetProjectName(string Projectname)
        {
            try
            {
                switch (Projectname)
                {
                    case "HD ATM":
                        return "Hardware ATM";
                    case "E2E":
                        return "E2E - ATM";
                    case "RTDS":
                        return "RTDS Automation";
                    case "SWAT":
                        return "SWAT";
                    default:
                        return "";
                }
            }
            catch (Exception ex)
            {
                return "";
            }

        }
        IWebDriver driver;
        public bool firstTime = true;
        private void UploadEmployeesDataToGTM()
        {
            string[] stringSeparators = new string[] { "\r\n" };
            try
            {
                if (driver != null)
                {
                    driver.Close();
                    driver.Quit();
                }
                driver = new ChromeDriver();
                driver.Navigate().GoToUrl(@"http://sjp2-cc/gtm/LoginPage.aspx");

                int d = 0;
                int user = 0;
                Thread.Sleep(2000);
                if (driver.Url.Contains("Login"))
                {
                    driver.FindElement(By.XPath("//*[@id=\"MiddleContents_TxtAccountId\"]")).SendKeys(userstemplist[user].UserName);
                    Thread.Sleep(2000);
                    driver.FindElement(By.XPath("//*[@id=\"MiddleContents_TxtPassword\"]")).SendKeys(userstemplist[user].Password);
                    driver.FindElement(By.XPath("//*[@id=\"MiddleContents_UserLogin\"]")).Click();
                    Thread.Sleep(2000);//*[@id="ctl00_MiddleContents_LnkUserLogin"]
                    Login = true;
                }
                WeeklyRecordTemp currentWeekRecordtemp = userstemplist[user].weekly_record;
                var gtmWeek = driver.FindElement(By.XPath(WEEK_XPATH)).Text.Split(stringSeparators, StringSplitOptions.None).ToList();

                //Pick up any dates year as it will remain same for complete month
                var resultantYear = currentWeekRecordtemp.DailyRecordList[0].RecordDate.Year.ToString();
                //Year
                driver.FindElement(By.XPath(YEAR_XPATH)).SendKeys(resultantYear);
                Thread.Sleep(MillisecondsTimeout);
                //Traverse through weeks daily records
                for (int day = 0; day < currentWeekRecordtemp.DailyRecordList.Count; day++)
                {
                    DailyRecordTemp dailyRecord = currentWeekRecordtemp.DailyRecordList[day];
                    //Week
                    var resultantWeek = gtmWeek.Find(x => x.Contains($"# {currentWeekRecordtemp.DailyRecordList[d].Week}"));
                    d++;
                    driver.FindElement(By.XPath(WEEK_XPATH)).SendKeys(resultantWeek);
                    Thread.Sleep(MillisecondsTimeout);
                    driver.FindElement(By.XPath(WEEK_XPATH)).Click();
                    Thread.Sleep(MillisecondsTimeout);

                    //Program name
                    driver.FindElement(By.XPath(PROGRAM_NAME_XPATH)).SendKeys(ProgramName);
                    Thread.Sleep(1000);

                    driver.FindElement(By.XPath(PROJECT_NAME_XPATH)).SendKeys(dailyRecord.Project);
                    Thread.Sleep(MillisecondsTimeout);
                    //date description hours and wbs
                    var resultantDate = dailyRecord.RecordDate.ToString("dd-MMM-yyyy");
                    var resultantHours = dailyRecord.Hours;
                    var resultantDescription = dailyRecord.Description;
                    //var resultantWBS = dailyRecord.WBS; //TO BE ADDED


                    //Date                     
                    driver.FindElement(By.XPath(DATE_XPATH)).SendKeys(resultantDate);
                    Thread.Sleep(MillisecondsTimeout);

                    //WBS
                    driver.FindElement(By.XPath(WBS_XPATH)).SendKeys("Testing");
                    Thread.Sleep(MillisecondsTimeout);

                    //Description
                    driver.FindElement(By.XPath(DESCRIPTION_XPATH)).Clear();
                    Thread.Sleep(MillisecondsTimeout);
                    driver.FindElement(By.XPath(DESCRIPTION_XPATH)).SendKeys(resultantDescription);

                    //Hours
                    Thread.Sleep(MillisecondsTimeout);
                    driver.FindElement(By.XPath(HOURS_XPATH)).SendKeys("-");
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath(HOURS_XPATH)).SendKeys(resultantHours);
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath(SAVEBUTTON_XPATH)).Click();
                }
                Thread.Sleep(1000);
                Login = false;
                Thread.Sleep(2000);

                this.Dispatcher.Invoke(() => WindowState = WindowState.Maximized);
                this.Dispatcher.Invoke(() => MessageBox.Show("Your data has successfully been uploaded on GTM portal"));

                #region CleanUpExcel
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                #endregion
            }
            catch (Exception ex)
            {
                #region CleanUpExcel
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                //Marshal.ReleaseComObject(xlRange);
                //Marshal.ReleaseComObject(xlWorksheet);

                ////close and release
                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                ////quit and release
                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
                MessageBox.Show(ex.Message);
                #endregion
                driver.Close();
                driver.Quit();
                MessageBox.Show(ex.Message);
            }

        }

        #endregion

        #region Private Methods      
        private void btnPauseLogging_Click(object sender, RoutedEventArgs e)
        {
            Pause();
        }

        private void btnResumeLogging_Click(object sender, RoutedEventArgs e)
        {
            Resume();
        }

        private void btnBrowseGtmFile_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog1;
            openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "txt files (*.xlsx)|*.xlsx";
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.CheckFileExists = true;

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                tbGtmFilePath.Text = openFileDialog1.FileName;
                cbWeeklyOrMonthly.SelectedIndex = 0;
            }
        }
        #endregion


        private void Combo_Monthly_SelectionChange(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                using (var stream = File.Open(tbGtmFilePath.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        comboBox_week.Items.Clear();
                        //comboBox_date.Items.Clear();

                        for (int i = 1; i < result.Tables[comboBox_Months.SelectedItem.ToString()].Rows.Count - 1; i++)
                        {

                            string emp = result.Tables[comboBox_Months.SelectedItem.ToString()].Rows[i][3].ToString();
                            string emp_pre = result.Tables[comboBox_Months.SelectedItem.ToString()].Rows[i - 1][3].ToString();

                            if (emp_pre != emp)
                            {
                                combo_employeename.Items.Add(result.Tables[comboBox_Months.SelectedItem.ToString()].Rows[i][3].ToString());
                            }
                            string date = result.Tables[comboBox_Months.SelectedItem.ToString()].Rows[i][0].ToString();
                            //if (!comboBox_date.Items.Contains(date))
                            //{
                            //    comboBox_date.Items.Add(date);
                            //}

                            string week = result.Tables[comboBox_Months.SelectedItem.ToString()].Rows[i][2].ToString();
                            if (week != "" && !(comboBox_week.Items.Contains(week)))
                            {
                                comboBox_week.Items.Add(week);

                            }

                        }

                    }
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }


        private void btn_closechrome_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (driver != null)
                {
                    driver.Close();
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }


        private void cbWeeklyOrMonthlyClick(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (cbWeeklyOrMonthly.SelectedItem.ToString().Contains("This month"))
            {

                labelWeek.Visibility = Visibility.Hidden;
                comboBox_week.Visibility = Visibility.Hidden;
                //labelDate.Visibility = Visibility.Visible;
                //comboBox_date.Visibility = Visibility.Visible;

            }
            else if (cbWeeklyOrMonthly.SelectedItem.ToString().Contains("Weekly"))
            {

                //labelDate.Visibility = Visibility.Hidden;
                //comboBox_date.Visibility = Visibility.Hidden;
                labelWeek.Visibility = Visibility.Visible;
                comboBox_week.Visibility = Visibility.Visible;


            }


        }

        private void tbGtmFilePathChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            try
            {
                string FilePath = tbGtmFilePath.Text;

                using (var stream = File.Open(FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        for (int i = 0; i < result.Tables.Count; i++)
                        {
                            string sheetName = result.Tables[i].TableName.ToString();
                            if (!comboBox_Months.Items.Contains(sheetName) &&
                               !sheetName.EndsWith("s") && sheetName != "Table" && sheetName != "Streams Mapping")
                            {
                                comboBox_Months.Items.Add(result.Tables[i].TableName.ToString());

                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }


    }
}


