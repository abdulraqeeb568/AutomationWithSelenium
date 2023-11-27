using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace GTMFiller
{
    public class Worker
    {
        public const int ROW_EMPLOYEE_COUNT = 0;
        public const int ROW_NAME = 0;
        public const int ROW_USERNAME = ROW_NAME + 1;
        public const int ROW_PASSWORD = ROW_USERNAME + 1;
        public const int START_ROW_RECORD = 4;
        public const int ROWS_TAKEN_BY_WEEK = 8;
        public const int ROWS_WEEK_COMPLETE = 11;
        public const int ROW_END_RECORD = 43;

        public const int COLUMN_EMPLOYEE_COUNT = 1;
        public const int COLUMN_RECORD_WEEK = 4;
        public const int COLUMN_RECORD_DATE = 6;
        public const int START_COLUMN_USER = 8;
        public const int COLUMNS_TAKEN_BY_USER = 3;

        public static List<MonthlyRecord> listOfEmployeesMonthlyRecords = new List<MonthlyRecord>();

        public static string excelPath = @"J:\SmartWiresATM\Smart Wires GTM.xlsm";

        #region Parse Excel To Employee(Old Implementation)
        //public static Status ParseExcelToEmployees()
        //{
        //    try
        //    {
        //        using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
        //        {
        //            using (var reader = ExcelReaderFactory.CreateReader(stream))
        //            {
        //                var result = reader.AsDataSet();
        //                //string sheetName = $"{DateTime.Now.ToString("MMMM")} {DateTime.Now.Year}";
        //                string sheetName = "Feb2021";

        //                int employeeCount = Convert.ToInt32(result.Tables[sheetName].Rows[ROW_EMPLOYEE_COUNT][COLUMN_EMPLOYEE_COUNT]);

        //                for (int employee = 0; employee < employeeCount; employee++)
        //                {
        //                    //If completed do not populate employee
        //                    if (Convert.ToString(result.Tables[sheetName].Rows[ROW_END_RECORD][START_COLUMN_USER + 1 + employee * COLUMNS_TAKEN_BY_USER]) == "c")
        //                    {
        //                        continue;
        //                    }

        //                    UserTemp usertemp = new UserTemp();
        //                    usertemp.Name = Convert.ToString(result.Tables[sheetName].Rows[ROW_NAME][START_COLUMN_USER + employee * COLUMNS_TAKEN_BY_USER]);
        //                    usertemp.UserName = Convert.ToString(result.Tables[sheetName].Rows[ROW_USERNAME][START_COLUMN_USER + employee * COLUMNS_TAKEN_BY_USER]);
        //                    usertemp.Password = Convert.ToString(result.Tables[sheetName].Rows[ROW_PASSWORD][START_COLUMN_USER + employee * COLUMNS_TAKEN_BY_USER]);
        //                    usertemp.DescriptionColumnNumber = START_COLUMN_USER + 1 + employee * COLUMNS_TAKEN_BY_USER;
        //                    List<WeeklyRecordTemp> listOfWeeklyRecordstemp = new List<WeeklyRecordTemp>();

        //                    //A Month can't have more than 5 records
        //                    for (int week = 0; week < 5; week++)
        //                    {
        //                        if (Convert.ToString(result.Tables[sheetName].Rows[ROWS_WEEK_COMPLETE + 8 * week][START_COLUMN_USER + 1 + employee * COLUMNS_TAKEN_BY_USER]) == "wc")
        //                            continue;

        //                        List<DailyRecordTemp> listOfDailyRecordstemp = new List<DailyRecordTemp>();

        //                        //A week can't have more than 7 working days
        //                        for (int day = 0; day < 7; day++)
        //                        {
        //                            DailyRecordTemp dailyRecord = new DailyRecordTemp();
        //                            var description = Convert.ToString(result.Tables[sheetName].
        //                                Rows[START_ROW_RECORD + week * ROWS_TAKEN_BY_WEEK + day][START_COLUMN_USER + 1 + employee * COLUMNS_TAKEN_BY_USER]);

        //                            if (description != "")
        //                            {
        //                                dailyRecord.RecordDate = Convert.ToDateTime(result.Tables[sheetName].
        //                                    Rows[START_ROW_RECORD + week * ROWS_TAKEN_BY_WEEK + day][COLUMN_RECORD_DATE]);

        //                                dailyRecord.Description = Convert.ToString(result.Tables[sheetName].
        //                                    Rows[START_ROW_RECORD + week * ROWS_TAKEN_BY_WEEK + day][START_COLUMN_USER + 1 + employee * COLUMNS_TAKEN_BY_USER]);
        //                                dailyRecord.Hours = Convert.ToString(result.Tables[sheetName].
        //                                    Rows[START_ROW_RECORD + week * ROWS_TAKEN_BY_WEEK + day][START_COLUMN_USER + 2 + employee * COLUMNS_TAKEN_BY_USER]);

        //                                listOfDailyRecordstemp.Add(dailyRecord);
        //                            }

        //                        }

        //                        if (listOfDailyRecordstemp.Count != 0)
        //                        {
        //                            WeeklyRecordTemp weeklyRecordtemp = new WeeklyRecordTemp();

        //                            weeklyRecordtemp.Week = Convert.ToString(result.Tables[sheetName].
        //                                    Rows[START_ROW_RECORD + week * ROWS_TAKEN_BY_WEEK][COLUMN_RECORD_WEEK]);
        //                            weeklyRecordtemp.Project = Convert.ToString(result.Tables[sheetName].
        //                                Rows[START_ROW_RECORD + week * ROWS_TAKEN_BY_WEEK][START_COLUMN_USER + employee * COLUMNS_TAKEN_BY_USER]);

        //                            weeklyRecordtemp.DailyRecordList = listOfDailyRecordstemp;

        //                            listOfWeeklyRecordstemp.Add(weeklyRecordtemp);
        //                        }
        //                    }

        //                    MonthlyRecord monthlyRecord = new MonthlyRecord();

        //                    monthlyRecord.Usertemp = usertemp;
        //                    monthlyRecord.WeeklyRecordsList = listOfWeeklyRecordstemp;

        //                    listOfEmployeesMonthlyRecords.Add(monthlyRecord);

        //                }
        //            }
        //        }

        //        return new Status(false);
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.WriteLine(LogType.Exception, $"{MethodBase.GetCurrentMethod().DeclaringType}|{ MethodBase.GetCurrentMethod().Name } {ex.StackTrace}");
        //        return new Status(true, $"Exception in + {MethodBase.GetCurrentMethod().DeclaringType}|{ MethodBase.GetCurrentMethod().Name }");

        //    }

        //}
        #endregion
    }
}
