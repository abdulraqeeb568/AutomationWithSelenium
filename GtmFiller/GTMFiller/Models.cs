using System;
using System.Collections.Generic;
using System.IO;

namespace GTMFiller
{

    public class MonthlyRecord
    {
        public List<DailyRecordTemp> SingleRecords { get; set; }    //TO BE REMOVED

        public UserTemp Usertemp { get; set; }
        public List<WeeklyRecordTemp> WeeklyRecordsList { get; set; }
    }
    public class DataGridData
    {
        public string Name { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public DateTime RecordDate { get; set; }
        public string Week { get; set; }        //TO BE REMOVED
        public string Project { get; set; }     //TO BE REMOVED
        public string Description { get; set; }
        public string Hours { get; set; }

      

    }

    public class WeeklyRecordTemp
    {
        //public string Project { get; set; }
        public string Week { get; set; }
        public List<DailyRecordTemp> DailyRecordList { get; set; }
    }
    public class DailyRecordTemp
    {
        public DateTime RecordDate { get; set; }
        public string Project { get; set; }     //TO BE REMOVED
        public string Description { get; set; }
        public string Hours { get; set; }
        public string Week { get; set; }

    }



    public class UserTemp
    {
        public string Name { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public int DescriptionColumnNumber { get; set; }
        public List<WeeklyRecordTemp> weekly_record_list { get; set; }

        public WeeklyRecordTemp weekly_record { get; set; }
    }

    public class Status
    {
        public string Message { get; set; } = string.Empty;

        public bool ErrorOccured { get; set; } = false;

        public Status()
        {

        }

        public Status(bool errorOccured)
        {
            ErrorOccured = errorOccured;
        }

        public Status(bool errorOccured, string errorMessage)
        {
            ErrorOccured = errorOccured;
            Message = errorMessage;
        }
    }

    enum LogType
    {
        Error,
        Status,
        Event,
        Exception
    }

    static class Log
    {
        static DateTime dateTime = DateTime.Now;
        public static void WriteLine(LogType type, string message)
        {
            string path = Path.Combine(Environment.CurrentDirectory, $"Log_{ dateTime.ToString("yy_mm_dd_hh_mm_ss")}.txt");
            message = $"{DateTime.Now} \t\t {type.ToString()} \t {message} \n";
            File.AppendAllText(path, message);
        }
    }
}
