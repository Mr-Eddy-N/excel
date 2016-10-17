using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;

namespace TEC.VacacionesColaboradoresWeb
{
    public class Logger
    {
        private static Logger instance;
        private Queue<Log> logQueue = new Queue<Log>();
        private string logFilename = "";
        private int queueSize = 1;
        private DateTime lastFlushed;
        private double maxLogAge = 60; //Segundos

        private Logger(){
           //string location= Directory.GetCurrentDirectory().ToString();
            string filePath = ConfigurationManager.AppSettings["PATHSLN"];
            logFilename = filePath+@"\log_" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + ".txt";
           // Console.WriteLine(logFilename);
        }

        ~Logger()
        {
            FlushLog();
        }

        public static Logger getInstance()
        {
            if (instance == null)
            {
                instance = new Logger();
            }
            return instance;
        }

        public void WriteLog(string message)
        {
            bool v = Convert.ToBoolean(ConfigurationManager.AppSettings["logger"]);
            if (v)
            {
                lock (logQueue)
                {
                    logQueue.Enqueue(new Log(DateTime.Now, message));
                    if (canFlush())
                    {
                        FlushLog();
                    }
                }
            }
        }
        public void WriteLog(Exception ex)
        {
            lock (logQueue)
            {
                logQueue.Enqueue(new Log(DateTime.Now, ex.Message +"\nTRACE:"+ex.StackTrace));
                if (canFlush())
                {
                    FlushLog();
                }
            }
        }

        private bool canFlush()
        {
            TimeSpan logAge = DateTime.Now - lastFlushed;
            if (logQueue.Count >= queueSize || logAge.TotalSeconds >= maxLogAge)
            {
                lastFlushed = DateTime.Now;
                return true;
            }

            return false;
        }

        private void FlushLog()
        {
            while (logQueue.Count > 0)
            {
                Log log = logQueue.Dequeue();

                //using (FileStream fs = File.Open(@logFilename, FileMode.Append, FileAccess.Write))
                {
                    //using (StreamWriter sw = new StreamWriter(fs))
                    using(StreamWriter sw = File.AppendText(@logFilename))
                    {
                        sw.WriteLine(string.Format("{0}\t{1}", log.logTime, log.message));
                    }
                }
            }
        }

        public class Log
        {
            public DateTime logTime { get; set; }
            public string message { get; set; }

            public Log(DateTime logTime, string message)
            {
                this.logTime = logTime;
                this.message = message;
            }
        }
    }
}