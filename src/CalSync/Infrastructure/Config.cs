using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace CalSync.Infrastructure
{
    class Config
    {
        public int SyncRangeDays { get; private set; }
        public string TargetEmailAddress { get; private set; }
        public bool Receive { get; private set; }
        public bool Send { get; private set; }
        public String ErrorMessage { get; private set; }

        public static Config Read()
        {
            try
            {
                return new Config()
                {
                    SyncRangeDays = int.Parse(ConfigurationManager.AppSettings["SyncRangeDays"]),
                    TargetEmailAddress = ConfigurationManager.AppSettings["TargetEmailAddress"],
                    Receive = bool.Parse(ConfigurationManager.AppSettings["Receive"] ?? "true"),
                    Send = bool.Parse(ConfigurationManager.AppSettings["Send"] ?? "true"),
                };
            } 
            catch(Exception e)
            {
                if(e is ConfigurationErrorsException || e is FormatException || e is OverflowException )
                {
                    return new Config()
                    {
                        ErrorMessage = e.Message
                    };
                }
                throw;
            }
            
        }
    }
}
