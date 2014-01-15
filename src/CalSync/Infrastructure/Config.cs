using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace CalSync.Infrastructure
{
    class Config
    {
        public bool ValidConfiguration { get; set; }
        public int SyncRangeDays { get; set; }
        public string TargetEmailAddress { get; set; }

        public static Config Read()
        {
            try
            {
                return new Config()
                {
                    SyncRangeDays = int.Parse(ConfigurationManager.AppSettings["SyncRangeDays"]),
                    TargetEmailAddress = ConfigurationManager.AppSettings["TargetEmailAddress"],
                    ValidConfiguration = true
                };
            } 
            catch(Exception e)
            {
                if(e is ConfigurationErrorsException || e is FormatException || e is OverflowException )
                {
                    return new Config()
                    {
                        ValidConfiguration = false
                    };
                }
                throw;
            }
            
        }
    }
}
