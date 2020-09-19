using System;
using System.Collections.Generic;
using System.Data;
using System.Text;


/// <summary>
/// {      "SiteType":"Any","ActivityType":"Any","DiskUsed":17790350,"DiskQuota":0,"DocumentCount":9,"TotalSites":11,"ActivityTotalSites":0,
///         "SitesWithOwnerActivities":0,"SitesWithNonOwnerActivities":0,"TimeFrame":"2020-07","ContentDate":"2020-07-24T00:00:00Z"    },
/// </summary>
namespace ReportsData
{
    public class TenantSharePointUsageModel
    {
        public string SiteType { get; set; }
        public string ActivityType { get; set; }
        public Int64 DiskUsed { get; set; }
        public Int64 DiskQuota { get; set; }
        public int DocumentCount { get; set; }
        public int TotalSites { get; set; }
        public int ActivityTotalSites { get; set; }
        public int SitesWithOwnerActivities { get; set; }
        public int SitesWithNonOwnerActivities { get; set; }
        public string TimeFrame { get; set; }
        public DateTime ContentDate { get; set; }

    }

    public class WeatherForecast
    {
        public DateTimeOffset Date { get; set; }
        public int TemperatureCelsius { get; set; }
        public string Summary { get; set; }
    }
}
