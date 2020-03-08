using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace taction.DTO
{
    public class DataDTO
    {
        public Tenant tenant { get; set; }
        public Issue issue { get; set; }
        public IEnumerable<History> history { get; set; }
        public Landlord landlord { get; set; }
        public IEnumerable<Availability> availability { get; set; }
        public IEnumerable<Defect> defects { get; set; }
    }

    public class Availability {
        public DateTime startAt { get; set; }
        public DateTime endAt { get; set; }
    }

    public class Tenant
    {
        public string name { get; set; }
        public string address1 { get; set; }
        public string address2 { get; set; }
        public string address3 { get; set; }
        public string postcode { get; set; }
        public string city { get; set; }
    }

    public class Landlord
    {
        public string name { get; set; }
        public string address1 { get; set; }
        public string address2 { get; set; }
        public string address3 { get; set; }
        public string city { get; set; }
        public string postcode { get; set; }
    }

    public class History
    {
        public DateTime date { get; set; }
        public string description { get; set; }
    }

    public class Defect
    {
        public string item { get; set; }
        public string room { get; set; }
        public string extentOfDamage { get; set; }
        public string inconvenienceSuffered { get; set; }
        public string costOfRepair { get; set; }
        public string photo { get; set; }
    }

    public class Issue
    {
        public string summary { get; set; }
        public string effects { get; set; }
    }
}
