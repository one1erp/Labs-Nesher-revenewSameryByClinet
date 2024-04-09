using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace revenewSameryByClinet
{
    public class ClientObj
    {
        public int ClientId { get; set; }
        public string ClientCode { get; set; }
        public string ClientName { get; set; }
        public string Adress { get; set; }
        public string Fhone { get; set; }
        public string Fax { get; set; }
        public string Email { get; set; }
        public string Remark { get; set; }
        public List<SdjObj> sdgs = new List<SdjObj>();
        public List<string> Columns = new List<string> {};
    }

    public class SdjObj 
    {
        public int SdjId { get; set; }
        public string SdjName { get; set; } 
        public string LabName { get; set; }       //שם מעבדה
        public string ExternalReference { get; set; }   //הזמנת לקוח
        public DateTime DeliveryDate { get; set; }      //תאריך
        public List<TestObj> tests = new List<TestObj>();
    }

    public class TestObj
    {
        public int TestId { get; set; }
        public string TestName { get; set; }
        public double Price { get; set; }
        public int CountTest { get; set; }

    }

    public class ClientContract
    {
        public int ClientId { get; set; }
        public List<TestPrice> tests = new List<TestPrice>();

    }
    public class TestPrice
    {
        public int TestId { get; set; }
        public double Price { get; set; }
    }
}
