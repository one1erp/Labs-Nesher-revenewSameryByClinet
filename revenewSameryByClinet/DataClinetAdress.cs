using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using DAL;

namespace revenewSameryByClinet
{
    public class DataClinetAdress
    {
        public DataTable DataTable { get; set; }
        public Client Client { get; set; }
        public Address Address { get; set; }

        public string Remark { get; set; }
    }
}
