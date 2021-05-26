using System;
using System.Collections.Generic;
using System.Text;

namespace MyKoloDAL.Models
{
    public class Expense
    {
        public int Id { get; set; }
        public DateTime CreatedDateTime { get; set; }
        public double Amount { get; set; }
        public string Description { get; set; }
    }
}
