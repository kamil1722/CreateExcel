using System;
using System.ComponentModel.DataAnnotations;

namespace MyExel.Models
{
    public class ErrorViewModel
    {
        public string RequestId { get; set; }

        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
    }
    public class med
    {
        [Key]
        public int mo { get; set; }
        public int cont { get; set; }
        public string date_1 { get; set; }
        public string date_2 { get; set; }
        public string ds1 { get; set; }

    }
    public class mkb
    {
        [Key]
        public string code { get; set; }
        public string name { get; set; }
        public string dbegin { get; set; }
        public string dend { get; set; }
    }
    public class mo
    {
        [Key]
        public int code { get; set; }
        public string name { get; set; }
        public string dbegin { get; set; }
        public string dend { get; set; }
    }
    public class smo
    {
        [Key]
        public int code { get; set; }
        public string name { get; set; }
        public string dbegin { get; set; }
        public string dend { get; set; }
    }
    public class ExelViewModel
    {
        public int Number { get; set; }
        public string Code { get; set; }
        public int CodeSMO { get; set; }
        public string NameSMO { get; set; }
        public int CodeMO { get; set; }
        public string NameMO { get; set; }
        public string CodeMKB { get; set; }
        public string NameMKB { get; set; }
        public int Count { get; set; }
        //public int CountIncident { get; set; }
    }
}
