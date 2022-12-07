using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQLiteApp
{
    public class User
    {
        public int PK { get; set; }
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? Family { get; set; }
        public string? Patronymicl { get; set; }
        public string? Telephone { get; set; }
        public string? Birthday { get; set; }
        public string? Department { get; set; }
    }
}
