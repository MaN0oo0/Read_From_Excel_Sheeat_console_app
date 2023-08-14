using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_From_Excel_Sheeat_console_app.Models
{
    public record ModelRes
    {
        public int seating_no { get; set; }
        public string arabic_name { get; set; }
        public double total_degree { get; set; }
        public int student_case { get; set; }
        public string student_case_desc { get; set; }
        public int c_flage { get; set;}
    }
}
