using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace skype_test_proj
{
    static class Data
    {
        public static int my_edited_count { get; set; }
        public static int my_filetransfer_count { get; set; }
        public static Int64 my_symbols_count { get; set; }
        public static int my_messages_count { get; set; }
        

        public static int comp_edited_count { get; set; }
        public static int comp_filetransfer_count { get; set; }
        public static Int64 comp_symbols_count { get; set; }
        public static int comp_messages_count { get; set; }
        public static string comp_name { get; set; }

        public static DateTime first_msg { get; set; }
        public static DateTime last_msg { get; set; }

        public static int global_I { get; set; }
        public static int global_Skype_smiles { get; set; }

    }
}
