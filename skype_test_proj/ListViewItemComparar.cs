using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Globalization;

namespace skype_test_proj
{
    class ListViewItemComparer : IComparer
    {
        private int col, por;
        public ListViewItemComparer()
        {
            col = 0;
            por = 1;
        }
        public ListViewItemComparer(int column, int sort_type)
        {
            col = column;
            por = sort_type;
        }
        public int Compare(object x, object y)
        {
            ListViewItem lvx = (ListViewItem)x;
            ListViewItem lvy = (ListViewItem)y;
            if (col == 0)
                return por * String.Compare(lvx.SubItems[col].Text, lvy.SubItems[col].Text, StringComparison.OrdinalIgnoreCase);
            else
                return -por * (int.Parse(lvx.SubItems[col].Text)).CompareTo(int.Parse(lvy.SubItems[col].Text)); 
        }
    }

}
