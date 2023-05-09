using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLayer
{
    public class Change
    {
        public int RowIndex { get; set; }
        public string? ColumnName { get; set; }
        public object OldValue { get; set; }
        public object NewValue { get; set; }
    }

}
