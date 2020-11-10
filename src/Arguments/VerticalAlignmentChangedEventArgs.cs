using NPOI.SS.UserModel;
using System;

namespace ExportToExcel
{
    public class VerticalAlignmentChangedEventArgs : EventArgs
    {
        public VerticalAlignment SelectedAlignment { get; }

        public VerticalAlignmentChangedEventArgs(VerticalAlignment args)
        {
            SelectedAlignment = args;
        }
    }
}
