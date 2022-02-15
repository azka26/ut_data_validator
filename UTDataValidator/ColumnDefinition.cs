using System.Collections.Generic;

namespace UTDataValidator
{
    public class ColumnDefinition
    {
        public ColumnDefinition(string columnValue, int columnIndex)
        {
            ColumnIndex = columnIndex;
            if (columnValue.Contains(":"))
            {
                string[] split = columnValue.Split(':');
                ColumnName = split[0].Trim();
                string needCheck = split[1].Trim();
                if (needCheck == "1")
                {
                    NeedValidation = true;
                }
            }
            else
            {
                ColumnName = columnValue;
                NeedValidation = false;
            }
        }

        public string ColumnName { get; private set; }
        public bool NeedValidation { get; private set; }
        public int ColumnIndex { get; set; }
    }
}
