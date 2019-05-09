using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public class ExcelRow : IEquatable<ExcelRow>
    {
        public int Index { get; private set; }
        public List<ExcelCell> Cells { get; private set; }

        public bool KeyCompare { get; private set; }
        public ExcelRow(int index, IEnumerable<ExcelCell> cells, bool keyCompare=false)
        {
            KeyCompare = keyCompare;
            Index = index;
            Cells = cells.ToList();
        }

        public override bool Equals(object obj)
        {
            if (KeyCompare)
            {
                return KeyEqual(obj);
            }
            else
            {
                var other = obj as ExcelRow;

                return Equals(other);
            }
        }

        public bool KeyEqual(object obj)
        {
            var other = obj as ExcelRow;
            return Cells[0].GetHashCode() == other.Cells[0].GetHashCode();
        }
        public override int GetHashCode()
        {
            var hash = 7;
            foreach (var cell in Cells)
            {
                hash = hash * 13 + cell.Value.GetHashCode();
            }

            return hash;
        }

        public bool Equals(ExcelRow other)
        {
            if (other == null)
                return false;

            return GetHashCode() == other.GetHashCode();
        }

        public bool IsBlank()
        {
            return Cells.All(c => string.IsNullOrEmpty(c.Value));
        }

        public void UpdateCells(IEnumerable<ExcelCell> cells)
        {
            Cells = cells.ToList();
        }
    }

    internal class RowComparer : IEqualityComparer<ExcelRow>
    {
        public HashSet<int> IgnoreColumns { get; private set; }

        public RowComparer(HashSet<int> ignoreColumns)
        {
            IgnoreColumns = ignoreColumns;
        }

        public bool Equals(ExcelRow x, ExcelRow y)
        {
            return GetHashCode(x).Equals(GetHashCode(y));
        }

        public int GetHashCode(ExcelRow obj)
        {
            //改为只判断第一个cell
            var hash = 7;
            var index = 0;
            foreach (var cell in obj.Cells)
            {
                if (IgnoreColumns.Contains(index))
                    continue;

                hash = hash * 13 + cell.Value.GetHashCode();

                index++;
            }

            return hash;            
        }
    }

    internal class RowKeyComparer : IEqualityComparer<ExcelRow>
    {
        public HashSet<int> IgnoreColumns { get; private set; }
        public RowKeyComparer(HashSet<int> ignoreColumns) 
        {
            IgnoreColumns = ignoreColumns;
        }
        public bool Equals(ExcelRow x, ExcelRow y)
        {
            string xHashCode = "-1";
            string yHashCode = "-1";
            if (x.Cells.Count>0)
                xHashCode= x.Cells[0].Value.ToString();
            if(y.Cells.Count>0)
                yHashCode = y.Cells[0].Value.ToString();
            return xHashCode == yHashCode;
        }
        public int GetHashCode(ExcelRow obj)
        {
            if (obj.Cells.Count > 0)
                return obj.Cells[0].GetHashCode();
            else
                return -1;
        }
    }

}
