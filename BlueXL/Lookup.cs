using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

namespace BlueXL
{
    public static class LookupFunctions
    {
        public static bool IsNumber(this object value)
        {
            return value is sbyte
                    || value is byte
                    || value is short
                    || value is ushort
                    || value is int
                    || value is uint
                    || value is long
                    || value is ulong
                    || value is float
                    || value is double
                    || value is decimal;
        }

        [ExcelFunction(Description = "Print regex matches")]
        public static object BXLookup(string tolookup, object[,] data, int lookupCol, object[,] columns)
        {
            Regex rgx = new Regex(tolookup);
            
            var colsToPrint = new List<long>();
            foreach (object col in columns)
            {
                if (IsNumber(col))
                {
                    double dCol = Convert.ToDouble(col);
                    if ((dCol%1) < Double.Epsilon)
                    {
                        int iCol = Convert.ToInt32(dCol);
                        colsToPrint.Add(iCol);
                    }
                }
            }

            var matchRows = new List<long>();
            for (long row = 0; row < data.GetLongLength(0); ++row)
            {
                if (rgx.Match(data[row, lookupCol].ToString()).Success)
                {
                    matchRows.Add(row);
                }
            }
            var toRet = new object[matchRows.Count, colsToPrint.Count];
            for (int matchRowIdx = 0 ; matchRowIdx < matchRows.Count ; ++matchRowIdx)
            {
                for (int colsToPrintIdx = 0; colsToPrintIdx < colsToPrint.Count; ++colsToPrintIdx)
                {
                    toRet[matchRowIdx, colsToPrintIdx] = data[matchRows[matchRowIdx], colsToPrint[colsToPrintIdx]];
                }
            }
            return toRet;
        }

        [ExcelFunction(Description = "Print regex matches")]
        public static object BXRegexMatch(string input, string pattern)
        {
            Regex rgx = new Regex(pattern);
            MatchCollection matches = rgx.Matches(input);

            var groupRow = new Dictionary<string, int>();
            int nbCols = rgx.GetGroupNames().Count();

            var toRet = new object[nbCols, (nbCols == 1 ? 0 : 1) + matches.Count];
            for (int row = 0; row < toRet.GetLength(0); ++row)
            {
                for (int col = 0; col < toRet.GetLength(1); ++col)
                {
                    toRet[row, col] = "";
                }
            }
            if (nbCols > 1)
            {
                for (var groupIdx = 1; groupIdx < rgx.GetGroupNames().Length; ++groupIdx)
                {
                    toRet[groupIdx, 0] = rgx.GetGroupNames()[groupIdx];
                }
            }
            for (var matchIdx = 0 ; matchIdx < matches.Count ; ++matchIdx)
            {
                var match = matches[matchIdx];
                for (var groupIdx = 0; groupIdx < match.Groups.Count; ++groupIdx)
                {
                    toRet[groupIdx, (nbCols == 1 ? 0 : 1)] = match.Groups[groupIdx].Value;
                }
            }
            return toRet;
        }
    }
}
