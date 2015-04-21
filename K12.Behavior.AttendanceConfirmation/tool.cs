using Aspose.Words;
using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Behavior.AttendanceConfirmation
{
    static public class tool
    {
        static public Run _run;

        static public double FontSize = 10;
        static public string FontName = "標楷體";

        static public List<string> GetPeriod()
        {
            //節次對照表
            List<string> list = new List<string>();

            //取得 Period List

            foreach (K12.Data.PeriodMappingInfo each in K12.Data.PeriodMapping.SelectAll())
            {
                if (!list.Contains(each.Name))
                    list.Add(each.Name);

            }

            return list;
        }

        /// <summary>
        /// Cell切割器
        /// </summary>
        /// <param name="_cell">傳入分割的儲存格</param>
        /// <param name="Count">傳入分割數目</param>
        static public void CellSplit(Cell _cell, int Count)
        {
            #region Cell切割器
            double MAXwidth = _cell.CellFormat.Width;
            double Cellwidth = MAXwidth / Count;

            CellFormat cf = _cell.CellFormat;
            cf.Width = Cellwidth;

            List<Cell> list = new List<Cell>();
            list.Add(_cell);

            Row _row = _cell.ParentNode as Row;


            for (int x = 0; x < Count - 1; x++)
            {
                list.Add((_row.InsertAfter(new Cell(_cell.Document), _cell)) as Cell);
            }

            foreach (Cell each in list)
            {
                each.CellFormat.Width = Cellwidth;
            }
            #endregion
        }

        /// <summary>
        /// 以Cell為基準,使用NextSibling向右移一格
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        static public Cell GetMoveRightCellByNextSibling(Cell cell, int count)
        {
            #region 以Cell為基準,使用NextSibling向右移一格
            if (count == 0) return cell;

            Node node = cell;
            for (int i = 0; i < count; i++)
                node = node.NextSibling;

            try
            {
                return (Cell)node;
            }
            catch (Exception ex)
            {
                return null;
            }
            #endregion
        }


        /// <summary>
        /// 以Cell為基準,向右移一格
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        static public Cell GetMoveRightCell(Cell cell, int count)
        {
            #region 以Cell為基準,向右移一格
            if (count == 0) return cell;

            Row row = cell.ParentRow;
            int col_index = row.IndexOf(cell);
            Table table = row.ParentTable;
            int row_index = table.Rows.IndexOf(row);

            try
            {
                return table.Rows[row_index].Cells[col_index + count];
            }
            catch (Exception ex)
            {
                return null;
            }
            #endregion
        }

        /// <summary>
        /// 寫入資料
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        static public void Write(Cell cell, string text)
        {
            #region 寫入資料
            if (cell.FirstParagraph == null)
                cell.Paragraphs.Add(new Paragraph(cell.Document));
            cell.FirstParagraph.Runs.Clear();
            _run.Text = text;
            _run.Font.Size = FontSize;
            _run.Font.Name = FontName;
            cell.FirstParagraph.Runs.Add(_run.Clone(true));
            #endregion
        }

        static public int sortdat(string a, string b)
        {
            DateTime dt1 = DateTime.Now;
            string aa = a.Remove(a.IndexOf('('));
            DateTime.TryParse(aa, out dt1);

            DateTime dt2 = DateTime.Now;
            string bb = b.Remove(b.IndexOf('('));
            DateTime.TryParse(bb, out dt2);

            return dt1.CompareTo(dt2);
        }

        public static int StudentComparer(StudentRecord x, StudentRecord y)
        {
            int xx = x.SeatNo.HasValue ? x.SeatNo.Value : 0;
            int yy = y.SeatNo.HasValue ? y.SeatNo.Value : 0;

            return xx.CompareTo(yy);
        }

        static public string HowManyWeek(DateTime OccurDate)
        {
            string stringDate = OccurDate.ToShortDateString();
            switch (OccurDate.DayOfWeek.ToString())
            {
                case "Monday":
                    stringDate += "(一)";
                    break;
                case "Tuesday":
                    stringDate += "(二)";
                    break;
                case "Wednesday":
                    stringDate += "(三)";
                    break;
                case "Thursday":
                    stringDate += "(四)";
                    break;
                case "Friday":
                    stringDate += "(五)";
                    break;
                case "Saturday":
                    stringDate += "(六)";
                    break;
                case "Sunday":
                    stringDate += "(日)";
                    break;
            }
            return stringDate;
        }
    }
}
