using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Text.RegularExpressions;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;

namespace ExcelHtmlAddin
{
    partial class Ribbon1
    {
        private string prefix = "<!-- ExcelHtmlAddin code start -->";
        private string sufix = "<!-- ExcelHtmlAddin code end -->";
        private double[] col_width_arr = { 9.13, 10.25, 10.25, 43.25, 30.75, 21.5 };

        //列幅調整（6列決め打ち）
        private void do_fit_col_width()
        {
            Excel.Range sa = Globals.ThisAddIn.Application.Selection;
            Excel.Worksheet ash = Globals.ThisAddIn.Application.ActiveSheet;
            int r, c1, c2 = 0;
            r = sa.Row;
            c1 = sa.Column;
            c2 = sa.Columns[sa.Columns.Count].Column;
            int cnt = 0;
            for(int j=c1; j<=c2; j++)
            {
                ash.Cells[r, j].ColumnWidth = col_width_arr[cnt];
                cnt++;
            }
            MessageBox.Show("列幅調整が完了しました。");
        }

        //表コードを表示
        private void do_table_tag_create_view()
        {
            string html = get_create_table_tag();
            if (PrevFormObj.Visible == false) PrevFormObj.Show();
            PrevFormObj.reportText.Text = html;
            //PrevFormObj.WindowState = FormWindowState.Normal;
            //PrevFormObj.Activate();
        }

        //表コードを保存
        private void do_table_tag_create_save()
        {
            string html = get_create_table_tag();
            string save_path = get_txt_save_path();
            System.Text.Encoding enc = new System.Text.UTF8Encoding(false);
            StreamWriter sw = new StreamWriter(save_path, false, enc);
            sw.Write(html);
            sw.Close();
        }

        //Excelの表組みからtable要素のhtmlソースを生成
        private string get_create_table_tag()
        {
            Excel.Range sa = Globals.ThisAddIn.Application.Selection;
            Excel.Worksheet ash = Globals.ThisAddIn.Application.ActiveSheet;

            int r1, r2, c1, c2 = 0;
            r1 = sa.Row;
            r2 = sa.Rows[sa.Rows.Count].Row;
            c1 = sa.Column;
            c2 = sa.Columns[sa.Columns.Count].Column;

            //列幅処理
            List<double> col_width_list = new List<double>();
            List<string> col_width_attrs = new List<string>();
            double col_total = 0;

            for (int x = c1; x <= c2; x++)
            {
                Excel.Range cell = ash.Cells[r1, x];
                double cw = cell.ColumnWidth;
                col_width_list.Add(cw);
                col_total += cw;
            }

            foreach (double cw in col_width_list)
            {
                double cpcw = (cw / col_total) * 100;
                string perc = Math.Floor(cpcw).ToString() + "%";
                col_width_attrs.Add(perc);
            }


            string html = "";
            html += @"<table class=""table-bordered"">" + "\r\n";


            //行ループ
            for (int i = r1; i <= r2; i++)
            {
                html += "<tr>\r\n";


                //列ループ
                for (int j = c1; j <= c2; j++)
                {
                    Excel.Range cell = ash.Cells[i, j];
                    int color_code = cell.Interior.ColorIndex;
                    Boolean align_flg = (cell.HorizontalAlignment == -4108) ? true : false;
                    Boolean border_top_none_flg = (cell.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle == -4142) ? true : false;
                    Boolean border_bottom_none_flg = (cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle == -4142) ? true : false;

                    //セル値を取得
                    string cell_val = "";
                    if (cell.Value == null)
                    {
                        cell_val = "";
                    }
                    else
                    {
                        Type t = cell.Value.GetType();
                        if(t.Equals(typeof(double)))
                        {
                            cell_val = cell.Value.ToString();
                        }
                        else if (t.Equals(typeof(DateTime)))
                        {
                            cell_val = cell.Value.ToString();
                        }
                        else if(t.Equals(typeof(string)))
                        {
                            cell_val = (string)cell.Value;
                        }
                    }
                    if(cell_val != "") cell_val = br_encode(cell_val);

                    //色無しセルはデータセル
                    if (color_code == -4142)
                    {
                        //セル結合判定
                        if (cell.MergeCells)
                        {
                            string merge_addr = cell.MergeArea.Address;
                            string scope = get_merge_scope(merge_addr);
                            if (scope.Equals("row"))
                            {
                                if (is_relative_cell(merge_addr, cell.Address))
                                {
                                    int merge_cnt = get_merge_count_row(merge_addr);
                                    html += @"<td rowspan=""" + merge_cnt + @"""";
                                    if (align_flg) html += @" style=""text-align: center;""";
                                    html += ">" + cell_val + "</td>\r\n";
                                }

                            }
                            else if (scope.Equals("col"))
                            {
                                if (is_relative_cell(merge_addr, cell.Address))
                                {
                                    int merge_cnt = get_merge_count_col(merge_addr);
                                    html += @"<td colspan=""" + merge_cnt + @"""";
                                    if (align_flg) html += @" style=""text-align: center;""";
                                    html += ">" + cell_val + "</td>\r\n";
                                }
                            }
                            else if (scope.Equals("cross"))
                            {
                                if (is_relative_cell(merge_addr, cell.Address))
                                {
                                    int merge_row_cnt = get_merge_count_row(merge_addr);
                                    int merge_col_cnt = get_merge_count_col(merge_addr);
                                    html += @"<td rowspan=""" + merge_row_cnt + @""" colspan=""" + merge_col_cnt + @"""";
                                    if (align_flg) html += @" style=""text-align: center;""";
                                    html += ">" + cell_val + "</td>\r\n";
                                }
                            }
                        }
                        else
                        {
                            html += "<td";
                            if (align_flg == true || border_top_none_flg == true || border_bottom_none_flg == true) html += @" style=""";
                            if (align_flg) html += "text-align: center;";
                            if(border_top_none_flg) html += @"border-top: none;";
                            if(border_bottom_none_flg) html += @"border-bottom: none;";
                            if (align_flg == true || border_top_none_flg == true || border_bottom_none_flg == true) html += @"""";

                            html += ">" + cell_val + "</td>\r\n";
                        }

                    }
                    //色つきセルは見出しセル
                    else
                    {
                        //セル結合判定
                        if (cell.MergeCells)
                        {
                            string merge_addr = cell.MergeArea.Address;
                            string scope = get_merge_scope(merge_addr);
                            if (scope.Equals("row"))
                            {
                                if (is_relative_cell(merge_addr, cell.Address))
                                {
                                    int merge_cnt = get_merge_count_row(merge_addr);
                                    html += @"<th rowspan=""" + merge_cnt + @"""";
                                    if (align_flg) html += @" style=""text-align: center;""";
                                    html += ">" + cell_val + "</th>\r\n";
                                }

                            }
                            else if (scope.Equals("col"))
                            {
                                if (is_relative_cell(merge_addr, cell.Address))
                                {
                                    int merge_cnt = get_merge_count_col(merge_addr);
                                    html += @"<th colspan=""" + merge_cnt + @"""";
                                    if (align_flg) html += @" style=""text-align: center;""";
                                    html += ">" + cell_val + "</th>\r\n";
                                }
                            }
                            else if (scope.Equals("cross"))
                            {
                                if (is_relative_cell(merge_addr, cell.Address))
                                {
                                    int merge_row_cnt = get_merge_count_row(merge_addr);
                                    int merge_col_cnt = get_merge_count_col(merge_addr);
                                    html += @"<th rowspan=""" + merge_row_cnt + @""" colspan=""" + merge_col_cnt + @"""";
                                    if (align_flg) html += @" style=""text-align: center;""";
                                    html += ">" + cell_val + "</th>\r\n";
                                }
                            }
                        }
                        else
                        {
                            html += @"<th style=""width: " + col_width_attrs[j - c1] + @";";
                            if (align_flg) html += @"text-align: center;";
                            if (border_top_none_flg) html += @"border-top: none;";
                            if (border_bottom_none_flg) html += @"border-bottom: none;";
                            if (align_flg == true || border_top_none_flg == true || border_bottom_none_flg == true) html += @"""";
                            html += ">" + cell_val + "</th>\r\n";
                        }

                    }
                }

                html += "</tr>\r\n";
            }

            html += "</table>";

            html = prefix + "\r\n" + html + "\r\n" + sufix + "\r\n";

            return html;

        }

        //結合方向判定
        private string get_merge_scope(string address)
        {
            string scope = "";
            Regex pt = new Regex(@"(\$)([a-zA-Z]+?)(\$)([0-9]+?)(:)(\$)([a-zA-Z]+?)(\$)([0-9]+)");
            if (!pt.IsMatch(address)) return "";
            Match mt = pt.Match(address);
            string stcol = mt.Groups[2].Value;
            string encol = mt.Groups[7].Value;
            int strow = Int32.Parse(mt.Groups[4].Value);
            int enrow = Int32.Parse(mt.Groups[9].Value);
            if (strow == enrow && !stcol.Equals(encol)) scope = "col";
            else if (strow != enrow && stcol.Equals(encol)) scope = "row";
            else if (strow != enrow && !stcol.Equals(encol)) scope = "cross";
            return scope;
        }

        //セル結合の基準セルかどうか判定
        private Boolean is_relative_cell(string merge_addr, string cell_addr)
        {
            Regex pt = new Regex(@"(\$)([a-zA-Z]+?)(\$)([0-9]+?)(:)(\$)([a-zA-Z]+?)(\$)([0-9]+)");
            if (pt.IsMatch(merge_addr))
            {
                Match mt = pt.Match(merge_addr);
                string capt = mt.Groups[1].Value + mt.Groups[2].Value + mt.Groups[3].Value + mt.Groups[4];
                if (cell_addr.Equals(capt)) return true;
                else return false;
            }
            return false;
        }

        //行結合数を取得
        private int get_merge_count_row(string address)
        {
            int cnt = 0;
            Regex pt = new Regex(@"(\$)([a-zA-Z]+?)(\$)([0-9]+?)(:)(\$)([a-zA-Z]+?)(\$)([0-9]+)");
            if (!pt.IsMatch(address)) return -1;
            Match mt = pt.Match(address);
            int strow = Int32.Parse(mt.Groups[4].Value);
            int enrow = Int32.Parse(mt.Groups[9].Value);
            cnt = (enrow - strow) + 1;
            return cnt;
        }

        //列結合数を取得
        private int get_merge_count_col(string address)
        {
            int cnt = 0;
            string[] cols = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ" };
            Regex pt = new Regex(@"(\$)([a-zA-Z]+?)(\$)([0-9]+?)(:)(\$)([a-zA-Z]+?)(\$)([0-9]+)");
            if (!pt.IsMatch(address)) return -1;
            Match mt = pt.Match(address);
            string stcol = mt.Groups[2].Value;
            string encol = mt.Groups[7].Value;
            int stcnt = 0;
            int encnt = 0;
            for (int i = 0; i < cols.Length; i++)
            {
                string vl = cols[i];
                if (vl.Equals(stcol)) stcnt = i;
                if (vl.Equals(encol)) encnt = i;
            }
            cnt = (encnt - stcnt) + 1;
            return cnt;
        }

        //br変換
        private string br_encode(string str)
        {
            if (str == "" || str == null) return "";
            Regex pat = new Regex(@"(\r\r\n|\r|\r\n)+", RegexOptions.Compiled | RegexOptions.Multiline);
            if (!pat.IsMatch(str)) return str;
            return pat.Replace(str, "<br>");
        }

    }
}
