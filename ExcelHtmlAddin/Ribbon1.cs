using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelHtmlAddin
{
    public partial class Ribbon1
    {

        private static PrevForm _PrevFormObj;

        //コンストラクタ
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //Formインスタンス取得
        public static PrevForm PrevFormObj
        {
            get
            {
                if(_PrevFormObj == null || _PrevFormObj.IsDisposed)
                {
                    _PrevFormObj = new PrevForm();
                }
                return _PrevFormObj;
            }
        }

        //表コード出力クリック
        private void doCreateTableTagButton_Click(object sender, RibbonControlEventArgs e)
        {
            do_table_tag_create_view();
        }

        //ファイルに保存クリック
        private void doSaveFileTableTagButton_Click(object sender, RibbonControlEventArgs e)
        {
            do_table_tag_create_save();
        }

        //列幅調整クリック
        private void doFitColumnWidthButton_Click(object sender, RibbonControlEventArgs e)
        {
            do_fit_col_width();
        }
    }
}
