using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelHtmlAddin
{
    public partial class PrevForm : Form
    {
        //コンストラクタ
        public PrevForm()
        {
            InitializeComponent();
            TopMost = true;
        }

        //Ctrl+A実装
        private void reportText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                e.SuppressKeyPress = true; //beep禁止
                reportText.SelectAll();
            }
        }

        //コピーして閉じる
        private void doCopyAndCloseButton_Click(object sender, EventArgs e)
        {
            string src = reportText.Text;
            try
            {
                Clipboard.SetDataObject(src);
            }
            catch(Exception ex)
            {
                MessageBox.Show("コピー失敗しました。再度実行してください。\r\nシステムエラー:" + ex.Message);
                return;
            }
            MessageBox.Show("コードをクリップボードにコピーしました。");
            this.Close();
        }

        //キャンセル
        private void doCancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
