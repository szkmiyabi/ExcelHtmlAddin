﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelHtmlAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void doCreateTableTagButton_Click(object sender, RibbonControlEventArgs e)
        {
            create_table_tag();
        }
    }
}
