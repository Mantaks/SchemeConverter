using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace SchemeConverter
{
    public partial class Convert
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void convertButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorksheet();
            //  Excel.Workbook workbook = Globals.ThisAddIn.GetActiveWorkbook();
            Excel.Shapes shapes = currentSheet.Shapes;
            string a=null;
            foreach (Excel.Shape item in shapes)
            {
                a = item.Name;
            }
            Excel.Shape shape = shapes.Item(a);
            SmartArt smart = shape.SmartArt;
            SmartArtNodes nodes = smart.AllNodes;

            List<TextFrame2> textFrame2s = new List<TextFrame2>();
            foreach (SmartArtNode node in nodes)
            {
                textFrame2s.Add(node.TextFrame2);
            }
            TextRange2 range;
            List<string> strings = new List<string>();
            foreach (TextFrame2 textframe2 in textFrame2s)
            {
                range = textframe2.TextRange;
                strings.Add(range.Text);
            }
        }
    }
}
