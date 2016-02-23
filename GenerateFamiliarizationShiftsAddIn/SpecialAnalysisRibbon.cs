using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace GenerateFamiliarizationShiftsAddIn
{
    public partial class SpecialAnalysisRibbon
    {
        private void SpecialAnalysisRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void runButton_Click(object sender, RibbonControlEventArgs e)
        {
            Window window = e.Control.Context;
            List<string> names = new List<string>();
            List<int> vals = new List<int>();
            List<int> output = new List<int>();

            try
            {
                foreach (Worksheet sheet in window.Application.ActiveWorkbook.Worksheets)
                {
                    names.Add(sheet.Name);

                    Range r0 = sheet.get_Range("F4");
                    Range r1 = sheet.get_Range("G3");

                    Range r2 = sheet.get_Range("I7");
                    Range r3 = sheet.get_Range("J6");

                    Range r4 = sheet.get_Range("L10");
                    Range r5 = sheet.get_Range("M9");

                    Range r6 = sheet.get_Range("O13");
                    Range r7 = sheet.get_Range("P12");

                    Range r8 = sheet.get_Range("R16");
                    Range r9 = sheet.get_Range("S15");

                    Range[] cells = { r0, r1, r2, r3, r4, r5, r6, r7, r8, r9 };
                    int o = 1;
                    for (int i = 0; i < cells.Length; i += 2)
                    {
                        int sum = 0;
                        if (cells[i].Value2 == null || cells[i + 1].Value2 == null)
                            o = 0;
                        if (!(cells[i].Value2 == null))
                            sum += (int)cells[i].Value2;
                        if (!(cells[i + 1].Value2 == null))
                            sum += (int)cells[i + 1].Value2;

                        vals.Add(sum);
                    }
                    output.Add(o);
                }
            }
            catch (Exception) { MessageBox.Show("Error parsing input files sheets."); return; }

            try {
                Worksheet s = window.Application.ActiveWorkbook.Worksheets.Add();

                s.Name = "FamShiftOutput";

                generateHeader(s);

                for (int i = 0; i < names.Count; i++)
                {
                    int i2 = i + 2;
                    s.get_Range("A" + i2).Value = names[i];
                    s.get_Range("B" + i2).Value = output[i];

                    int root = i * 5;

                    s.get_Range("C" + i2).Value = vals[root + 0];
                    s.get_Range("D" + i2).Value = vals[root + 1];
                    s.get_Range("E" + i2).Value = vals[root + 2];
                    s.get_Range("F" + i2).Value = vals[root + 3];
                    s.get_Range("G" + i2).Value = vals[root + 4];

                    s.get_Range("C" + i2).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189, 215, 238));
                    s.get_Range("D" + i2).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 204, 204));
                    s.get_Range("E" + i2).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(198, 244, 180));
                    s.get_Range("F" + i2).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 204, 255));
                    s.get_Range("G" + i2).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(248, 203, 173));
                }

                s.Columns.AutoFit();
            }
            catch (Exception) { MessageBox.Show("Error generating output sheet."); return; }

            MessageBox.Show("Success");
        }

        private void generateHeader(Worksheet s)
        {
            s.get_Range("A1").Value = "ID";
            s.get_Range("B1").Value = "Correct_Output";
            s.get_Range("C1").Value = "Fam1_Shifts";
            s.get_Range("D1").Value = "Fam2_Shifts";
            s.get_Range("E1").Value = "Fam3_Shifts";
            s.get_Range("F1").Value = "Fam4_Shifts";
            s.get_Range("G1").Value = "Fam5_Shifts";
        }
    }
}
