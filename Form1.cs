using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AUVPermissionGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.            

            if (result == DialogResult.OK) // Test result.
            {
                var fileName = openFileDialog1.FileName;
                var file = new FileInfo(fileName);

                var dicPerm = ReadPermissions(file);
            }
        }

        private Dictionary<string, List<string>> ReadPermissions(FileInfo file)
        {
            try
            {
                using (var package = new ExcelPackage(file))
                {
                    ExcelWorksheet workSheet = package.Workbook.Worksheets[1];
                    var start = workSheet.Dimension.Start;
                    var end = workSheet.Dimension.End;

                    var roles = new string[]
                    {
                            "FIRM_ADMIN", "FIRM_USER", "BUSINES_ADMIN", "BUSINES_USER", "SUPER_ADMIN",
                            "AUVENIR_ADMIN", "LEAD_AUDITOR", "GENERAL_AUDITOR", "LEAD_CLIENT", "GENERAL_CLIENT"
                    };

                    var dic = new Dictionary<string, List<string>>();
                    foreach (var role in roles)
                    {
                        dic.Add(role, new List<string>());
                    }

                    int permCounts = 0;

                    for (int rowNum = start.Row; rowNum <= end.Row; rowNum++)
                    {
                        var row = workSheet.Cells[string.Format("{0}:{0}", rowNum)];
                        var permCell = row[$"D{rowNum}"];

                        if (permCell != null && permCell.Value != null)
                        {
                            var perm = permCell.GetValue<string>();
                            if (!string.IsNullOrWhiteSpace(perm) && perm.StartsWith("PM_"))
                            {
                                permCounts++;

                                int startCol = 6;
                                for (int colNum = startCol; colNum < 17; colNum++)
                                {
                                    if (colNum == 12) continue;

                                    var cell = row[rowNum, colNum];
                                    if (HasPermission(cell))
                                    {
                                        var roleIndex = colNum < 12 ? (colNum - startCol) : (colNum - startCol - 1);
                                        var role = roles[roleIndex];
                                        dic[role].Add(perm);
                                    }
                                }
                            }
                        }
                    }

                    return dic;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private bool HasPermission(ExcelRange cell)
        {
            if (cell == null || cell.Value == null)
            {
                return false;
            }

            var cellValue = cell.GetValue<string>();
            if (string.IsNullOrWhiteSpace(cellValue))
            {
                return false;
            }

            return cellValue.Equals("Y", StringComparison.InvariantCultureIgnoreCase);
        }
    }
}
