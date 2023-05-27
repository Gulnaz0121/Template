using System;
using Microsoft.Win32;
using System.Globalization;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;

namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для Yunusova_4337.xaml
    /// </summary>
    public partial class Yunusova_4337 : Window
    {
        public Yunusova_4337()
        {
            InitializeComponent();
        }
        private void ButtonImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }

            using (EmployeeEntities db = new EmployeeEntities())
            {
                for (var i = 1; i < 11; i++)
                {
                    db.Employee.Add(new Employee()
                    {
                        RoleEmp = list[i, 0],
                        FIO = list[i, 1],
                        LoginEmp = list[i, 2],
                        PasswordEmp = list[i, 3]
                    });

                }
                db.SaveChanges();
                MessageBox.Show("Данные добавлены!");
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
            }
        }

        private void ButtonExport_Click(object sender, RoutedEventArgs e)
        {
            List<Employee> employees;
            List<Role> roles;
            using (EmployeeEntities db = new EmployeeEntities())
            {
                employees = db.Employee.ToList();
                roles = db.Role.ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = roles.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < roles.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = roles[i].RoleEmp;
                worksheet.Cells[1][startRowIndex + 1] = "Логин";
                worksheet.Cells[2][startRowIndex + 1] = "Пароль";
                startRowIndex++;
                var categ = employees.GroupBy(s => s.RoleEmp).ToList();
                foreach (var c in categ)
                {
                    if (c.Key == roles[i].RoleEmp)
                    {
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                        headerRange.Merge();
                        headerRange.Value = roles[i].RoleEmp;
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;
                        foreach (Employee c1 in employees)
                        {
                            if (c1.RoleEmp == c.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = c1.LoginEmp;
                                worksheet.Cells[2][startRowIndex] = c1.PasswordEmp;
                                startRowIndex++;
                            }
                        }
                        worksheet.Cells[1][startRowIndex].Formula = $"=СЧЁТ(A3:A{startRowIndex - 1})";
                        worksheet.Cells[1][startRowIndex].Font.Bold = true;
                    }
                    else
                    {
                        continue;
                    }
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }

        private void ButtonImportJSON_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void ButtonExportWord_Click(object sender, RoutedEventArgs e)
        {

        }

        private string GetHashString(string s)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(s);
            MD5CryptoServiceProvider CSP = new
            MD5CryptoServiceProvider();
            byte[] byteHash = CSP.ComputeHash(bytes);
            string hash = "";
            foreach (byte b in byteHash)
            {
                hash += string.Format("{0:x2}", b);
            }
            return hash;
        }
    }
}
