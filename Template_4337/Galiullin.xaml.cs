using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Newtonsoft.Json;
using Microsoft.Win32;
using System.Data.SqlClient;
using System.IO;
using System.Data.Entity;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Windows.Markup;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для Galiullin.xaml
    /// </summary>

    public class Service
    {
        public int IdServices { get; set; }
        public string NameServices { get; set; }
        public string TypeOfService { get; set; }
        public string CodeService { get; set; }
        public int Cost { get; set; }
    }
    public partial class Galiullin : System.Windows.Window
    {
        public Galiullin()
        {
            InitializeComponent();
        }

        private void toDBfromExel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON Files (*.json)|*.json";
            openFileDialog.Title = "Выберите файл JSON";

            if (openFileDialog.ShowDialog() == true)
            {
                string jsonFilePath = openFileDialog.FileName; // Выбранный путь к файлу 
                // Загрузка и десериализация данных из файла JSON
                List<Service> services;
                using (StreamReader r = new StreamReader(jsonFilePath))
                {
                    string json = r.ReadToEnd();
                    services = JsonConvert.DeserializeObject<List<Service>>(json);
                }

                // Сохранение данных в базу данных
                using (var dbContext = new Model2())
                {

                    foreach (Service service in services)
                    {
                        dbContext.MyTables.Add(new MyTable
                        {
                            Id = service.IdServices,
                            Name = service.NameServices,
                            Type = service.TypeOfService,
                            Price = service.Cost
                        });
                        int i = 0;
                        
                    }
                    dbContext.SaveChanges();
                }

                MessageBox.Show("Данные успешно загружены и сохранены в базу данных.");
            }
        }

        private void toExelfromDB_Click(object sender, RoutedEventArgs e)
        {
            var dbContext = new Model2();
            // Замените GetLoadedData() на ваш метод получения данных
            List<MyTable> first = new List<MyTable>();
            List<MyTable> second = new List<MyTable>();
            List<MyTable> third = new List<MyTable>();

            foreach ( var data in dbContext.MyTables.ToList())
            {
                if(data.Price>0 && data.Price<= 350)
                {
                    first.Add(data);
                }
                else if (data.Price > 350 && data.Price <= 800)
                {
                    second.Add(data);
                }
                else
                {
                    third.Add(data);
                }
            }

            // Создание нового документа Word
            var wordApp = new Word.Application();
            var doc = wordApp.Documents.Add();
            List<MyTable> datas = new List<MyTable>();
            for(int i =0; i < 3; i++)
            {
                if (i == 0)
                {
                    datas = first;
                }
                if (i == 1)
                {
                    datas = second;
                }
                if (i == 2)
                {
                    datas = third;
                }

                doc.Words.Last.InsertBreak(WdBreakType.wdPageBreak);

                // Добавление заголовка группы
                var groupHeading = doc.Paragraphs.Add();
                groupHeading.Range.Text = $"Группа: {i+1}";
                groupHeading.Range.Font.Bold = 1;
                groupHeading.Range.InsertParagraphAfter();

                    for (int j =0; j<datas.Count;j++)
                    {
                        var dataRow = doc.Paragraphs.Add();
                        dataRow.Range.Text = $"Id: {datas[j].Id}, Name: {datas[j].Name}, Type: {datas[j].Type} Price: {datas[j].Price}";
                        dataRow.Range.InsertParagraphAfter();
                    }   
            }
            string outputPath = @"C:\ISRPOLR2\Template\Template_4337\result\result.docx";
            doc.SaveAs(outputPath);
            doc.Close();
            wordApp.Quit();
            MessageBox.Show("Успешно");
        }

        private void toExel_Click(object sender, RoutedEventArgs e)
        {
            List<MyTable> first = new List<MyTable>();
            List<MyTable> second = new List<MyTable>();
            List<MyTable> third = new List<MyTable>();
            using(Model2 dbContext = new Model2())
            {
                first = (from a in dbContext.MyTables
                         where a.Price < 350
                         select a).ToList();
                second = (from a in dbContext.MyTables
                         where a.Price < 800 && a.Price>350
                         select a).ToList();
                third = (from a in dbContext.MyTables
                         where a.Price > 800
                         select a).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < 3; i++)
            {
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Cells[1][1] = "Порядковый номер";
                worksheet.Cells[2][1] = "Название услуги";
                worksheet.Cells[3][1] = "Вид услуги";
                worksheet.Cells[4][1] = "Стоимость";
                if (i == 0){
                    for (int j = 2; j < first.Count() + 2; j++)
                    {
                        worksheet.Cells[1][j] = first[j - 2].Id;
                        worksheet.Cells[2][j] = first[j - 2].Name;
                        worksheet.Cells[3][j] = first[j - 2].Type;
                        worksheet.Cells[4][j] = first[j - 2].Price;
                    }
                }
                if (i == 1)
                {
                    for (int k = 2; k < second.Count() + 2; k++)
                    {
                        worksheet.Cells[1][k] = second[k - 2].Id;
                        worksheet.Cells[2][k] = second[k - 2].Name;
                        worksheet.Cells[3][k] = second[k - 2].Type;
                        worksheet.Cells[4][k] = second[k - 2].Price;
                    }
                }
                if (i == 2)
                {
                    for (int c = 2; c < third.Count() + 2; c++)
                    {
                        worksheet.Cells[1][c] = third[c - 2].Id;
                        worksheet.Cells[2][c] = third[c - 2].Name;
                        worksheet.Cells[3][c] = third[c - 2].Type;
                        worksheet.Cells[4][c] = third[c - 2].Price;
                    }
                }
            }
            string filePath = @"C:\Users\79393\OneDrive\Рабочий стол\result.xlsx";
            workbook.SaveAs(filePath);

            // Закрытие книги и приложения Excel
            workbook.Close();
            app.Quit();
        }

        private void toDB_Click(object sender, RoutedEventArgs e)
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
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (var forisrpEntities1 = new Model2())
            {
                for (int i = 1; i < _rows; i++)
                {
                    forisrpEntities1.MyTables.Add(new MyTable() { Id = Convert.ToInt32(list[i, 0]), Name = list[i, 1], Type = list[i, 2], Price = Convert.ToInt32(list[i, 4]) });
                }
                forisrpEntities1.SaveChanges();
            }
            MessageBox.Show("Успешно добавлено!");
        }
    }
}
