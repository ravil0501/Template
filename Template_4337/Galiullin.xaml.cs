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

        }
    }
}
