using Lab2.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;
namespace Lab2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window

    {
        public List<Crimes> recordsFull = new List<Crimes>();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void My_Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                TakeDataFromExcel("thrlist.xlsx");
                MessageBox.Show("Выполнен поиск локальных данных. Данные были загружены");
            }
            catch (Exception)
            {

                if (MessageBox.Show("Выполнен поиск локальных данных. Данные отсутствуют. Загрузить их?", "Запуск", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    DownloadData();
                    TakeDataFromExcel("tempered_thrlist.xlsx");
                }
            }
        }

        private void ShowInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (recordsFull.Count == 0) MessageBox.Show("Таблицы не существует. Загрузите её"); else MessageBox.Show("Показываю таблицу");
            Grid.Items.Refresh();
        }

        private void RecordInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (recordsFull.Count == 0)
            {
                try
                {
                    DownloadData();
                    TakeDataFromExcel("tempered_thrlist.xlsx");
                    MessageBox.Show("Я получил новые данные");

                }
                catch (COMException)
                {
                    MessageBox.Show("Проверьте соединение с интернетом и повторите попытку");
                }
                catch (Exception)
                {
                    MessageBox.Show("Неизвестная ошибка");
                }
            }
            else
            {
                DownloadData();
                Updater("tempered_thrlist.xlsx");

                MessageBox.Show("Я обновил данные");
            }





        }

        private void Saver()
        {
            try
            {
                if (File.Exists("thrlist.xlsx"))
                {

                    File.Replace("tempered_thrlist.xlsx", "thrlist.xlsx", null, true); //хочу спать
                    File.Delete("tempered_thrlist.xlsx");
                    MessageBox.Show("FileDeleted");
                }
                else
                {
                    File.Copy("tempered_thrlist.xlsx", "thrlist.xlsx");
                    File.Delete("tempered_thrlist.xlsx");
                    MessageBox.Show("Данные сохранены");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }

        }

        private void DownloadData()
        {
            using (WebClient wc = new WebClient())
            {
                wc.DownloadFileAsync(new System.Uri("https://bdu.fstec.ru/files/documents/thrlist.xlsx"), "tempered_thrlist.xlsx");
            }
        }

        private void TakeDataFromExcel(string road)
        {
            string path = Path.GetFullPath(road);
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            string[,] list = new string[lastCell.Row - 2, lastCell.Column - 2];
            for (int i = 0; i < (int)lastCell.Row - 2; i++)
                for (int j = 0; j < (int)lastCell.Column - 2; j++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 3, j + 1].Text.ToString();
                }
            recordsFull = Inizialisation(list);

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
        }


        private List<Crimes> Inizialisation(string[,] list)
        {
            List<Crimes> newList = new List<Crimes>();
            for (int i = 0; i < list.GetLongLength(0); i++)
            {
                newList.Add(new Crimes(list[i, 0], list[i, 1], list[i, 2], list[i, 3], list[i, 4], list[i, 5], list[i, 6], list[i, 7]));
            }
            return newList;
        }

        private void Updater(string road)
        {
            string path = Path.GetFullPath(road);
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            string[,] list = new string[lastCell.Row - 2, lastCell.Column - 2];
            for (int i = 0; i < (int)lastCell.Row - 2; i++)
                for (int j = 0; j < (int)lastCell.Column - 2; j++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 3, j + 1].Text.ToString();
                }
            List<Crimes> tempList = Inizialisation(list);
            string changes = null;
            int count = 0;
            if (tempList.Count > recordsFull.Count) count = tempList.Count; else count = recordsFull.Count;
            for (int i = 0; i < count; i++)
            {
                if (tempList[i].Id != recordsFull[i].Id) changes += ($"Было: {recordsFull[i].Id}. Стало: {tempList[i].Id}\n");
                if (tempList[i].Name != recordsFull[i].Name) changes += ($"Было: {recordsFull[i].Name}. Стало: {tempList[i].Name}\n");
                if (tempList[i].Discription != recordsFull[i].Discription) changes += ($"Было: {recordsFull[i].Discription}. Стало: {tempList[i].Discription}\n");
                if (tempList[i].Source != recordsFull[i].Source) changes += ($"Было: {recordsFull[i].Source}. Стало: {tempList[i].Source}\n");
                if (tempList[i].ImpactedObject != recordsFull[i].ImpactedObject) changes += ($"Было: {recordsFull[i].ImpactedObject}. Стало: {tempList[i].ImpactedObject}\n");
                if (tempList[i].Confidentiality != recordsFull[i].Confidentiality) changes += ($"Было: {recordsFull[i].Confidentiality}. Стало: {tempList[i].Confidentiality}\n");
                if (tempList[i].Integrity != recordsFull[i].Integrity) changes += ($"Было: {recordsFull[i].Integrity}. Стало: {tempList[i].Integrity}\n");
                if (tempList[i].Accessibility != recordsFull[i].Accessibility) changes += ($"Было: {recordsFull[i].Accessibility}. Стало: {tempList[i].Accessibility}\n");
            }
            recordsFull = tempList;
            if (changes != null)
            {
                MessageBox.Show(changes);
            }

        }


        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Данные сохранены на этот компьютер. Они будут автоматически загружены при следующем запуске программы! Также ты можешь найти таблицу в одной из папок");
            Saver();

        }

        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            Grid.ItemsSource = recordsFull;
        }

        private void Grid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Crimes path = Grid.SelectedItem as Crimes;
                MessageBox.Show("Идентификатор угрозы: " + path.Id + "\nНаименование угрозы: " + path.Name +
                    "\nОписание угрозы: " + path.Discription + "\nИсточник угрозы: " + path.Source
                    + "\nОбъект воздействия угрозы: " + path.ImpactedObject + "\nНарушение конфиденциальности: "
                    + path.Confidentiality + "\nНарушение целостности: " + path.Integrity + "\nНарушение доступности: " + path.Accessibility);
            }
            catch (Exception) { MessageBox.Show("Error"); }
        }
        private void OnAutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            PropertyDescriptor propertyDescriptor = (PropertyDescriptor)e.PropertyDescriptor;
            e.Column.Header = propertyDescriptor.DisplayName;
            if (propertyDescriptor.DisplayName == "Discription") e.Cancel = true;
            if (propertyDescriptor.DisplayName == "Source") e.Cancel = true;
            if (propertyDescriptor.DisplayName == "ImpactedObject") e.Cancel = true;
            if (propertyDescriptor.DisplayName == "Confidentiality") e.Cancel = true;
            if (propertyDescriptor.DisplayName == "Integrity") e.Cancel = true;
            if (propertyDescriptor.DisplayName == "Accessibility") e.Cancel = true;
        }
    }

}
