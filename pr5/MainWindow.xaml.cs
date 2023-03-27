using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Word;


namespace pr5
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Btn_exit_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        private void Btn_create_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = wordApp.Documents.Add();
            Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(doc.Range(), 9, 2);


            table.Cell(1, 1).Range.Text = "Андрей-раб";
            table.Cell(1, 2).Range.Text = "274-88-17";

            table.Cell(2, 1).Range.Text = "Света-Х";
            table.Cell(2, 2).Range.Text = "+38(067)7030356";

            table.Cell(3, 1).Range.Text = "ЖКХ";
            table.Cell(3, 2).Range.Text = "22-345-72";

            table.Cell(4, 1).Range.Text = "Справка";
            table.Cell(4, 2).Range.Text = "009";

            table.Cell(5, 1).Range.Text = "Александр Степанович";
            table.Cell(5, 2).Range.Text = "223-67-67 доп 32-67";

            table.Cell(6, 1).Range.Text = "Мама-дом";
            table.Cell(6, 2).Range.Text = "570-38-76";

            table.Cell(7, 1).Range.Text = "Карапузова Таня";
            table.Cell(7, 2).Range.Text = "201-72-23 пямой моб";

            table.Cell(8, 1).Range.Text = "Погода сегодня";
            table.Cell(8, 2).Range.Text = "001";

            table.Cell(9, 1).Range.Text = "Театр Браво";
            table.Cell(9, 2).Range.Text = "216-40-22";


            // Установка границ для всех ячеек таблицы
            foreach (Cell cell in table.Range.Cells)
            {
                // Установка стиля границ
                cell.Range.Borders.Enable = 1;
                cell.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                cell.Range.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
                // Установка цвета границ
                cell.Range.Borders.OutsideColor = WdColor.wdColorBlack;
                cell.Range.Borders.InsideColor = WdColor.wdColorBlack;
            }


            doc.SaveAs2(@"C:\\VisualStudio\\Vs_Projects\\C#\\PR_Titov\\pr5\tableC#");
            doc.Close();
            wordApp.Quit();


            MessageBoxResult result = MessageBox.Show("Хотите узнать имя файла?", "Создан Word файл", MessageBoxButton.YesNo);

            if (result == MessageBoxResult.Yes)
            {
                MessageBox.Show("tableC#.docx", "Название файла");
            }
            else
            {
                MessageBox.Show("Файл сохранен");
            }
        }
    }
}
