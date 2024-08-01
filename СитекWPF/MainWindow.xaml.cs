using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Globalization;
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
using System.Windows.Threading;
using СитекWPF.DataFiles;
using Microsoft.Office.Interop;
using Word = Microsoft.Office.Interop.Word;
using Path = System.IO;
using System.Reflection;
using Xceed.Words.NET;
using Xceed.Document.NET;
using System.IO;

namespace СитекWPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            // Заполнение БД
            DBCon.conObj = new FIAS_GAREntities();

            // Таймер
            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(2);
            timer.Tick += LoadADDRTable; // Приявязка метода для заполнения таблицы
            timer.Start();
            LoadDatesCombo();
        }

        private void LoadDatesCombo() // Метод для заполнения Combobox датами 
        {
            try
            {
                var fillquery = DBCon.conObj.ADDR_OBJ.Select(a => a.UPDATEDATE).Distinct().OrderByDescending(date => date);
                dateCombo.Items.Clear();
                foreach (var date in fillquery)
                {
                    dateCombo.Items.Add(Convert.ToDateTime(date));
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }
        }

        private void LoadADDRTable(object sender, object e) // Метод для заполнения таблицы
        {
            try
            {
                if (dateCombo.SelectedIndex == -1)
                {
                    var fillquery = DBCon.conObj.ADDR_OBJ.Join(DBCon.conObj.AS_OBJECT_LEVELS,
                                        ad => ad.LEVEL,
                                        aso => aso.LEVEL,
                                        (ad, aso) => new
                                        {
                                            OBJECT = ad.OBJECTID,
                                            CHANGED = ad.CHANGEID,
                                            Name = ad.NAME,
                                            TYPENAME = ad.TYPENAME,
                                            NAME = aso.NAME,
                                            OPERTYPED = ad.OPERTYPEID,
                                            PREV = ad.PREVID,
                                            NEXT = ad.NEXTID,
                                            UPDATEDATE = ad.UPDATEDATE,
                                            STARTDATE = ad.STARTDATE,
                                            ENDDATE = ad.ENDDATE,
                                            ISACTUAL = ad.ISACTUAL,
                                            ISACTIVE = ad.ISACTIVE
                                        }).Where(ad => ad.ISACTIVE == 1).ToList();
                    addrObjTable.ItemsSource = fillquery;

                }
                else
                {
                    string selectedDate = Convert.ToString(dateCombo.SelectedItem);
                    DateTime date = DateTime.Parse(selectedDate);
                    //MessageBox.Show(date.ToString("yyyy-MM-dd"));
                    var fillquery2 = DBCon.conObj.ADDR_OBJ.Join(DBCon.conObj.AS_OBJECT_LEVELS,
                                                ad => ad.LEVEL,
                                                aso => aso.LEVEL,
                                                (ad, aso) => new
                                                {
                                                    OBJECT = ad.OBJECTID,
                                                    CHANGED = ad.CHANGEID,
                                                    Name = ad.NAME,
                                                    TYPENAME = ad.TYPENAME,
                                                    NAME = aso.NAME,
                                                    OPERTYPED = ad.OPERTYPEID,
                                                    PREV = ad.PREVID,
                                                    NEXT = ad.NEXTID,
                                                    UPDATEDATE = ad.UPDATEDATE,
                                                    STARTDATE = ad.STARTDATE,
                                                    ENDDATE = ad.ENDDATE,
                                                    ISACTUAL = ad.ISACTUAL,
                                                    ISACTIVE = ad.ISACTIVE
                                                }).Where(ad => ad.UPDATEDATE == date && ad.ISACTIVE == 1).ToList();
                    addrObjTable.ItemsSource = fillquery2;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void clearBtn_Click(object sender, RoutedEventArgs e) // Кнопка для очистки выпадающего списка
        {
            try
            {
                dateCombo.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void createReportBtn_Click(object sender, RoutedEventArgs e) // Кнопка для вызова метода по сохранению отчёта
        {
            try
            {
                if (dateCombo.SelectedIndex == -1)
                {
                    MessageBox.Show("Выберите дату!");
                }
                else
                {
                    string selectedDate = Convert.ToString(dateCombo.SelectedItem);
                    DateTime date = DateTime.Parse(selectedDate);
                    DateTime dateSort = DateTime.Parse(selectedDate);
                    GetLastDownloadFileInfo(date.ToString("dd-MM-yyyy"), dateSort);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void GetLastDownloadFileInfo(string date, DateTime dateSort) // Метод для сохранения отчёта
        {
            try
            {
                if (Directory.Exists(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Reports")))
                {

                }
                else
                {
                    Directory.CreateDirectory(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Reports"));
                }
                string filePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Reports\Report.docx");
                DocX document = DocX.Create(filePath);
                document.InsertParagraph($"Отчёт по добавленным адресным объектам за {date} \n —--------------------------------------")
                .Font("Arial").FontSize(20).Color(System.Drawing.Color.LightBlue).Alignment = Xceed.Document.NET.Alignment.both;
                Xceed.Document.NET.Paragraph paragraph = document.InsertParagraph();
                paragraph.Alignment = Alignment.left;

                // Получение типов уровней
                var getCountOfLevels = from ad in DBCon.conObj.ADDR_OBJ
                                       join aso in DBCon.conObj.AS_OBJECT_LEVELS on ad.LEVEL equals aso.LEVEL
                                       where ad.UPDATEDATE == dateSort && ad.ISACTIVE == 1
                                       select aso.NAME;
                var distinctLevelNames = getCountOfLevels.Distinct().ToList();

                foreach (var levelName in distinctLevelNames)
                {
                    string paragraphTitle = levelName;
                    // Получение наименований зданий
                    var getCountOfNames = from ad in DBCon.conObj.ADDR_OBJ
                                          join aso in DBCon.conObj.AS_OBJECT_LEVELS on ad.LEVEL equals aso.LEVEL
                                          orderby ad.NAME ascending
                                          where ad.UPDATEDATE == dateSort && ad.ISACTIVE == 1 && aso.NAME == paragraphTitle
                                          select ad.NAME;
                    var distinctNames = getCountOfNames.Distinct().ToList();

                    // Получение типов зданий
                    var getCountOfTypes = from ad in DBCon.conObj.ADDR_OBJ
                                          join aso in DBCon.conObj.AS_OBJECT_LEVELS on ad.LEVEL equals aso.LEVEL
                                          where ad.UPDATEDATE == dateSort && ad.ISACTIVE == 1 && aso.NAME == paragraphTitle
                                          select ad.TYPENAME;
                    var distinctTypes = getCountOfTypes.Distinct().ToList();

                    // Создание таблицы
                    Xceed.Document.NET.Table table = document.AddTable(distinctNames.Count + 1, 2);
                    table.Alignment = Alignment.center;
                    table.Design = TableDesign.TableGrid;
                    table.Rows[0].Cells[0].Paragraphs[0].Append("Тип объекта").Bold().Alignment = Alignment.center;
                    table.Rows[0].Cells[1].Paragraphs[0].Append("Наименование").Bold().Alignment = Alignment.center;

                    for (int j = 0; j < distinctNames.Count; j++)
                    {
                        table.Rows[j + 1].Cells[1].Paragraphs[0].Append(distinctNames[j]);
                        for (int k = 0; k < distinctTypes.Count; k++)
                        {
                            table.Rows[j + 1].Cells[0].Paragraphs[0].Append(distinctTypes[k]);
                        }
                    }

                    document.InsertParagraph(paragraphTitle).FontSize(14).Font("Arial").InsertTableAfterSelf(table);
                    document.InsertParagraph("").Alignment = Alignment.center;
                }

                document.Save();
                MessageBox.Show($"Отчёт создан в {filePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
