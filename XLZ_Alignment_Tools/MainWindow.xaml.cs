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
using System.Windows.Media.Animation;
using System.Windows.Navigation;
using System.Reflection;
using ResxFileParser;
using OfficeOpenXml;
using System.IO;
using SentenceSplitter;
using Lionbridge.Sterling.Shared.XLZ;


namespace XLZ_Alignment_Tools
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ByFile_rBtn_Checked(object sender, RoutedEventArgs e)
        {
            SRC_File_label.Content = "Source File :";
            TGT_File_label1.IsEnabled = true;

            TGT_File_label1.Content = "Target File :";
            TGT_File_label1.Visibility = Visibility.Visible;

            TGT_File_tbx.IsEnabled = true;
            TGT_File_tbx.Visibility = Visibility.Visible;
        }

        private void ByMapplist_rBtn_Checked(object sender, RoutedEventArgs e)
        {
            SRC_File_label.Content = "File List :";
            TGT_File_label1.IsEnabled = false;
            TGT_File_label1.Visibility = Visibility.Hidden;
            TGT_File_tbx.IsEnabled = false;
            TGT_File_tbx.Visibility = Visibility.Hidden;
        }


        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            ResxFileParser.ResxFileParser SRC_RESX = new ResxFileParser.ResxFileParser();
            SRC_RESX.FileParser(SRC_File_tbx.Text);

            ResxFileParser.ResxFileParser TGT_RESX = new ResxFileParser.ResxFileParser();
            TGT_RESX.FileParser(TGT_File_tbx.Text);

            #region Generate Excel file to store resx file.
            FileInfo Bilingual_Info = new FileInfo(Path.GetDirectoryName(SRC_File_tbx.Text) + @"\" + Path.GetFileNameWithoutExtension(SRC_File_tbx.Text) + ".xlsx");
            #endregion
            ExcelPackage Bilingual_Package = new ExcelPackage();
            ExcelWorksheet Bilingual_WS = Bilingual_Package.Workbook.Worksheets.Add(Path.GetFileNameWithoutExtension(SRC_File_tbx.Text));
            int current_position = 2;
            foreach (var SRC_Value in SRC_RESX.DataName)
            {
                foreach(var TGT_Value in TGT_RESX.DataName)
                {
                    if(SRC_Value.Key==TGT_Value.Key)
                    {
                        Bilingual_WS.Cells[current_position, 1].Value = SRC_Value.Key;
                        Bilingual_WS.Cells[current_position, 2].Value = SRC_Value.Value;
                        Bilingual_WS.Cells[current_position, 3].Value = TGT_Value.Value;
                        current_position++;
                    }
                }
            }
            XLZDocument ReviewXLZ = new XLZDocument("en-us", "", "zh-cn");
            int processCount = 1;
            for (int i = 2; i <= Bilingual_WS.Dimension.End.Row; i++)
            {
                string[][] split_String = SentenceSplitter.SentenceSplitter.GetSplitedSegments(Bilingual_WS.Cells[i, 2].Text, Bilingual_WS.Cells[i, 3].Text, "en-US", "zh_CN");
                for (int j = 0; j < split_String[0].Count() ; j++)
                {
                    string currentSRC_String = split_String[0][j];
                    string currentTGT_String = split_String[1][j];
                    processCount++;
                    var TU = ReviewXLZ.AddTransUnit(new XliffTransUnit((ReviewXLZ.TransUnits.Count + 1).ToString(), true), true);
                    TU.SourceRaw = System.Security.SecurityElement.Escape(currentSRC_String);
                    TU.TranslationRaw = System.Security.SecurityElement.Escape(currentTGT_String);
                    ReviewXLZ.AppendFormattingAfterTransUnit(TU, "\r\n");
                }
            }
            Bilingual_Package.SaveAs(Bilingual_Info);
            ReviewXLZ.Save(Path.GetDirectoryName(SRC_File_tbx.Text) + @"\" + Path.GetFileNameWithoutExtension(SRC_File_tbx.Text) + ".xlz");

            MessageBox.Show("Done");
            //MessageBox.Show(SRC_RESX.DataName.Count.ToString());
            //MessageBox.Show(TGT_RESX.DataName.Count.ToString());
        }

        private void LangaugeCode_Loaded(object sender, RoutedEventArgs e)
        {
            List<string> Language = Properties.Resources.LanguageCode.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).ToList();

            SRC_Cbox.ItemsSource = Language;
            TGT_Cbox.ItemsSource = Language;
        }
    }
}
