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
using System.Text.RegularExpressions;


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

        private void LangaugeCode_Loaded(object sender, RoutedEventArgs e)
        {
            List<string> Language = Properties.Resources.LanguageCode.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).ToList();

            SRC_Cbox.ItemsSource = Language;
            TGT_Cbox.ItemsSource = Language;
        }

        private void ByFile_rBtn_Checked(object sender, RoutedEventArgs e)
        {
            SRC_File_label.Content = "Source File :";
            TGT_File_label1.IsEnabled = true;

            TGT_File_label1.Content = "Target File :";
            TGT_File_label1.Visibility = Visibility.Visible;

            TGT_File_tbx.IsEnabled = true;
            TGT_File_tbx.Visibility = Visibility.Visible;

            SRC_Column_label.IsEnabled = false;
            SRC_Column_tbx.IsEnabled = false;
            TGT_Column_label.IsEnabled = false;
            TGT_Column_tbx.IsEnabled = false;
        }

        private void ByMapplist_rBtn_Checked(object sender, RoutedEventArgs e)
        {
            SRC_File_label.Content = "File List :";
            TGT_File_label1.IsEnabled = false;
            TGT_File_label1.Visibility = Visibility.Hidden;
            TGT_File_tbx.IsEnabled = false;
            TGT_File_tbx.Visibility = Visibility.Hidden;

            SRC_Column_label.IsEnabled = false;
            SRC_Column_tbx.IsEnabled = false;
            TGT_Column_label.IsEnabled = false;
            TGT_Column_tbx.IsEnabled = false;
        }

        private void ByExcel_rBtn_Checked(object sender, RoutedEventArgs e)
        {
            SRC_File_label.Content = "Excel File :";
            TGT_File_label1.IsEnabled = false;
            TGT_File_label1.Visibility = Visibility.Hidden;
            TGT_File_tbx.IsEnabled = false;
            TGT_File_tbx.Visibility = Visibility.Hidden;

            SRC_Column_label.IsEnabled = true;
            SRC_Column_tbx.IsEnabled = true;
            TGT_Column_label.IsEnabled = true;
            TGT_Column_tbx.IsEnabled = true;
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            Regex Lang_pattern = new Regex(@"(.+?)\s->(?<Code>.+?-.+?)\b");
            Match SRC_Code_Match = Lang_pattern.Match(SRC_Cbox.SelectedValue.ToString());
            Match TGT_Code_Match = Lang_pattern.Match(TGT_Cbox.SelectedValue.ToString());

            #region By FileList
            if (ByMapplist_rBtn.IsChecked == true)
            {
                FileInfo List_Info = new FileInfo(SRC_File_tbx.Text);
                ExcelPackage List_Package = new ExcelPackage(List_Info);
                ExcelWorksheet List_WS = List_Package.Workbook.Worksheets[1];

                for (int i = 2; i <= List_WS.Dimension.End.Row; i++)
                {
                    ResxFileParser.ResxFileParser SRC_RESX = new ResxFileParser.ResxFileParser();
                    SRC_RESX.FileParser(List_WS.Cells[i, 1].Text);

                    ResxFileParser.ResxFileParser TGT_RESX = new ResxFileParser.ResxFileParser();
                    TGT_RESX.FileParser(List_WS.Cells[i, 2].Text);

                    FileInfo Bilingual_Info = new FileInfo(Path.GetDirectoryName(List_WS.Cells[i, 1].Text) + @"\" + Path.GetFileNameWithoutExtension(List_WS.Cells[i, 1].Text) + ".xlsx");
                    ExcelPackage Bilingual_Package = new ExcelPackage();
                    ExcelWorksheet Bilingual_WS = Bilingual_Package.Workbook.Worksheets.Add(Path.GetFileNameWithoutExtension(List_WS.Cells[i, 1].Text));
                    int current_position = 2;
                    foreach (var SRC_Value in SRC_RESX.DataName)
                    {
                        foreach (var TGT_Value in TGT_RESX.DataName)
                        {
                            if (SRC_Value.Key == TGT_Value.Key)
                            {
                                Bilingual_WS.Cells[current_position, 1].Value = SRC_Value.Key;
                                Bilingual_WS.Cells[current_position, 2].Value = SRC_Value.Value;
                                Bilingual_WS.Cells[current_position, 3].Value = TGT_Value.Value;
                                current_position++;
                            }
                        }
                    }
                    XLZDocument ReviewXLZ = new XLZDocument(SRC_Code_Match.Groups["Code"].Value, "", TGT_Code_Match.Groups["Code"].Value);
                    int processCount = 1;
                    for (int z = 2; z <= Bilingual_WS.Dimension.End.Row; z++)
                    {
                        string[][] split_String = SentenceSplitter.SentenceSplitter.GetSplitedSegments(Bilingual_WS.Cells[z, 2].Text, Bilingual_WS.Cells[z, 3].Text, SRC_Code_Match.Groups["Code"].Value, TGT_Code_Match.Groups["Code"].Value);
                        for (int j = 0; j < split_String[0].Count(); j++)
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
                    ReviewXLZ.Save(Path.GetDirectoryName(List_WS.Cells[i, 1].Text) + @"\" + Path.GetFileNameWithoutExtension(List_WS.Cells[i, 1].Text) + ".xlz");
                }
            }
            #endregion
            #region By File
            if (ByFile_rBtn.IsChecked == true)
            {
                ResxFileParser.ResxFileParser SRC_RESX = new ResxFileParser.ResxFileParser();
                SRC_RESX.FileParser(SRC_File_tbx.Text);

                ResxFileParser.ResxFileParser TGT_RESX = new ResxFileParser.ResxFileParser();
                TGT_RESX.FileParser(TGT_File_tbx.Text);

                FileInfo Bilingual_Info = new FileInfo(Path.GetDirectoryName(SRC_File_tbx.Text) + @"\" + Path.GetFileNameWithoutExtension(SRC_File_tbx.Text) + ".xlsx");
                ExcelPackage Bilingual_Package = new ExcelPackage();
                ExcelWorksheet Bilingual_WS = Bilingual_Package.Workbook.Worksheets.Add(Path.GetFileNameWithoutExtension(SRC_File_tbx.Text));
                int current_position = 2;
                foreach (var SRC_Value in SRC_RESX.DataName)
                {
                    foreach (var TGT_Value in TGT_RESX.DataName)
                    {
                        if (SRC_Value.Key == TGT_Value.Key)
                        {
                            Bilingual_WS.Cells[current_position, 1].Value = SRC_Value.Key;
                            Bilingual_WS.Cells[current_position, 2].Value = SRC_Value.Value;
                            Bilingual_WS.Cells[current_position, 3].Value = TGT_Value.Value;
                            current_position++;
                        }
                    }
                }

                XLZDocument ReviewXLZ = new XLZDocument(SRC_Code_Match.Groups["Code"].Value, "", TGT_Code_Match.Groups["Code"].Value);
                int processCount = 1;
                for (int i = 2; i <= Bilingual_WS.Dimension.End.Row; i++)
                {
                    string[][] split_String = SentenceSplitter.SentenceSplitter.GetSplitedSegments(Bilingual_WS.Cells[i, 2].Text, Bilingual_WS.Cells[i, 3].Text, SRC_Code_Match.Groups["Code"].Value, TGT_Code_Match.Groups["Code"].Value);
                    for (int j = 0; j < split_String[0].Count(); j++)
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
            }
            #endregion
            #region Excel
            if (ByExcel_rBtn.IsChecked == true)
            {
                int StartRow_Position = int.Parse(Start_Row_tbx.Text);
                FileInfo Bilingual_Info = new FileInfo(SRC_File_tbx.Text);
                ExcelPackage Bilingual_Package = new ExcelPackage(Bilingual_Info);

                XLZDocument ReviewXLZ = new XLZDocument(SRC_Code_Match.Groups["Code"].Value, "", TGT_Code_Match.Groups["Code"].Value);
                XLZDocument NewXLZ = new XLZDocument(SRC_Code_Match.Groups["Code"].Value, "", TGT_Code_Match.Groups["Code"].Value);
                int processCount = 1;
                for (int z = 1; z <= Bilingual_Package.Workbook.Worksheets.Count; z++)
                {
                    ExcelWorksheet Bilingual_WS = Bilingual_Package.Workbook.Worksheets[z];
                    if (Bilingual_WS.Dimension != null)
                    {
                        for (int i = StartRow_Position; i <= Bilingual_WS.Dimension.End.Row; i++)
                        {
                            if (Bilingual_WS.Cells[TGT_Column_tbx.Text + i].Text.Length < 1)
                            {
                                string[][] split_String = SentenceSplitter.SentenceSplitter.GetSplitedSegments(Bilingual_WS.Cells[SRC_Column_tbx.Text + i].Text, Bilingual_WS.Cells[SRC_Column_tbx.Text + i].Text, SRC_Code_Match.Groups["Code"].Value, TGT_Code_Match.Groups["Code"].Value);
                                for (int j = 0; j < split_String[0].Count(); j++)
                                {
                                    string currentSRC_String = split_String[0][j];
                                    var TU = NewXLZ.AddTransUnit(new XliffTransUnit((NewXLZ.TransUnits.Count + 1).ToString(), true), true);
                                    TU.SourceRaw = System.Security.SecurityElement.Escape(currentSRC_String);
                                    NewXLZ.AppendFormattingAfterTransUnit(TU, "\r\n");
                                }

                            }
                            else
                            {
                                string[][] split_String = SentenceSplitter.SentenceSplitter.GetSplitedSegments(Bilingual_WS.Cells[SRC_Column_tbx.Text + i].Text, Bilingual_WS.Cells[TGT_Column_tbx.Text + i].Text, SRC_Code_Match.Groups["Code"].Value, TGT_Code_Match.Groups["Code"].Value);
                                for (int j = 0; j < split_String[0].Count(); j++)
                                {
                                    string currentSRC_String = split_String[0][j];
                                    string currentTGT_String = split_String[1][j];
                                    processCount++;
                                    var TU = ReviewXLZ.AddTransUnit(new XliffTransUnit((ReviewXLZ.TransUnits.Count + 1).ToString(), true), true);
                                    TU.SourceRaw = System.Security.SecurityElement.Escape(currentSRC_String);
                                    TU.TranslationRaw = System.Security.SecurityElement.Escape(currentTGT_String);
                                    TU.MatchPercent = 100;
                                    ReviewXLZ.AppendFormattingAfterTransUnit(TU, "\r\n");
                                }
                            }
                        }
                        NewXLZ.Save(Path.GetDirectoryName(SRC_File_tbx.Text) + @"\" + Path.GetFileNameWithoutExtension(SRC_File_tbx.Text) + "_Sheet" + z.ToString() + "_New.xlz");
                        ReviewXLZ.Save(Path.GetDirectoryName(SRC_File_tbx.Text) + @"\" + Path.GetFileNameWithoutExtension(SRC_File_tbx.Text) + "_Sheet" + z.ToString() + "_Review.xlz");
                    }
                }
            }
            #endregion
        MessageBox.Show("Done");
        }
    }
}
