using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;

namespace CfxUtilityGUI
{
    /// <summary>
    /// </summary>
    public partial class CfxManipurateSettingsWindow : Window
    {
        public CfxManipurateSettingsWindow()
        {
            InitializeComponent();
        }

        private void button_BrowseDbPath_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = GUtility.CreateOpenFileDialog(
                  "ChemFinder files (*.cfx)|*.cfx|"
                + "All files (*.*)|*.*");
            if (openFileDialog.ShowDialog() != false)
            {
                this.textBox_DbPath.Text = openFileDialog.FileName;
            }
        }

        private void buttonOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void button_ScaffordCdxFileName_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = GUtility.CreateOpenFileDialog(
                  "ChemDraw files (*.cdx)|*.cdx|"
                + "All files (*.*)|*.*");
            if (openFileDialog.ShowDialog() != false)
            {
                this.textBox_ScaffordCdxFileName.Text = openFileDialog.FileName;
            }
        }

        public void SetValuedFromSetting()
        {
            this.textBox_DbPath.Text = Properties.Settings.Default.LocalDbPath;
            this.textBox_MolIDFieldName.Text = Properties.Settings.Default.MolIDInLocalDb_FieldName;
            this.textBox_LocalCodeFieldName.Text = Properties.Settings.Default.LocalCodeInLocalDb_FieldName;
            this.textBox_ChemNameFieldName.Text = Properties.Settings.Default.ChemName_FieldName;
            this.textBox_ScaffordCdxFileName.Text = Properties.Settings.Default.ScaffordCdx_FileName;
            this.checkBoxLoadImage.IsChecked = Properties.Settings.Default.LoadSubstanceImage;
            this.checkBoxOverwrite.IsChecked = Properties.Settings.Default.Overwrite;
        }

        public void SetSettingFromDialog()
        {
            Properties.Settings.Default.LocalDbPath = this.textBox_DbPath.Text;
            Properties.Settings.Default.MolIDInLocalDb_FieldName = this.textBox_MolIDFieldName.Text;
            Properties.Settings.Default.LocalCodeInLocalDb_FieldName = this.textBox_LocalCodeFieldName.Text;
            Properties.Settings.Default.ChemName_FieldName = this.textBox_ChemNameFieldName.Text;
            Properties.Settings.Default.ScaffordCdx_FileName = this.textBox_ScaffordCdxFileName.Text;
            Properties.Settings.Default.LoadSubstanceImage = this.checkBoxLoadImage.IsChecked ?? false;
            Properties.Settings.Default.Overwrite = this.checkBoxOverwrite.IsChecked ?? false;
        }

        private void button_ResetValues_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.Reset();
            SetValuedFromSetting();
        }
    }
}
