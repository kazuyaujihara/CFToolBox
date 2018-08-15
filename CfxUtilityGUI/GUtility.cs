using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace CfxUtilityGUI
{
    static class GUtility
    {
        public static OpenFileDialog CreateOpenFileDialog(string filter)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.ValidateNames = true;
            openFileDialog.Filter = filter;
            return openFileDialog;
        }

        public static SaveFileDialog CreateSaveFileDialog(string filter, string defaltExt)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.OverwritePrompt = true;
            saveFileDialog.DefaultExt = defaltExt;
            saveFileDialog.Filter = filter;
            saveFileDialog.AddExtension = true;
            return saveFileDialog;
        }
    }
}
