using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using ExcelRefineAddIn.Catalog;
using ExcelRefineAddIn.Service;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ExcelRefineAddIn
{
    public partial class ExcelRefineRibbon
    {
        private readonly VstoExcelService _excelService = VstoExcelService.Instance;
        private readonly CsvService _csvService = CsvService.Instance;
        private readonly SerilogService _logService = SerilogService.Instance;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            // initialize components
            foreach (var co in CharsetCatalog.CharsetOptions)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = co.DisplayName;
                item.Tag = co.Encoding;
                charsetDrd.Items.Add(item);
            }

            foreach(var nl in NewLineCatalog.NewLineOptions)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = nl.DisplayName; ;
                item.Tag = nl.NewLine;
                newLineDrd.Items.Add(item);
            }

            // Note: To modify width of RibbonDropDown, set SizeString in designer manually.
        }
        private void chooseFolderTbt_Click(object sender, RibbonControlEventArgs e)
        {
            chooseFolderTbt.Checked = true;
            saveToBookFolderTbt.Checked = false;
        }
        private void saveToBookFolderTbt_Click(object sender, RibbonControlEventArgs e)
        {
            saveToBookFolderTbt.Checked = true;
            chooseFolderTbt.Checked = false;
        }

        private void saveAsCsvBtn_Click(object sender, RibbonControlEventArgs e)
        {
            saveFile(",", ".csv");
        }

        private void saveAsTsvBtn_Click(object sender, RibbonControlEventArgs e)
        {
            saveFile("\t", ".tsv");
        }

        private void saveFile(String delimiter, String extensionWithDot)
        {
            Workbook book = null;
            Worksheet sheet = null;

            try
            {
                book = Globals.ThisAddIn.Application.ActiveWorkbook;
                String fullName = book.FullName;
                String newExtFullName = Path.ChangeExtension(fullName, extensionWithDot);

                // if the workbook is not saved yet, use MyDocuments as default folder
                if (String.IsNullOrEmpty(Path.GetDirectoryName(newExtFullName)))
                {
                    newExtFullName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), newExtFullName);
                }

                //Note: This addin does not support saving file to cloud location such as OneDrive and Sharepoint.
                //If the workbook is located in a cloud folder, even when the user select "Save to book's folder",
                //prompt the user to choose local folder to save. 
                if (newExtFullName.StartsWith("http") || chooseFolderTbt.Checked)
                {
                    if(newExtFullName.StartsWith("http"))
                    {
                        MessageBox.Show("The current workbook is stored on cloud, OneDrive, Sharepiint, etc. Please choose a local folder to save the CSV/TSV file. (Save to cloud is not supported)",
                            "Cloud workbook", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    var dialog = new SaveFileDialog
                    {
                        Title = "Choose location to save",
                        Filter = "CSV file (*.csv)|*.csv|All file (*.*)|*.*",
                        DefaultExt = "csv",
                        FileName = Path.GetFileName(newExtFullName),
                    };
                    if(dialog.ShowDialog() == DialogResult.OK)
                    {
                        newExtFullName = dialog.FileName;
                    }
                    else
                    {
                        Marshal.ReleaseComObject(book);
                        return;
                    }
                }
                else
                {
                    // check if file exists
                    if (File.Exists(newExtFullName))
                    {
                        var dr = MessageBox.Show($"The following file already exists. Do you want to overwrite?\r\n" +
                            $"Filename: {Path.GetFileName(newExtFullName)}\r\nFolder path: {Path.GetDirectoryName(newExtFullName)}",
                            "File already exists", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dr == DialogResult.No)
                        {
                            Marshal.ReleaseComObject(book);
                            return;
                        }
                    }
                }

                //check file lock
                if (File.Exists(newExtFullName))
                {
                    FileStream fs = null;
                    try
                    {
                        fs = File.Open(newExtFullName, FileMode.Open, FileAccess.Read, FileShare.None);
                    }
                    catch
                    {
                        MessageBox.Show("Failed to save file. The existing file is locked. Close the file if opened by another program",
                            "Failed to save", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        Marshal.ReleaseComObject(book);
                        return;
                    }
                    finally
                    {
                        if (fs != null)
                        {
                            fs.Close();
                        }
                    }
                }

                sheet = Globals.ThisAddIn.Application.ActiveSheet;
                var rows = _excelService.ReadActiveSheet(sheet);
                _csvService.save(rows, newExtFullName, delimiter, (String)newLineDrd.SelectedItem.Tag, (Encoding)charsetDrd.SelectedItem.Tag);

                MessageBox.Show(String.Format("File saved as below.\r\nFilename: {0}\r\nFolder path: {1}", Path.GetFileName(newExtFullName),
                    Path.GetDirectoryName(newExtFullName)), "File saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save file. " + ex.Message, "Failed to save", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                _logService.Error("Error", ex);
            }
            finally
            {
                if(sheet != null)
                {
                    Marshal.ReleaseComObject(sheet);
                }

                if(book != null)
                {
                    Marshal.ReleaseComObject(book);
                }
            }
        }

        private void aboutBtn_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new Form
            {
                Text = "About ExcelRefine",
                Size = new Size(420, 200),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var lblProduct = new System.Windows.Forms.Label
            {
                Text = "ExcelRefine",
                Font = new System.Drawing.Font("Segoe UI", 12, FontStyle.Bold),
                Location = new System.Drawing.Point(20, 20),
                AutoSize = true
            };

            var lblVersion = new System.Windows.Forms.Label
            {
                Text = $"Version: {Assembly.GetExecutingAssembly().GetName().Version}",
                Location = new System.Drawing.Point(20, 50),
                AutoSize = true
            };

            var link = new LinkLabel
            {
                Text = "GitHub: https://github.com/oshossan/ExcelRefine",
                Location = new System.Drawing.Point(20, 80),
                AutoSize = true
            };
            link.Links[0].LinkData = "https://github.com/oshossan/ExcelRefine/";
            link.LinkClicked += (s, args) => Process.Start(link.Links[0].LinkData.ToString());

            form.Controls.Add(lblProduct);
            form.Controls.Add(lblVersion);
            form.Controls.Add(link);
            form.ShowDialog();
        }
    }
}
