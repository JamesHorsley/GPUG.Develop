using System;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using Microsoft.Dexterity.Bridge;
using Microsoft.Dexterity.Applications;
using Microsoft.Dexterity.Applications.DynamicsDictionary;
using System.Diagnostics;
using System.Drawing;

namespace GPUG.Develop
{
    /// <summary>
    /// Access GP Data with selecting from global table buffer
    /// </summary>
    public class Lesson_3e : IDexterityAddIn
    {
        private PmVendorMaintenanceForm vendorForm = Dynamics.Forms.PmVendorMaintenance;

        public void Initialize()
        {
            vendorForm.AddMenuHandler(getRangeVendorData, "Get Ranged Vendor Data", null);
        }

        // creates a new vendor record
        private void getRangeVendorData(object sender, EventArgs e)
        {
            var data = new DataTable();

            // declare table error variable for responses from SQL
            TableError error;

            // reference the global table for the vendor maintenance table PM00200
            var row = Dynamics.Tables.PmVendorMstr;

            // define a key and set the key's value
            row.Key = 1;

            // set the range start
            row.Clear();
            row.VendorId.Value = "A";
            row.RangeStart();

            // set the range end
            row.Clear();
            row.VendorId.Value = "BZZZZZZZZZZZ";
            row.RangeEnd();

            // set record to update, check for errors, and make changes
            error = row.GetFirst();
            if (error == TableError.NoError)
            {
                // add the row to our data table
                data = this.GetTableProperties(row, data);
            }
            else
            {
                MessageBox.Show(error.ToString());
            }

            // create a loop to get the remaining records from the range
            // and add them to the data table until the end is reached
            do
            {
                // get the next record
                error = row.GetNext();
                if (error == TableError.NoError)
                {
                    // add the row to our data table
                    data = this.GetTableProperties(row, data);
                }
                
            } while (error == TableError.NoError);

            // close the table since this is a global table open
            row.Close();

            // show the data retrieved to the user
            ViewDataTable("Data results", data);
        }



        // create an input modal window that shows rows of data and returns a string value of the first column of the selected row
        // changes the column names with "_" as spaces in the column names shown to the user
        private DialogResult ViewDataTable(string title, DataTable datasource, string btnOkTxt = "OK", int winWidth = 640, int winHeight = 480)
        {
            int heightOffset = winHeight - 480;
            int widthOffset = winWidth - 640;

            Form form = new Form();
            Label label = new Label();
            DataGridView dgv = new DataGridView();
            Button buttonOk = new Button();

            form.Text = title;
            dgv.ReadOnly = true;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.AllowUserToResizeColumns = true;
            dgv.AllowUserToResizeRows = false;
            dgv.RowHeadersVisible = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = false;
            dgv.BackgroundColor = Color.White;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.None;
            dgv.AutoGenerateColumns = false;

            // add the columns needed for the datasource
            foreach (DataColumn col in datasource.Columns)
            {
                var dvgCol = new DataGridViewTextBoxColumn();
                dvgCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dvgCol.DataPropertyName = col.ColumnName;
                dvgCol.HeaderText = col.ColumnName.Replace("_", " ");
                dvgCol.Name = col.ColumnName;
                dvgCol.ReadOnly = true;
                dgv.Columns.Add(dvgCol);
            }

            // add a blank column on the end to fill out the grid
            var lastCol = new DataGridViewTextBoxColumn();
            lastCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            lastCol.HeaderText = "";
            lastCol.ReadOnly = true;
            dgv.Columns.Add(lastCol);

            // show the data
            dgv.DataSource = datasource;

            buttonOk.Text = btnOkTxt;
            buttonOk.DialogResult = DialogResult.OK;

            dgv.SetBounds(0, 0, 640 + widthOffset, 440 + heightOffset);
            buttonOk.SetBounds(535 + widthOffset, 445 + heightOffset, 100, 30);

            dgv.Anchor = dgv.Anchor | AnchorStyles.Right | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Top;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(winWidth, winHeight);
            form.Controls.AddRange(new Control[] { dgv, buttonOk });
            form.FormBorderStyle = FormBorderStyle.Sizable;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;

            DialogResult dialogResult = form.ShowDialog();

            // set the return value if returning one
            if (dialogResult == DialogResult.OK)
            {

            }

            return dialogResult;
        }


        // retrieve a datatable from Dexterity table information
        private DataTable GetTableProperties(object row, DataTable dataTable)
        {
            var fields = row.GetType().GetFields();
            var properties = row.GetType().GetProperties();


            // check the DataTable for columns, and if missing, add them
            if (dataTable.Columns.Count == 0)
            {
                foreach (var p in properties)
                {
                    var val = p.GetValue(row);

                    switch (val.ToString())
                    {
                        case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.String]":
                            dataTable.Columns.Add(p.Name, typeof(string));
                            break;
                        case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.Decimal]":
                            dataTable.Columns.Add(p.Name, typeof(decimal));
                            break;
                        case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.Int16]":
                            dataTable.Columns.Add(p.Name, typeof(short));
                            break;
                        case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.Int32]":
                            dataTable.Columns.Add(p.Name, typeof(int));
                            break;
                        case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.Int64]":
                            dataTable.Columns.Add(p.Name, typeof(long));
                            break;
                        case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.DateTime]":
                            dataTable.Columns.Add(p.Name, typeof(DateTime));
                            break;
                        case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.Boolean]":
                            dataTable.Columns.Add(p.Name, typeof(bool));
                            break;
                        default:
                            break;
                    }
                }
            }

            // now get the values and add a datarow
            var newRow = dataTable.NewRow();
            foreach (var p in properties)
            {
                var val = p.GetValue(row);

                switch (val.ToString())
                {
                    case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.String]":
                        newRow[p.Name] = ((FieldReadWrite<string>)val).Value;
                        break;
                    case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.Decimal]":
                        newRow[p.Name] = ((FieldReadWrite<decimal>)val).Value;
                        break;
                    case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.Int16]":
                        newRow[p.Name] = ((FieldReadWrite<short>)val).Value;
                        break;
                    case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.Int32]":
                        newRow[p.Name] = ((FieldReadWrite<int>)val).Value;
                        break;
                    case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.Int64]":
                        newRow[p.Name] = ((FieldReadWrite<long>)val).Value;
                        break;
                    case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.DateTime]":
                        newRow[p.Name] = ((FieldReadWrite<DateTime>)val).Value;
                        break;
                    case "Microsoft.Dexterity.Bridge.FieldReadWrite`1[System.Boolean]":
                        newRow[p.Name] = ((FieldReadWrite<bool>)val).Value;
                        break;
                    default:
                        break;
                }
            }
            dataTable.Rows.Add(newRow);

            return dataTable;
        }

    }
}
