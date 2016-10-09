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
    /// Access form properties via name, and interacting with controls on the dex form
    /// </summary>

    public class Lesson_2 : IDexterityAddIn
    {
        // vendor maintenance form
        private PmVendorMaintenanceForm vendorForm = Dynamics.Forms.PmVendorMaintenance;
        private Forms.formVendorMaintenancePopup vendorPopup = new Forms.formVendorMaintenancePopup();

        public void Initialize()
        {
            // when the window is opened, show the additional form, and close it when the vendor master closes
            this.vendorForm.OpenAfterOriginal += new EventHandler(vendorMasterOpened);
            this.vendorForm.CloseAfterOriginal += new EventHandler(vendorMasterClosed);

            // adding a menu option to open forms or fire new events
            this.vendorForm.AddMenuHandler(showFields, "Show field names and properties", null);
        }

        // add menu handling to a window
        private void vendorMasterOpened(object sender, EventArgs e)
        { 
            Debug.Print("The vendor master has opened");

            if (!this.WindowIsOpen(this.vendorPopup))
            {
                // this.vendorPopup.Show();
            }
        }

        // add menu handling to a window
        private void vendorMasterClosed(object sender, EventArgs e)
        {
            Debug.Print("The vendor master is closed");

            if (this.WindowIsOpen(this.vendorPopup))
            {
                // this.vendorPopup.Close();
            }
        }

        // create a menu handler method
        private void showFields(object sender, EventArgs e)
        {
            var curForm = sender as Microsoft.Dexterity.Bridge.DictionaryForm;

            // show the names of the windows available to the forms
            foreach (Window win in curForm.Windows)
            {
                this.ViewDataTable(win.Name, this.GetWindowProperties(win));
            }

        }

        #region Helper Functions

        private bool WindowIsOpen(Form form)
        {
            try
            {
                FormCollection forms = Application.OpenForms;
                foreach (Form f in forms)
                {
                    if (f.Name == form.Name)
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return false;
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

        // retrieve a value by name from a scroll window
        // adapted from James Lyn
        // https://jamesdlyn.wordpress.com/tag/microsoft-dynamics-gp-sdk/
        private DataTable GetWindowProperties(Window dexterityWindow)
        {
            // output
            DataTable dt = new DataTable();
            dt.Columns.Add("Name", typeof(String));
            dt.Columns.Add("Value", typeof(String));

            //Loop through each field
            foreach (var field in dexterityWindow.Fields)
            {
                string name = "";
                string value = "";
                try
                {
                    //Loop through the properties of the field looking for Name and Value
                    foreach (var prop in field.GetType().GetProperties())
                    {
                        switch (prop.Name)
                        {
                            case "Name":
                                name = Convert.ToString(prop.GetValue(field, null)).Trim();
                                break;
                            case "Value":
                                value = Convert.ToString(prop.GetValue(field, null)).Trim();
                                break;
                        }
                    }
                }
                catch { }

                //Do not output rows that have blank fields
                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(value))
                {
                    DataRow row = dt.NewRow();
                    row["Name"] = name;
                    row["Value"] = value;
                    dt.Rows.Add(row);
                }
            }

            return dt;
        }


        // retrieve a value by name from a scroll window
        // adapted from James Lyn
        // https://jamesdlyn.wordpress.com/tag/microsoft-dynamics-gp-sdk/
        private string GetScrollWindowPropertyValue(object sender, string fieldName)
        {
            string result = "";

            // temp output
            DataTable dt = new DataTable();
            dt.Columns.Add("Name", typeof(String));
            dt.Columns.Add("Value", typeof(String));

            //Loop through each field
            foreach (var field in ((ScrollingWindow)sender).Fields)
            {
                string name = "";
                string value = "";
                try
                {
                    //Loop through the properties of the field looking for Name and Value
                    foreach (var prop in field.GetType().GetProperties())
                    {
                        switch (prop.Name)
                        {
                            case "Name":
                                name = Convert.ToString(prop.GetValue(field, null)).Trim();
                                break;
                            case "Value":
                                value = Convert.ToString(prop.GetValue(field, null)).Trim();
                                break;
                        }
                    }
                }
                catch { }

                //Do not output rows that have blank fields
                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(value))
                {
                    DataRow row = dt.NewRow();
                    row["Name"] = name;
                    row["Value"] = value;
                    dt.Rows.Add(row);
                }

                // assign the output value
                if (name == fieldName)
                {
                    result = value;
                    break;
                }
            }

            return result;
        }

        #endregion

    }
}
