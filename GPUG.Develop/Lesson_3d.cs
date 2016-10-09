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
    public class Lesson_3d : IDexterityAddIn
    {
        private PmVendorMaintenanceForm vendorForm = Dynamics.Forms.PmVendorMaintenance;

        public void Initialize()
        {
            vendorForm.AddMenuHandler(deleteVendorData, "Delete Vendor Data", null);
        }


        // creates a new vendor record
        private void deleteVendorData(object sender, EventArgs e)
        {
            // update the vendor maintenance window to take the record off screen if visible prior to performing the delete
            if (!string.IsNullOrEmpty(vendorForm.PmVendorMaintenance.VendorId.Value))
                vendorForm.PmVendorMaintenance.SaveButton.RunValidate();


            // declare table error variable for responses from SQL
            TableError error;

            // reference the global table for the vendor maintenance table PM00200
            var row = Dynamics.Tables.PmVendorMstr;

            // define a key and set the key's value
            row.Key = 1;
            row.VendorId.Value = "0000TESTVENDOR";
            
            // set record to update, check for errors, and make changes
            error = row.Change();
            if (error == TableError.NoError)
            {
                // remove the record, and check for errors
                if (row.Remove() != TableError.NoError)
                    MessageBox.Show(error.ToString());
            }
            else
            {
                MessageBox.Show(error.ToString());
            }

            // close the table since this is a global table open
            row.Close();

        }

    }
}
