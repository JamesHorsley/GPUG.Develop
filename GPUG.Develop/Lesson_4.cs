using System;
using Microsoft.Dexterity.Bridge;
using Microsoft.Dexterity.Applications;
using Microsoft.Dexterity.Applications.DynamicsDictionary;
using System.Text.RegularExpressions;

namespace GPUG.Develop
{
    /// <summary>
    /// To use this Add-in
    /// 
    /// In the vendor maintenance form, the user will type in the name of the vendor they wish to create.
    /// Once the user leaves the name field, the add-in will pull the next available vendor id using an 
    /// alpha-numeric structure (XXX######) where the X = the first three characters of the vendor name 
    /// without any special characters. The # = an auto number that is sequencial, and will automatically
    /// populate the last number found + 1
    /// 
    /// The vendor id will automatically populate with this information, and allow the user to finish
    /// filling in the remaining fields to create the vendor record.
    /// 
    /// Created by Joshua Pelkola, BKD Technologies
    /// Updated 10/2/2016
    /// jpelkola@bkd.com
    /// 
    /// </summary>
    public class Lesson_4 : IDexterityAddIn
    {
        // reference the vendor maintenance form
        private PmVendorMaintenanceForm vendorForm = Dynamics.Forms.PmVendorMaintenance;

        // required by the IDexterityAddIn interface
        public void Initialize()
        {
            // make a reference to the vendor name's after change event
            this.vendorForm.PmVendorMaintenance.VendorName.LeaveAfterOriginal += new EventHandler(VendorName_LeaveAfterOriginal);
        }

        // local parameters to be used for creating the next vendor number
        private short _numberAlphaCharacters = 5;
        private char _alphaFillerCharacter = '0';
        private short _numberDigits = 5;

        /// <summary>
        /// handles the vendor name change and setting the vendor id if a new record.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VendorName_LeaveAfterOriginal(object sender, EventArgs e)
        {
            try
            {
                // check to see if the vendor id is already populated, and cancel if it is
                if (!string.IsNullOrEmpty(this.vendorForm.PmVendorMaintenance.VendorId.Value))
                {
                    return;
                }

                // get the vendor name
                var vendorName = this.vendorForm.PmVendorMaintenance.VendorName.Value;

                // get the alpha code
                var alphaString = this.FormatVendorIdString(vendorName, this._numberAlphaCharacters, this._alphaFillerCharacter);

                // get the new vendor number, and assign it to the vendor id field
                var vendorNumber = this.GetNextVendorNumber(alphaString, this._numberDigits);

                // assign, and tell the window to validate the vendor id
                if (vendorNumber.Length <= 15)
                {
                    this.vendorForm.PmVendorMaintenance.VendorId.Value = vendorNumber;
                    this.vendorForm.PmVendorMaintenance.VendorId.RunValidate();
                }
                else
                {
                    throw new Exception("The vendor id is too long. Please check the parameters for the add-in to create one.");
                }
            }
            catch (Exception ex)
            {
                Dynamics.Forms.SyVisualStudioHelper.Functions.DexWarning.Invoke(ex.Message);
            }
        }

        /// <summary>
        /// formats the given vendor name into a leading alpha character set to use for a new vendor id
        /// </summary>
        /// <param name="vendorName">Unformatted vendor name to use as basis for getting the alpha codes</param>
        /// <param name="numCharacters">the number of characters to pull from the beginning of the vendor name</param>
        /// <param name="fillerCharacter">the character to be used when the vendor name is shorter than the numCharacters</param>
        /// <returns></returns>
        private string FormatVendorIdString(string vendorName, short numCharacters, char fillerCharacter)
        {
            var vendorAlphaChars = string.Empty;

            try
            {
                // define a regular expression to remove all non-alpha characters, and convert to upper case
                var regex = new Regex("[^A-Za-z']");
                vendorAlphaChars = regex.Replace(vendorName.ToUpper(), "");

                // format the length if greater than number of characters
                if (vendorAlphaChars.Length > numCharacters)
                {
                    vendorAlphaChars = vendorAlphaChars.Substring(0, numCharacters);
                }

                // format the length if less than number of characters
                if (vendorAlphaChars.Length < numCharacters)
                {
                    // add the suffix characters
                    for (int i = 0; i < numCharacters; i++)
                    {
                        vendorAlphaChars += fillerCharacter;
                    }

                    // get the alpha code from the updated string
                    vendorAlphaChars = vendorAlphaChars.Substring(0, numCharacters);
                }
            }
            catch (Exception ex)
            {
                Dynamics.Forms.SyVisualStudioHelper.Functions.DexWarning.Invoke(ex.Message);
            }

            return vendorAlphaChars;
        }

        /// <summary>
        /// Access vendor global table buffer to find the next available vendor id in the GP company
        /// </summary>
        /// <param name="alphaString">Formatted Alpha string to prefix the vendor id's number</param>
        /// <param name="numDigits">The number of numeric digits to follow the alpha string</param>
        /// <returns>A unique vendor number</returns>
        private string GetNextVendorNumber(string alphaString, short numDigits)
        {
            // set the default digit string 
            string digitString = this.AddLeadingZeros(0, numDigits);

            // set the return vendor number
            var vendorNumber = alphaString + digitString;

            // set the upper limit to check for all vendors in a range using the alpha codes
            var vendorNumberUpperLimit = alphaString + this.AddLeadingZeros(9, numDigits);

            try
            {
                // check vendors and see if there is a matching alpha vendor to iterate the next number if available
                TableError err;
                var vendorTable = Dynamics.Tables.PmVendorMstr;
                vendorTable.Key = 1;

                // check the range of vendors that start and stop the possible ranges of new vendor numeric ids
                // setting the beginning of the range as alpha string + 0000000 to the number of digits listed
                // setting the end of the range as the alpha string + 999999 to pull all possible instances listed
                vendorTable.Clear();
                vendorTable.VendorId.Value = vendorNumber;
                vendorTable.RangeStart();

                vendorTable.Clear();
                vendorTable.VendorId.Value = vendorNumberUpperLimit;
                vendorTable.RangeEnd();

                // get the last record
                err = vendorTable.GetLast();

                // check for the record
                // if no error, there is a record
                if (err == TableError.NoError)
                {
                    string id = vendorTable.VendorId.Value;

                    // check the length of the id
                    if (id.Length >= numDigits)
                    {
                        // get the number digits from the right
                        var numString = id.Substring(id.Length - numDigits);

                        // try and convert the string to a number
                        int num = -1;
                        var ok = int.TryParse(numString, out num);

                        // if the try parse is good, add a digit
                        if (ok)
                        {
                            num += 1;
                        }

                        // convert the number back to a string, and assign the new value to the vendor number
                        numString = this.AddLeadingZeros(num, numDigits);
                        vendorNumber = alphaString + numString;
                    }
                }
                else // there is no record, use the default created at the top of the method.
                {
                    // no record found, let the add-in use the default
                }

                vendorTable.Close();
            }
            catch (Exception ex)
            {
                Dynamics.Forms.SyVisualStudioHelper.Functions.DexWarning.Invoke(ex.Message);
            }

            return vendorNumber;
        }

        /// <summary>
        /// Converts a number to a string, and adds leading zeros to it
        /// </summary>
        /// <param name="number">The number preceded by zeros to be formatted as a string</param>
        /// <param name="numDigits">The number of numeric digits to follow the alpha string</param>
        /// <returns>a string value for the number passed in with leading zeros</returns>
        private string AddLeadingZeros(int number, short numDigits)
        {
            // set the default digit string 
            string digitString = string.Empty;
            for (int i = 0; i < numDigits; i++)
            {
                digitString += "0";
            }

            // add the number and shorten the string from the right
            digitString += number.ToString();
            digitString = digitString.Substring(digitString.Length - numDigits);

            return digitString;
        }
    }
}
