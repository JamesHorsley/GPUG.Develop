using System;
using System.Data;
using System.Collections.Generic;
using System.Reflection;
using System.ComponentModel;
using System.Diagnostics;

using Microsoft.Dexterity.Bridge;
using Microsoft.Dexterity.Applications;
using Microsoft.Dexterity.Applications.DynamicsDictionary;


namespace GPUG.Develop.Integrations
{
    /// <summary>
    /// Having an addin in a different namespace will still load into GP
    /// </summary>
    public class GeneralLedgerAddIn : IDexterityAddIn
    {
        private Repository Database;

        /// <summary>
        /// Constructor
        /// </summary>
        public GeneralLedgerAddIn()
        {
            this.Database = new GeneralLedgerAddIn.Repository();
        }

        /// <summary>
        /// Called by GP to initialize the addin
        /// A required method for the IDexterityAddIn interface
        /// </summary>
        public void Initialize()
        {
            Dynamics.Forms.AboutBox.AddMenuHandler(Test, "Dynamics GP Test");

            Dynamics.Forms.Login.Login.Password.Value = "Pass@word1";
        }


        private void Test(object sender, EventArgs e)
        {
            var accts = this.Database.GetAccountRange("000-1100-00", "000-1301-00");
            var acct = this.Database.GetAccount("000-1100-00");

            var csData = "000-1100-00".ToCompositeAccount();
        }



        /// <summary>
        /// Nested class to manage access to the database
        /// Contains the methods needed to interact with the database objects
        /// </summary>
        private class Repository
        {
            // point references to the GP tables needed for the database connections
            private GlAccountMstrTable GL00100 = Dynamics.Tables.GlAccountMstr;
            private GlAccountCategoryMstrTable GL00102 = Dynamics.Tables.GlAccountCategoryMstr;
            private GlAccountIndexMstrTable GL00105 = Dynamics.Tables.GlAccountIndexMstr;
            private TableError TableError;


            /// <summary>
            /// Constructor
            /// </summary>
            public Repository()
            {

            }


            /// <summary>
            /// Get an account index from an account number
            /// </summary>
            public int GetAccountIndex(string accountNumberString)
            {
                var index = 0;

                try
                {
                    // set the required key for the table
                    this.GL00105.Key = (short)GL00105Keys.NumberString;
                    this.GL00105.AccountNumberString.Value = accountNumberString;

                    // get a data row
                    this.TableError = this.GL00105.Get();

                    // check for an error, and set the return value, or return the error
                    if (this.TableError == TableError.NoError)
                    {
                        index = this.GL00105.AccountIndex;
                    }
                    else
                    {
                        throw new Exception(this.TableError.ToString());
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Error encountered: " + ex.Message);
                }
                finally
                {
                    // close the connection to the table accessed
                    // if not, users will receive a system processes are running message when attempting to close GP
                    this.GL00105.Close();
                }

                return index;
            }


            /// <summary>
            /// Get an account index from an account number
            /// </summary>
            public List<int> GetAccountIndexRange(string accountNumberStart, string accountNumberEnd)
            {
                var index = new List<int>();

                try
                {
                    // set the required key for the table
                    this.GL00105.Key = (short)GL00105Keys.NumberString;


                    // set the start of the range
                    this.GL00105.Clear();
                    this.GL00105.AccountNumberString.Value = accountNumberStart;
                    this.GL00105.RangeStart();


                    // set the end of the range
                    this.GL00105.Clear();
                    this.GL00105.AccountNumberString.Value = accountNumberEnd;
                    this.GL00105.RangeEnd();


                    // get the first data row
                    this.TableError = this.GL00105.GetFirst();


                    // check for an error, and set the return value, or return the error
                    if (this.TableError == TableError.NoError)
                    {
                        index.Add(this.GL00105.AccountIndex.Value);
                    }
                    else
                    {
                        throw new Exception(this.TableError.ToString());
                    }


                    // loop through the remaining rows in the range, and add them to the dataset
                    while (this.TableError != TableError.EndOfTable)
                    {
                        // get the next record
                        this.TableError = this.GL00105.GetNext();

                        switch (this.TableError)
                        {
                            case TableError.NoError:
                                index.Add(this.GL00105.AccountIndex.Value);
                                break;

                            case TableError.EndOfTable:
                                break;

                            default:
                                throw new Exception(this.TableError.ToString());
                        }

                    };
                }
                catch (Exception ex)
                {
                    throw new Exception("Error encountered: " + ex.Message);
                }
                finally
                {
                    // close the connection to the table accessed
                    // if not, users will receive a system processes are running message when attempting to close GP
                    this.GL00105.Close();
                }

                return index;
            }


            /// <summary>
            /// Get an account index from an account number
            /// </summary>
            public DataTable GetAccount(string accountNumberString)
            {
                var result = GLHelper.CreateBlankGLTable();

                // get the account index for the account master key requirement
                var accountIndex = this.GetAccountIndex(accountNumberString);

                // if there is an error, return null
                if (this.TableError != TableError.NoError)
                    return null;

                try
                {
                    // set the required key for the table
                    this.GL00100.Key = (short)GL00100Keys.Index;
                    this.GL00100.AccountIndex.Value = accountIndex;

                    // get the data
                    this.TableError = this.GL00100.Get();

                    // check for an error, and set the return value, or return the error
                    if (this.TableError == TableError.NoError)
                    {
                        // add the data row using a mapping helper
                        result.AddGLRow(this.GL00100);
                    }
                    else
                    {
                        throw new Exception(this.TableError.ToString());
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Error encountered: " + ex.Message);
                }
                finally
                {
                    // close the connection to the table accessed
                    // if not, users will receive a system processes are running message when attempting to close GP
                    this.GL00100.Close();
                }

                return result;
            }


            /// <summary>
            /// Get an account index from an account number
            /// </summary>
            public DataTable GetAccountRange(string accountNumberStart, string accountNumberEnd)
            {
                var result = GLHelper.CreateBlankGLTable();
                var indexList = this.GetAccountIndexRange(accountNumberStart, accountNumberEnd);


                try
                {
                    // for each account integer in the index list
                    // get the account data row, and map to the GL data table
                    foreach (int index in indexList)
                    {
                        // set the required key for the table
                        this.GL00100.Key = (short)GL00100Keys.Index;
                        this.GL00100.AccountIndex.Value = index;

                        // get the data
                        this.TableError = this.GL00100.Get();

                        // check for an error, and set the return value, or return the error
                        if (this.TableError == TableError.NoError)
                        {
                            // add the data row using a mapping helper
                            result.AddGLRow(this.GL00100);
                        }
                        else
                        {
                            throw new Exception(this.TableError.ToString());
                        }
                    }

                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                finally
                {
                    // close the connection to the table accessed
                    // if not, users will receive a system processes are running message when attempting to close GP
                    this.GL00100.Close();
                }

                return result;
            }

        }

        public enum GL00100Keys
        {
            Index = 1,
            AliasAndIndex = 2,
            DescriptionAndNumber = 3,
            TypeAndNumber = 4,
            CategoryAndMainSegmentAndNumber = 5,
            Number = 6,
            MainSegmentAndNumber = 7
        }

        public enum GL00102Keys
        {
            Number = 1,
            DescriptionAndNumber = 2
        }

        public enum GL00105Keys
        {
            Index = 1,
            Number = 2,
            NumberString = 3
        }

    }

    /// <summary>
    /// Helpers and conversion methods
    /// </summary>
    public static class GLHelper
    {
        /// <summary>
        /// Converts GL composite data to an account string value
        /// </summary>
        public static string CompositeToString(this AccountNumberCompositeData accountCS)
        {
            // get the account composite value and convert it to a string
            string accountNumber = string.Empty;
            for (int i = 0; i < accountCS.Length; i++)
            {
                // check to make sure theere are segment values in the composite value
                if (!string.IsNullOrEmpty(accountCS[i]))
                {
                    // add the account separator only if the segment is not the first one
                    // TODO: pull the configuration for the account separator 
                    if (i > 0)
                    {
                        accountNumber += "-";
                    }

                    // add the account segment value
                    accountNumber += accountCS[i];
                }
            }

            return accountNumber;
        }


        /// <summary>
        /// Converts a GL account string to a composite data object
        /// </summary>
        public static AccountNumberCompositeReadOnly ToCompositeAccount(this string accountCString)
        {
            var accountCs = new AccountNumberCompositeData();

            // AccountNumberCompositeReadOnly acctReadOnly;

            //TODO: get the setup delimiter from company setup
            // split the string by the delimiter
            char delimiter = '-';

            // split the segments
            var acctSegments = accountCString.Split(delimiter);

            // add each segment to the composite data
            for (int i = 0; i < acctSegments.Length; i++)
            {
                accountCs[i] = acctSegments[i];
            }

            //// create the readonly composite data
            //acctReadOnly = new AccountNumberCompositeReadOnly(accountCs);

            // Microsoft.Dexterity.Applications.DynamicsDictionary.AccountNumberCompositeData accountNumberData = new AccountNumberCompositeData();
            //accountNumberData[0] = "000";
            //accountNumberData[1] = "1100";
            //accountNumberData[2] = "00";
            //accountNumberData[3] = "";

            // Create the read-only composite that is used for the function call
            var accountNumber = new AccountNumberCompositeReadOnly(accountCs);

            // Convert the account number to its account alias
            string alias = Dynamics.Forms.GlAcctBase.Functions.ConvertAcctToAliasStr.Invoke(accountNumber);

            return accountNumber;
        }


        /// <summary>
        /// Returns a blank GL datatable
        /// </summary>
        public static DataTable CreateBlankGLTable()
        {
            var table = new DataTable();

            table.Columns.Add("AccountAlias", typeof(string));
            table.Columns.Add("AccountCategoryNumber", typeof(short));
            table.Columns.Add("AccountDescription", typeof(string));
            table.Columns.Add("AccountIndex", typeof(int));
            table.Columns.Add("AccountNumber", typeof(string));
            table.Columns.Add("AccountType", typeof(short));
            table.Columns.Add("Active", typeof(bool));
            table.Columns.Add("AdjustForInflation", typeof(bool));
            table.Columns.Add("AllowAccountEntry", typeof(bool));
            table.Columns.Add("BalanceForCalculation", typeof(short));
            table.Columns.Add("ClearBalance", typeof(bool));
            table.Columns.Add("ConversionMethod", typeof(short));
            table.Columns.Add("CreatedDate", typeof(DateTime));
            table.Columns.Add("DecimalPlaces", typeof(short));
            table.Columns.Add("DisplayInLookups", typeof(int));
            table.Columns.Add("FixedOrVariable", typeof(short));
            table.Columns.Add("HistoricalRate", typeof(decimal));
            table.Columns.Add("InflationEquityAccountIndex", typeof(int));
            table.Columns.Add("InflationRevenueAccountIndex", typeof(int));
            table.Columns.Add("MainAccountSegment", typeof(string));
            table.Columns.Add("ModifiedDate", typeof(DateTime));
            table.Columns.Add("NoteIndex", typeof(decimal));
            table.Columns.Add("PostingType", typeof(short));
            table.Columns.Add("PostInventoryIn", typeof(short));
            table.Columns.Add("PostPayrollIn", typeof(short));
            table.Columns.Add("PostPurchasingIn", typeof(short));
            table.Columns.Add("PostSalesIn", typeof(short));
            table.Columns.Add("TypicalBalance", typeof(short));
            table.Columns.Add("UserDefined1", typeof(string));
            table.Columns.Add("UserDefined2", typeof(string));
            table.Columns.Add("UserDefinedString1", typeof(string));
            table.Columns.Add("UserDefinedString2", typeof(string));

            //var props = this.GL00100.GetType().GetProperties();
            //foreach (PropertyInfo p in props)
            //{
            //    string name = p.Name;
            //}

            return table;
        }


        /// <summary>
        /// Mapper for the GL accounts table
        /// </summary>
        public static DataTable AddGLRow(this DataTable glData, GlAccountMstrTable dataRow)
        {
            if (dataRow != null)
            {
                // Add the row to the datatable
                var row = glData.NewRow();

                // map the values to the datatable
                row["AccountAlias"] = dataRow.AccountAlias.Value;
                row["AccountCategoryNumber"] = dataRow.AccountCategoryNumber.Value;
                row["AccountDescription"] = dataRow.AccountDescription.Value;
                row["AccountIndex"] = dataRow.AccountIndex.Value;
                row["AccountNumber"] = dataRow.AccountNumber.Value.CompositeToString();
                row["AccountType"] = dataRow.AccountType.Value;
                row["Active"] = dataRow.Active.Value;
                row["AdjustForInflation"] = dataRow.AdjustForInflation.Value;
                row["AllowAccountEntry"] = dataRow.AllowAccountEntry.Value;
                row["BalanceForCalculation"] = dataRow.BalanceForCalculation.Value;
                row["ClearBalance"] = dataRow.ClearBalance.Value;
                row["ConversionMethod"] = dataRow.ConversionMethod.Value;
                row["CreatedDate"] = dataRow.CreatedDate.Value;
                row["DecimalPlaces"] = dataRow.DecimalPlaces.Value;
                row["DisplayInLookups"] = dataRow.DisplayInLookups.Value;
                row["FixedOrVariable"] = dataRow.FixedOrVariable.Value;
                row["HistoricalRate"] = dataRow.HistoricalRate.Value;
                row["InflationEquityAccountIndex"] = dataRow.InflationEquityAccountIndex.Value;
                row["InflationRevenueAccountIndex"] = dataRow.InflationRevenueAccountIndex.Value;
                row["MainAccountSegment"] = dataRow.MainAccountSegment.Value;
                row["ModifiedDate"] = dataRow.ModifiedDate.Value;
                row["NoteIndex"] = dataRow.NoteIndex.Value;
                row["PostingType"] = dataRow.PostingType.Value;
                row["PostInventoryIn"] = dataRow.PostInventoryIn.Value;
                row["PostPayrollIn"] = dataRow.PostPayrollIn.Value;
                row["PostPurchasingIn"] = dataRow.PostPurchasingIn.Value;
                row["PostSalesIn"] = dataRow.PostSalesIn.Value;
                row["TypicalBalance"] = dataRow.TypicalBalance.Value;
                row["UserDefined1"] = dataRow.UserDefined1.Value;
                row["UserDefined2"] = dataRow.UserDefined2.Value;
                row["UserDefinedString1"] = dataRow.UserDefinedString1.Value;
                row["UserDefinedString2"] = dataRow.UserDefinedString2.Value;

                glData.Rows.Add(row);
            }

            return glData;
        }

    }
}
