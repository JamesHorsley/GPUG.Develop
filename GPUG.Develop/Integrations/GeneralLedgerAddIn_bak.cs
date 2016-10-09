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
    public class GeneralLedgerAddIn_bak //: IDexterityAddIn
    {
        /*
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
                var result = this.CreateBlankGLTable();

                // get the account index for the account master key requirement
                var accountIndex = this.GetAccountIndex(accountNumberString);

                // if there is an error, return null
                if (this.TableError != TableError.NoError) return null;

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
                        // Add the row to the datatable
                        var row = result.NewRow();

                        // map the values to the datatable
                        row["AccountAlias"] = this.GL00100.AccountAlias.Value;
                        row["AccountCategoryNumber"] = this.GL00100.AccountCategoryNumber.Value;
                        row["AccountDescription"] = this.GL00100.AccountDescription.Value;
                        row["AccountIndex"] = this.GL00100.AccountIndex.Value;
                        row["AccountNumber"] = this.GL00100.AccountNumber.Value.CompositeToString();
                        row["AccountType"] = this.GL00100.AccountType.Value;
                        row["Active"] = this.GL00100.Active.Value;
                        row["AdjustForInflation"] = this.GL00100.AdjustForInflation.Value;
                        row["AllowAccountEntry"] = this.GL00100.AllowAccountEntry.Value;
                        row["BalanceForCalculation"] = this.GL00100.BalanceForCalculation.Value;
                        row["ClearBalance"] = this.GL00100.ClearBalance.Value;
                        row["ConversionMethod"] = this.GL00100.ConversionMethod.Value;
                        row["CreatedDate"] = this.GL00100.CreatedDate.Value;
                        row["DecimalPlaces"] = this.GL00100.DecimalPlaces.Value;
                        row["DisplayInLookups"] = this.GL00100.DisplayInLookups.Value;
                        row["FixedOrVariable"] = this.GL00100.FixedOrVariable.Value;
                        row["HistoricalRate"] = this.GL00100.HistoricalRate.Value;
                        row["InflationEquityAccountIndex"] = this.GL00100.InflationEquityAccountIndex.Value;
                        row["InflationRevenueAccountIndex"] = this.GL00100.InflationRevenueAccountIndex.Value;
                        row["MainAccountSegment"] = this.GL00100.MainAccountSegment.Value;
                        row["ModifiedDate"] = this.GL00100.ModifiedDate.Value;
                        row["NoteIndex"] = this.GL00100.NoteIndex.Value;
                        row["PostingType"] = this.GL00100.PostingType.Value;
                        row["PostInventoryIn"] = this.GL00100.PostInventoryIn.Value;
                        row["PostPayrollIn"] = this.GL00100.PostPayrollIn.Value;
                        row["PostPurchasingIn"] = this.GL00100.PostPurchasingIn.Value;
                        row["PostSalesIn"] = this.GL00100.PostSalesIn.Value;
                        row["TypicalBalance"] = this.GL00100.TypicalBalance.Value;
                        row["UserDefined1"] = this.GL00100.UserDefined1.Value;
                        row["UserDefined2"] = this.GL00100.UserDefined2.Value;
                        row["UserDefinedString1"] = this.GL00100.UserDefinedString1.Value;
                        row["UserDefinedString2"] = this.GL00100.UserDefinedString2.Value;

                        result.Rows.Add(row);
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
                var result = this.CreateBlankGLTable();
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
                            // Add the row to the datatable
                            var row = result.NewRow();


                            // map the values to the datatable
                            row["AccountAlias"] = this.GL00100.AccountAlias.Value;
                            row["AccountCategoryNumber"] = this.GL00100.AccountCategoryNumber.Value;
                            row["AccountDescription"] = this.GL00100.AccountDescription.Value;
                            row["AccountIndex"] = this.GL00100.AccountIndex.Value;
                            row["AccountNumber"] = this.GL00100.AccountNumber.Value.CompositeToString();
                            row["AccountType"] = this.GL00100.AccountType.Value;
                            row["Active"] = this.GL00100.Active.Value;
                            row["AdjustForInflation"] = this.GL00100.AdjustForInflation.Value;
                            row["AllowAccountEntry"] = this.GL00100.AllowAccountEntry.Value;
                            row["BalanceForCalculation"] = this.GL00100.BalanceForCalculation.Value;
                            row["ClearBalance"] = this.GL00100.ClearBalance.Value;
                            row["ConversionMethod"] = this.GL00100.ConversionMethod.Value;
                            row["CreatedDate"] = this.GL00100.CreatedDate.Value;
                            row["DecimalPlaces"] = this.GL00100.DecimalPlaces.Value;
                            row["DisplayInLookups"] = this.GL00100.DisplayInLookups.Value;
                            row["FixedOrVariable"] = this.GL00100.FixedOrVariable.Value;
                            row["HistoricalRate"] = this.GL00100.HistoricalRate.Value;
                            row["InflationEquityAccountIndex"] = this.GL00100.InflationEquityAccountIndex.Value;
                            row["InflationRevenueAccountIndex"] = this.GL00100.InflationRevenueAccountIndex.Value;
                            row["MainAccountSegment"] = this.GL00100.MainAccountSegment.Value;
                            row["ModifiedDate"] = this.GL00100.ModifiedDate.Value;
                            row["NoteIndex"] = this.GL00100.NoteIndex.Value;
                            row["PostingType"] = this.GL00100.PostingType.Value;
                            row["PostInventoryIn"] = this.GL00100.PostInventoryIn.Value;
                            row["PostPayrollIn"] = this.GL00100.PostPayrollIn.Value;
                            row["PostPurchasingIn"] = this.GL00100.PostPurchasingIn.Value;
                            row["PostSalesIn"] = this.GL00100.PostSalesIn.Value;
                            row["TypicalBalance"] = this.GL00100.TypicalBalance.Value;
                            row["UserDefined1"] = this.GL00100.UserDefined1.Value;
                            row["UserDefined2"] = this.GL00100.UserDefined2.Value;
                            row["UserDefinedString1"] = this.GL00100.UserDefinedString1.Value;
                            row["UserDefinedString2"] = this.GL00100.UserDefinedString2.Value;

                            result.Rows.Add(row);
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


            /// <summary>
            /// Returns a blank GL datatable
            /// </summary>
            public DataTable CreateBlankGLTable()
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
        public static AccountNumberCompositeData ToCompositeAccount(this string accountCString)
        {
            var accountCs = new AccountNumberCompositeData();

            try
            {
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

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }

            return accountCs;
        }
        */

    }
}
