// V004 20110207    * completed
// V005 20110207    * maximum row is now set to 10000
//                  * we will look for the word End to end our line processing
// V006 20110212    * don't submit lines that have a 0 value
//                  * truncate the decimals to two decimals
// V007 20120512    * updated error checking to something a tad more sensible
//                      so it will work with LSO 10
//      20120513    * addressed an issue with "Old format or invalid type library"
//                      http://support.microsoft.com/kb/320369
//                      issue raised by Hendrik
//                      http://potatoit.wordpress.com/2011/01/23/journal-importing-jscript-handling-of-error-messages/#comment-45
// V008 20120516    * if a forumla is used in the debit column and it is zero, 
//                      then we incorrectly don't process the credit column
// V009 20120516    * we stopped on blank lines when we should continue processing
// V010 20120527    * improved error reporting - if there is an error we display a dialog box which
//                     moves the focus to the script where the user can select to save the error to the
//                     spreadsheet or to end the import.  We don't create the confusion with the
//                     focus disappearing


import System;
import System.Text;
import System.Windows;
import System.Windows.Controls;
import System.Windows.Media;
import System.Windows.Media.Media3D;

import MForms;
import Mango.UI.Core;
import Mango.UI.Core.Util;
import Mango.UI.Services;
import Mango.Services;

import Excel;
import System.Reflection;

package MForms.JScript
{
	class GLS100_JournalImport_V10
	{
		var giicInstanceController : IInstanceController = null;    // this is where we will store the IInstanceController to make it available to the rest of the class
		var ggrdContentGrid : Grid = null;                          // this is the Grid that we get passed by the Init()
		var gexaApplication = null;                                 // Excel.Application object
		
		var gbtnImportFromExcel : Button = null;                    // this is the button that we will put on to the panel that will kick off the whole import
		var glvListView : ListView = null;                          // this is the ListView on the panel
		
		var gwbWorkbook = null;                                     // here we will store the Workbook object
		
		var giStartRow : int = 15;                                  // the starting row in the Spreadsheet
		var giMaxRow : int = 10000;                                    // the end row in the Spreadsheet
		var giCurrentRow : int = 15;                                // the current row in the Spreadsheet
		
		var gbLookForResponse = false;                              // should we be looking for a response?

		var gobjStatusJ1 = null;                                      // the statusbar
		var gobjStatusE = null;

		var gbRequest : boolean = false;                            // the request event 


		var gstrVoucherType : String = null;                        // the voucher type

        var gstrError : String = null;                              // keep track of the errors

		public function Init(element: Object, args: Object, controller : Object, debug : Object)
		{
			// lets make some of the controls and other
			// bits pieces available to other sections of our code
			ggrdContentGrid = controller.RenderEngine.Content;
			giicInstanceController = controller;
			glvListView = controller.RenderEngine.ListControl.ListView;
			
			try
			{
				// create the button for importing
				gbtnImportFromExcel = new Button();
				gbtnImportFromExcel.Content = "Import";

				Grid.SetColumnSpan(gbtnImportFromExcel, 10);
				Grid.SetColumn(gbtnImportFromExcel, 1);
				Grid.SetRow(gbtnImportFromExcel, 22);
				
				// finally add the control to the grid
				ggrdContentGrid.Children.Add(gbtnImportFromExcel);
				
				// ----- Events -----
				gbtnImportFromExcel.add_Click(OnImportFromExcelClicked);
				gbtnImportFromExcel.add_Unloaded(OnImportFromExcelUnloaded);

			}
			catch(exException)
			{
				MessageBox.Show("Error: " + exException.Message + Environment.NewLine + exException.StackTrace);
			}
		
		}
		
		// check for errors, we need to check
		// for errors on BOTH the E panel
		// and J1 panel
		// We will go out and look for the status control if we 
		// don't have it
		private function checkForError()
		{
			var strResult : String = null;
            // var objRuntime = MForms.Runtime.Runtime(giicInstanceController.Runtime).Result;
			var strStatusMessage : String = MForms.Runtime.Runtime(giicInstanceController.Runtime).Result;  // objRuntime.Result;

			try
			{
				var iStartPosition : int = 0;

				iStartPosition = strStatusMessage.IndexOf("<Msg>");

				// if Msg doesn't exist, then we didn't have an error
				if(-1 == iStartPosition)
				{
					// we are all good!
				}
				else
				{
					var iEndPosition : int = strStatusMessage.IndexOf("</Msg>");
					if((-1 == iEndPosition) && (0 != iEndPosition))
					{
						iEndPosition = strStatusMessage.length-1;
					}
					strResult = strStatusMessage.substring(iStartPosition+5, iEndPosition);

                    // 20120527 V010
                    gstrError = "Row: " + giCurrentRow + " error: " + strResult;
				}
			}
			catch(ex)
			{
				MessageBox.Show("checkForError() exception: " + ex.message);
			}

			return(strResult);
		}

		// display an OpenFileDialog box
		// and extract the result
		private function retrieveImportFile()
		{
			var result : String = null;
			var ofdFile = new System.Windows.Forms.OpenFileDialog();    // we have to use the forms OpenFileDialog unfortunately
			if(null != ofdFile)
			{
				ofdFile.Multiselect = false;
				ofdFile.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx|All Files (*.*)|*.*"; // filter on xls or xlsx files only
				
				if(true == ofdFile.ShowDialog())
				{
					result = ofdFile.FileName;
				}
			}
			return(result);
		}
		
		// this is where we actually do the import
		private function OnImportFromExcelClicked(sender : Object, e : RoutedEventArgs)
		{
			gstrVoucherType = null;
			try
			{
				// here we do some initialisation of Excel
				InitialiseExcel();

				var strFilename : String = retrieveImportFile();            // retrieve the filename of the Excel spreadsheet to open
				if((null != strFilename) && (null != gexaApplication))      // ensure that not only do we have a filename, but we also managed to initialise Excel
				{
					gwbWorkbook = gexaApplication.Workbooks.Open(strFilename);  // open the spreadsheet
					if(null != gwbWorkbook)
					{
						giicInstanceController.add_RequestCompleted(OnRequestCompleted);
						giicInstanceController.add_RequestCompleted(OnRequested);
						gbRequest = true;

						gwbWorkbook.Saved = true;                               // get rid of those annoying save messages
						var strVoucherType : String = retrieveVoucherType();    // we want to get the voucher type from the spreadsheet (the GLS100 voucher)
						
						if(!String.IsNullOrEmpty(strVoucherType))               // we need to ensure that we have a voucher type
						{
							giicInstanceController.RenderEngine.SetFocusOnList();
							selectFAMFunction(strVoucherType);              // now we need to go out and select the function
							
							// from where on out, we start using the events
						}
					}
					else MessageBox.Show("Failed to Open Workbook");
				}
				else MessageBox.Show("Filename or Excel doesn't exist: " + strFilename);
			}
			catch(exException)
			{
				MessageBox.Show("Error: " + exException.description);
			}
		}

		// set the VoucherText within GLS100/E
		private function setM3VoucherText(astrVoucherText : String)
		{
			var tbVoucherText : TextBox = ScriptUtil.FindChild(ggrdContentGrid, "WWGVTX");
			if(null != tbVoucherText)
			{
				tbVoucherText.Text = astrVoucherText;
			}
			else MessageBox.Show("setM3VoucherText() - Child not found");
		}

		// set the Year within GLS100/E
		private function setM3YEA4(astrYEA4Text : String)
		{
			var tbYEA4Text : TextBox = ScriptUtil.FindChild(ggrdContentGrid, "WWYEA4");
			if(null != tbYEA4Text)
			{
				tbYEA4Text.Text = astrYEA4Text;
			}
			else MessageBox.Show("setM3VoucherText() - Child not found");
		}

		// we need to set the reversal date
		private function setM3ReversalDate(astrReversalText : String)
		{
			var tbReversalText = ScriptUtil.FindChild(ggrdContentGrid, "WWSHDT");
			if(null != tbReversalText)
			{
				try
				{
					var dtValue : DateTime = DateTime.FromOADate(Convert.ToDouble(astrReversalText));
					tbReversalText.Value = dtValue;
				}
				catch(ex)
				{
					MessageBox.Show(ex.description);
				}
			}
			else MessageBox.Show("setM3ReversalDate() - Child not found");
		}

		// set the accounting date within GLS100/E
		private function setM3AccountingDate(astrAccountingDate : String)
		{
			var tbITNO = ScriptUtil.FindChild(ggrdContentGrid, "WWACDT");
			if(null != tbITNO)
			{
				try
				{
					var dtValue : DateTime = DateTime.FromOADate(Convert.ToDouble(astrAccountingDate));
					tbITNO.Value = dtValue;
				}
				catch(ex)
				{
					MessageBox.Show(ex.description);
				}
			}
			else MessageBox.Show("Accounting Date not found");
		}

		// retrieve the voucher text from the spreadsheet
		private function retrieveVoucherText()
		{
			return(gwbWorkbook.ActiveSheet.Range("E7").Value);
		}
		
		// retrieve the accounting date from the spreadsheet
		// we need to use Value2 in this instance to get
		// a value that we can actually use
		private function retrieveAccountingDate()
		{
			try
			{
				return(gwbWorkbook.ActiveSheet.Range("R5").Value2);
			}
			catch(ex)
			{
				MessageBox.Show("Exception: " + ex.message);
			}
			
		}

		// retrieve the voucher type from the Spreadsheet
		private function retrieveVoucherType()
		{
			gstrVoucherType = gwbWorkbook.ActiveSheet.Range("K5").Value;
			return(gstrVoucherType);
		}		

		// retrieve the reversing date
		private function retrieveReversingDate()
		{
			try
			{
				return(gwbWorkbook.ActiveSheet.Range("R7").Value2);
			}
			catch(ex)
			{
			}
		}		


		// GLS100/B set the FAM Function
		private function selectFAMFunction(astrFAMFunction : String)
		{
			var bFound : boolean = false;
			
			if(!String.IsNullOrEmpty(astrFAMFunction))
			{
				// search through the ListView for the FAM function
				for(var iCount : int = 0; iCount < glvListView.Items.Count; iCount++)
				{
					var itmCurrentItem = glvListView.Items[iCount];
					if(null != itmCurrentItem)
					{
						if(!String.IsNullOrEmpty(itmCurrentItem[0]))
						{
							var strCurrentString = itmCurrentItem[0].ToString();
							if(0 == String.Compare(strCurrentString, astrFAMFunction))
							{
								glvListView.SelectedItem = itmCurrentItem;
								bFound = true;
								break;
							}
						}
					}
				}
			}
			if(true == bFound)
			{
				// ok, we've found the FAM Function on the ListView
				// now we need to SELECT it
				giicInstanceController.ListOption("1");	// SELECT
			}
			
		}
		
		// our Import button is being unloaded, now's a good time to clean
		// everything up
		private function OnImportFromExcelUnloaded(sender : Object, e : RoutedEventArgs)
		{
			if(null != gbtnImportFromExcel)
			{
				gbtnImportFromExcel.remove_Click(OnImportFromExcelClicked);
				gbtnImportFromExcel.remove_Unloaded(OnImportFromExcelUnloaded);
			}
		}
		
		public function OnRequested(sender: Object, e: RequestEventArgs)
		{
			// we don't really use this at all at the moment
		}

		// set the error line against the spreadsheet
		private function setLineStatus(astrError : String)
		{
			if(null != gwbWorkbook)
			{
				gwbWorkbook.ActiveSheet.Range("S" + (giCurrentRow-1).ToString()).Value = astrError;
			}
		}
	
		public function OnRequestCompleted(sender: Object, e: RequestEventArgs)
		{
			var strError : String = null;
			try
			{
				if(e.CommandType == MNEProtocol.CommandTypeKey)     // we're looking for a key event
				{
					if(e.CommandValue == MNEProtocol.KeyEnter)      // specifically we're looking the enter key event
					{
						strError = checkForError();

						if(true == String.IsNullOrEmpty(strError))
						{
							if(true == giicInstanceController.RenderEngine.PanelHeader.EndsWith("E"))   // are we on panel E?
							{
								handleEPanel();     // handle panel E
								strError = checkForError();
							}
							else if(true == giicInstanceController.RenderEngine.PanelHeader.EndsWith("J1")) // are we on panel G1 (this should be GLS120/G1)
							{
								strError = checkForError();

								if(true != String.IsNullOrEmpty(strError))
								{
									giCurrentRow = giMaxRow + 1;
								}
								else
								{
									handleJ1Panel();    // handle panel j
								}
							}
						}
						else
						{
							setLineStatus(strError);
						}
					}
				}
				else if(e.CommandType == MNEProtocol.CommandTypeListOption)
				{
					if(e.CommandValue == MNEProtocol.OptionSelect)
					{
						if(true == String.IsNullOrEmpty(strError))
						{
							handleEPanel();
						}
					}
				}
				if(null != giicInstanceController.Response)
				{
					if(0 == String.Compare(giicInstanceController.Response.Request.RequestType.ToString(), "Panel"))
					{
						if((MNEProtocol.CommandTypeKey == giicInstanceController.Response.Request.CommandType) && (MNEProtocol.KeyF03 == giicInstanceController.Response.Request.CommandValue))
						{
							CleanUp();
						}
					}
				}
			}
			catch(ex)
			{
				MessageBox.Show(ex.message);
			}
			if(null != strError)
			{
				CleanUp();
			}
		}
		
		// this is where we do the actual handling of the J1 Panel
		private function handleJ1Panel()
		{
			if(giCurrentRow <= giMaxRow)    // the spreadsheet has a limited number of rows...
			{
				// extract the lines from the spreadsheet
				var strWWADIV : String = retrieveFromActiveSheet("H" + giCurrentRow);       // division
				var strWXAIT1 : String = retrieveFromActiveSheet("B" + giCurrentRow);
				var strWXAIT2 : String = retrieveFromActiveSheet("C" + giCurrentRow);
				var strWXAIT3 : String = retrieveFromActiveSheet("D" + giCurrentRow);
				var strWXAIT4 : String = retrieveFromActiveSheet("E" + giCurrentRow);
				var strWXAIT5 : String = retrieveFromActiveSheet("F" + giCurrentRow);
				var strWXAIT6 : String = retrieveFromActiveSheet("G" + giCurrentRow);
				var strWWCUAMDebit : String = retrieveFromActiveSheet("I" + giCurrentRow);
				var strWWCUAMCredit : String = retrieveFromActiveSheet("L" + giCurrentRow);
				var strWWVTXT : String = retrieveFromActiveSheet("O" + giCurrentRow);
				var strWWVTCD : String = retrieveFromActiveSheet("N" + giCurrentRow);

				// this is the current row
				giCurrentRow = giCurrentRow + 1;
				if(!String.IsNullOrEmpty(strWXAIT1))
				{
					if(0 != String.Compare(strWXAIT1,"undefined"))  // verify that we actually have content
					{
						if(0 != String.Compare(strWXAIT1,"End", true))
						{
							var bDoWeHaveAValue : boolean = false;
							
							if(!String.IsNullOrEmpty(strWWCUAMDebit))
							{
								var dblDebit : double = strWWCUAMDebit;
								var strDebit : String = dblDebit.ToString("#.##");  // make sure that we are formatted to only 2 decimal places
								if(!String.IsNullOrEmpty(strDebit))                 // ensure we actually have a value now that we have converted it
								{
									if(0 != String.Compare(strDebit, "0.00"))       // ensure that the value isn't 0!
									{
										bDoWeHaveAValue = true;
										setM3TextField("WWCUAM", strDebit);	// Value
									}
								}
							}
							if((!String.IsNullOrEmpty(strWWCUAMCredit)) && (false == bDoWeHaveAValue))  // 20120516 - if a forumla is used in the debit column and it is zero, then we incorrectly don't process the credit column
							{
								var dblCredit : double = strWWCUAMCredit;
								var strCredit : String = dblCredit.ToString("#.##");    // make sure that we are formatted to only 2 decimal places
								if(!String.IsNullOrEmpty(strCredit))                    // ensure we actually have a value now that we have converted it
								{
									if(0 != String.Compare(strCredit, "0.00"))          // ensure that the value isn't 0!
									{
										bDoWeHaveAValue = true;
										setM3TextField("WWCUAM", "-" + strCredit);	// Value
									}
								}
							}
							if(true == bDoWeHaveAValue)     // if the value is 0, then we shouldn't submit it
							{
								// strWWADIV
								if(!String.IsNullOrEmpty(strWWADIV))
								{
									setM3TextField("WWADIV", strWWADIV);	// division
								}

								setM3TextField("WXAIT1", strWXAIT1);	// account
	
								setM3TextField("WXAIT2", strWXAIT2);	// Dept
								setM3TextField("WXAIT3", strWXAIT3);	// Dim3
								setM3TextField("WXAIT4", strWXAIT4);	// Dim4
								setM3TextField("WXAIT5", strWXAIT5);	// Dim5
								setM3TextField("WXAIT6", strWXAIT6);	// Dim6
								//setM3TextField("", retrieveFromActiveSheet("H" + giCurrentRow));	// Division

						
								setM3TextField("WWVTXT", strWWVTXT);	// Voucher Text
								setM3TextField("WWVTCD", strWWVTCD);    // VAT type

								giicInstanceController.PressKey("ENTER");   // press the enter key
							}
							else
							{
								// we need to go to the next line to process
								handleJ1Panel();
							}
						}
						else
						{
							giCurrentRow = giMaxRow + 1;    // 20110207 - end this loop
							CleanUp();                      // and do our cleanup
						}

					}
                    else handleJ1Panel();   //MessageBox.Show("2. No content in column B");
				}
                else handleJ1Panel();   //MessageBox.Show("1. No content in column B");
			}
			else
			{
				//MessageBox.Show("handleJ1Panel(): " + giCurrentRow + " - " + giMaxRow);
				CleanUp();
			}
		}
		
		// a little wee helper function that will search for a TextBox name
		// and set the TextBox value
		private function setM3TextField(astrName : String, astrValue : String)
		{
			var tbTextBox : TextBox = ScriptUtil.FindChild(ggrdContentGrid, astrName);
			if(null != tbTextBox)
			{
				tbTextBox.Text = astrValue;
			}
			else MessageBox.Show("Can't find: " + tbTextBox.Text);
		}

		// retrieve some data from the active spreadsheet
		// at a specific location
		private function retrieveFromActiveSheet(astrPosition : String)
		{
			var strValue : String = gwbWorkbook.ActiveSheet.Range(astrPosition).Value;
			if(true == String.IsNullOrEmpty(strValue))
			{
				strValue = "";
			}
			else if(0 == String.Compare(strValue, "undefined"))
			{
				strValue = "";
			}

			return(strValue);
		}
		
		// retrieveVoucherType
		// gstrVoucherType
		// handle the E Panel
		private function handleEPanel()
		{
			var strAccountingDate : String = retrieveAccountingDate();  // retrieve the accounting date from the Spreadsheet
			var strVoucherText : String = retrieveVoucherText();        // retroeve the voucher text from the Spreadsheet

			var strReversalDate : String = retrieveReversingDate();     // retrieve the reversal date
			var bHaveAllTheFields : Boolean = false;                    // do we have all the fields that we require?



			if( (0 == String.Compare(gstrVoucherType,"100")) || (0 == String.Compare(gstrVoucherType,"200")) || (0 == String.Compare(gstrVoucherType,"900")))
			{
				if((!String.IsNullOrEmpty(strAccountingDate)) && (!String.IsNullOrEmpty(strVoucherText)))
				{
					if(0 == String.Compare(gstrVoucherType,"900"))
					{
						setM3YEA4(strAccountingDate);
					}
					else
					{
						setM3AccountingDate(strAccountingDate);     // now we actually set the accounting date in the TextBox
					}
					
					setM3VoucherText(strVoucherText);           // and the Voucher TextBox
					bHaveAllTheFields = true;
				}
			}
			else if(0 == String.Compare(gstrVoucherType,"300"))
			{
				if((!String.IsNullOrEmpty(strReversalDate)) && (!String.IsNullOrEmpty(strAccountingDate)) && (!String.IsNullOrEmpty(strVoucherText)))
				{
					setM3AccountingDate(strAccountingDate);     // now we actually set the accounting date in the TextBox
					setM3VoucherText(strVoucherText);           // and the Voucher TextBox
					setM3ReversalDate(strReversalDate);         // now set the reversal date

					bHaveAllTheFields = true;
				}
			}
			else
			{
				MessageBox.Show("Sorry, but we can't handle the voucher type: " + gstrVoucherType);
			}
			// if((!String.IsNullOrEmpty(strAccountingDate)) && (!String.IsNullOrEmpty(strVoucherText)))
			if(true == bHaveAllTheFields)
			{
				giicInstanceController.PressKey("ENTER");   // now we press enter - this will fire off a Request event and should take us to GLS120/G1
			}
			else MessageBox.Show("We require an Account Date and Voucher Text");
		}
		
		private function CleanUp()
		{
			if(true == gbRequest)
			{
				giicInstanceController.remove_RequestCompleted(OnRequestCompleted);
				giicInstanceController.remove_RequestCompleted(OnRequested);
			}
			gbRequest = false;
			CleanUpExcel();
			//MessageBox.Show("Cleaned up");
		}

		private function CleanUpExcel()
		{
				// check to ensure we have a Workbook object
				// before we attempt to close the workbook
				if(null != gwbWorkbook)
				{
                    // 20120527 start
                    if(null != gstrError)
                    {
                        if(System.Windows.MessageBoxResult.Yes == MessageBox.Show("Sadly there was an error processing, did you want to save the error on to the spreadsheet" + Environment.NewLine + gstrError, "Error", System.Windows.MessageBoxButton.YesNo))
                        {
                            gwbWorkbook.Save();
                        }
                    }
                    
					gwbWorkbook.Close(false);   // this prevents Excel from opening a prompt to save
                    // gwbWorkbook.Close();
                    // 20120527 end
					gwbWorkbook = null;
				}
				// make sure we have actually created
				// the Excel Application object before
				// we Quit
				if(null != gexaApplication)
				{
					gexaApplication.Quit();
					gexaApplication = null;
				}
		}

		private function InitialiseExcel()
		{
			var result = null;
			try
			{
				gexaApplication = new ActiveXObject("Excel.Application");
				gexaApplication.Visible = true;
				
                // Address a 'bug' "Old format or invalid type library" where
                // you run an english version of Excel but the regional settings of the
                // computer is configured for a non-English language (and the language pack
                // isn't installed)
                //  http://support.microsoft.com/kb/320369
                // 
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

				result = gexaApplication;
			}
			catch(exException)
			{
				MessageBox.Show("Error: " + exException.Message + Environment.NewLine + exException.StackTrace);
			}
			return(result);
		}
		
		//
	}
}
