using System;
using System.ComponentModel;
using System.Diagnostics;

using Microsoft.Dexterity.Bridge;
using Microsoft.Dexterity.Applications;
using Microsoft.Dexterity.Applications.DynamicsDictionary;

namespace GPUG.Develop
{
    /// <summary>
    /// Event handling on various controls in Dynamics GP also provides access to user control interactions
    /// </summary>
    //public class GP_Addin_WorkWithEvents_2 : IDexterityAddIn
    //{
    //    // local reference to the AboutBox form in Dynamics GP
    //    private SopEntryForm sopEntry = Dynamics.Forms.SopEntry;

    //    public void Initialize()
    //    {
    //        // working with events on the form
    //        this.sopEntry.OpenBeforeOriginal += delegate { Debug.Print("SOP Entry Form is opening"); };
    //        this.sopEntry.OpenAfterOriginal += delegate { Debug.Print("SOP Entry Form is open"); };
    //        this.sopEntry.CloseBeforeOriginal += delegate { Debug.Print("SOP Entry Form is closing"); };
    //        this.sopEntry.CloseAfterOriginal += delegate { Debug.Print("SOP Entry Form is closed"); };


    //        // working with events on a window in a form
    //        this.sopEntry.SopEntry.OpenBeforeOriginal += delegate { Debug.Print("SOP Entry Window is opening"); };
    //        this.sopEntry.SopEntry.OpenAfterOriginal += delegate { Debug.Print("SOP Entry Window is open"); };
    //        this.sopEntry.SopEntry.ActivateBeforeOriginal += delegate { Debug.Print("SOP Entry Window is receiving focus"); };
    //        this.sopEntry.SopEntry.ActivateAfterOriginal += delegate { Debug.Print("SOP Entry Window has focus"); };
    //        this.sopEntry.SopEntry.CloseBeforeOriginal += delegate { Debug.Print("SOP Entry Window is closing"); };
    //        this.sopEntry.SopEntry.CloseAfterOriginal += delegate { Debug.Print("SOP Entry Window is closed"); };
    //        this.sopEntry.SopEntry.PrintBeforeOriginal += delegate { Debug.Print("SOP Entry Window is opening a printer dialog"); };
    //        this.sopEntry.SopEntry.PrintAfterOriginal += delegate { Debug.Print("SOP Entry Window has opened a printer dialog"); };
    //        this.sopEntry.SopEntry.BeforeModalDialog += delegate { Debug.Print("SOP Entry Window is opening a modal dialog"); };
    //        this.sopEntry.SopEntry.AfterModalDialog += delegate { Debug.Print("SOP Entry Window has opened a modal dialog"); };


    //        // working with events on buttons
    //        this.sopEntry.SopEntry.SaveButton.ClickBeforeOriginal += new CancelEventHandler(Before_SaveButtonClick);
    //        this.sopEntry.SopEntry.SaveButton.ClickAfterOriginal += new EventHandler(After_SaveButtonClick);


    //        // working with events in text fields
    //        this.sopEntry.SopEntry.BatchNumber.Change += delegate { Debug.Print("The batch number field change"); };
    //        this.sopEntry.SopEntry.BatchNumber.EnterBeforeOriginal += delegate { Debug.Print("The batch number field before entry"); };
    //        this.sopEntry.SopEntry.BatchNumber.EnterAfterOriginal += delegate { Debug.Print("The batch number field after entry"); };
    //        this.sopEntry.SopEntry.BatchNumber.LeaveBeforeOriginal += delegate { Debug.Print("The batch number field before leaving"); };
    //        this.sopEntry.SopEntry.BatchNumber.LeaveAfterOriginal += delegate { Debug.Print("The batch number field after leaving"); };
    //        this.sopEntry.SopEntry.BatchNumber.ValidateBeforeOriginal += delegate { Debug.Print("The batch number field before validation"); };
    //        this.sopEntry.SopEntry.BatchNumber.ValidateAfterOriginal += delegate { Debug.Print("The batch number field after validation"); };


    //        // working with events in dropdown boxes
    //        // note:    drop down boxes cause the parent window to fire the open before and after events. 
    //        //          This will also cause the change events to fire from the text box controls
    //        this.sopEntry.SopEntry.SopType.Change += delegate { Debug.Print("The batch number field change"); };
    //        this.sopEntry.SopEntry.SopType.EnterBeforeOriginal += delegate { Debug.Print("The SOP type dropdown before entry"); };
    //        this.sopEntry.SopEntry.SopType.EnterAfterOriginal += delegate { Debug.Print("The SOP type dropdown after entry"); };
    //        this.sopEntry.SopEntry.SopType.LeaveBeforeOriginal += delegate { Debug.Print("The SOP type dropdown before leaving"); };
    //        this.sopEntry.SopEntry.SopType.LeaveAfterOriginal += delegate { Debug.Print("The SOP type dropdown after leaving"); };
    //        this.sopEntry.SopEntry.SopType.ValidateBeforeOriginal += delegate { Debug.Print("The SOP type dropdown before validation"); };
    //        this.sopEntry.SopEntry.SopType.ValidateAfterOriginal += delegate { Debug.Print("The SOP type dropdown after validation"); };


    //        // working with scroll windows
    //        this.sopEntry.SopEntry.LineScroll.LineEnterBeforeOriginal += delegate { Debug.Print("Line before entry"); };
    //        this.sopEntry.SopEntry.LineScroll.LineEnterAfterOriginal += delegate { Debug.Print("Line after entry"); };
    //        this.sopEntry.SopEntry.LineScroll.LineChangeBeforeOriginal += delegate { Debug.Print("Line before change"); };
    //        this.sopEntry.SopEntry.LineScroll.LineChangeAfterOriginal += delegate { Debug.Print("Line after change"); };
    //        this.sopEntry.SopEntry.LineScroll.LineLeaveBeforeOriginal += delegate { Debug.Print("Line before leaving"); };
    //        this.sopEntry.SopEntry.LineScroll.LineLeaveAfterOriginal += delegate { Debug.Print("Line after leaving"); };
    //        this.sopEntry.SopEntry.LineScroll.LineInsertBeforeOriginal += delegate { Debug.Print("Line before insert"); };
    //        this.sopEntry.SopEntry.LineScroll.LineInsertAfterOriginal += delegate { Debug.Print("Line after insert"); };
    //        this.sopEntry.SopEntry.LineScroll.LineFillBeforeOriginal += delegate { Debug.Print("Line before fill"); };
    //        this.sopEntry.SopEntry.LineScroll.LineFillAfterOriginal += delegate { Debug.Print("Line after fill"); };
    //        this.sopEntry.SopEntry.LineScroll.LineDeleteBeforeOriginal += delegate { Debug.Print("Line before delete"); };
    //        this.sopEntry.SopEntry.LineScroll.LineDeleteAfterOriginal += delegate { Debug.Print("Line after delete"); };

    //    }


    //    /// <summary>
    //    /// Working with button clicks on a Dyanmics GP window
    //    /// </summary>
    //    private void Before_SaveButtonClick(object sender, CancelEventArgs e)
    //    {
    //        Debug.Print("The save button is being clicked, and the user data in the window is still available.");

    //        string customerNumber = this.sopEntry.SopEntry.CustomerNumber.Value;
    //        Debug.Print("The customer number is " + customerNumber);
    //    }

    //    /// <summary>
    //    /// Working with button clicks on a Dyanmics GP window
    //    /// </summary>
    //    private void After_SaveButtonClick(object sender, EventArgs e)
    //    {
    //        Debug.Print("The save button is finished being clicked, and the user data in the window is no longer available.");

    //        string customerNumber = this.sopEntry.SopEntry.CustomerNumber.Value;
    //        Debug.Print("The customer number is " + customerNumber);
    //    }

        
    //}
}


