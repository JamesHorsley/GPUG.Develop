using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Dexterity.Bridge;
using Microsoft.Dexterity.Applications;
using Microsoft.Dexterity.Applications.DynamicsDictionary;

namespace GPUG.Develop
{
    ///// <summary>
    ///// Event handling is the primary way to move interactions between Dynamics GP and Visual Studio Tools 
    ///// </summary>
    //public class GP_Addin_WorkWithEvents : IDexterityAddIn
    //{
    //    // local reference to the AboutBox form in Dynamics GP
    //    private AboutBoxForm aboutGP = Dynamics.Forms.AboutBox;

    //    public void Initialize()
    //    {
    //        // Initialize add-in start up code here.
    //        this.aboutGP.AboutBox.OpenBeforeOriginal += delegate { Debug.Print("OpenBeforeOriginal - Before the window opens"); };
    //        this.aboutGP.AboutBox.OpenAfterOriginal += delegate { Debug.Print("OpenAfterOriginal - After the window opens."); };
    //        this.aboutGP.AboutBox.ActivateBeforeOriginal += delegate { Debug.Print("ActivateBeforeOriginal - Before window receives focus"); };
    //        this.aboutGP.AboutBox.ActivateAfterOriginal += delegate { Debug.Print("ActivateAfterOriginal - After window receives focus"); };
    //        this.aboutGP.AboutBox.CloseBeforeOriginal += delegate { Debug.Print("CloseBeforeOriginal - Before the window closes."); };
    //        this.aboutGP.AboutBox.CloseAfterOriginal += delegate { Debug.Print("CloseAfterOriginal - After the window closes."); };
    //        this.aboutGP.AboutBox.BeforeModalDialog += delegate { Debug.Print("BeforeModalDialog - Before a popup window shows"); };
    //        this.aboutGP.AboutBox.AfterModalDialog += delegate { Debug.Print("AfterModalDialog - After a popup window shows"); };
    //        this.aboutGP.AboutBox.PrintBeforeOriginal += delegate { Debug.Print("PrintBeforeOriginal - Before the printer dialog opens"); };
    //        this.aboutGP.AboutBox.PrintAfterOriginal += delegate { Debug.Print("PrintAfterOriginal - After the printer dialog opens"); };

    //    }
    //}
}
