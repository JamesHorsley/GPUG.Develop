using System;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Dexterity.Bridge;
using Microsoft.Dexterity.Applications;
using Microsoft.Dexterity.Applications.DynamicsDictionary;
using System.Diagnostics;

namespace GPUG.Develop
{
    /// <summary>
    /// Creating a simple plugin framework to access some events on the About Dynamics GP form
    /// </summary>
    public class Lesson_1b : IDexterityAddIn
    {
        // create a private reference to the form to make the instance available to the addin, and to initialize any event handling
        private AboutBoxForm aboutForm = Dynamics.Forms.AboutBox;

        public void Initialize()
        {
            // adding to the "Additional" menu for custom functions and introducing custom windows forms
            // the "0" is a shortcut declaration that allows hot-key functionality in the menu
            this.aboutForm.AddMenuHandler(customMenuFunction, "GPUG AddIn Menu Option", "0");
        }

        // add menu handling to a window
        private void customMenuFunction(object sender, EventArgs e)
        {
            // make a note of the debug console here when the message box is opened
            // this will fire off the activate window before and after events again
            MessageBox.Show("I'm a custom menu function with lots of potential");
        }
    }
}
