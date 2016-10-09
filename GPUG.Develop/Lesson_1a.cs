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
    public class Lesson_1a : IDexterityAddIn
    {
        // create a private reference to the form to make the instance available to the addin, and to initialize any event handling
        private AboutBoxForm aboutForm = Dynamics.Forms.AboutBox;

        public void Initialize()
        {
            // Initialize add-in start up code here.
            this.aboutForm.AboutBox.OpenAfterOriginal += new EventHandler(openAbout);
            this.aboutForm.AboutBox.CloseAfterOriginal += new EventHandler(closeAbout);

            // activation of the window
            this.aboutForm.AboutBox.ActivateBeforeOriginal += new CancelEventHandler(activateBeforeAbout);
            this.aboutForm.AboutBox.ActivateAfterOriginal += new EventHandler(activateAfterAbout);
        }
        
        // form open event
        private void openAbout(object sender, EventArgs e)
        {
            Debug.Print("Hi! I am going to open the About GP form");
        }

        // form close event
        private void closeAbout(object sender, EventArgs e)
        {
            Debug.Print("I am so sad you are closing About GP form :(");
        }

        // form > window before activate
        private void activateBeforeAbout(object sender, CancelEventArgs e)
        {
            Debug.Print("The About GP window is activating. Think of me as before getting focus.");
        }

        // form > window after activate
        private void activateAfterAbout(object sender, EventArgs e)
        {
            Debug.Print("The About GP window has activated something fun.");
        }
    }
}
