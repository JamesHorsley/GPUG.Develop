using System.Diagnostics;
using Microsoft.Dexterity.Bridge;

namespace GPUG.Develop
{
    /// <summary>
    /// Creating our first plugin for GP
    /// </summary>
    public class Lesson_1 : IDexterityAddIn
    {
        public void Initialize()
        {
            Debug.Print("Hello GPUG! I am loaded when all the addins are processed.");
        }
    }
}
