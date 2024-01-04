
using ConsoleApp1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Media3D;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
//[assembly: assemblyversion("1.0.0.3")]
//[assembly: assemblyfileversion("1.0.0.3")]
//[assembly: assemblyinformationalversion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
//[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context, System.Windows.Window window) //ScriptEnvironment environment)
        {
            //UserControl1 userControl = new UserControl1(context);
            UserControl1 userControl = new UserControl1(context, window);
            //window.Hide();
            //window.Height = 450;
            //window.Width = 500;
            //window.Content = userControl;


        }
    }
}