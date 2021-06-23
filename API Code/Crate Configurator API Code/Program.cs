using System;
using System.Runtime.InteropServices;
using Inventor;
using System.Windows.Forms;

namespace Crate_Configurator_API_Code
{
    class Program
    {
        static void Main(string[] args)
        {

            Inventor.Application _invApp;

            try
            {
                _invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
                _invApp.Visible = true;
                MessageBox.Show("Connected to existing instance of Autodesk Inventor");
            }
            catch
            {
                MessageBox.Show("Couldn't connect to Autodesk Inventor, please open Autodesk Inventor, opening a new instance");

                Type invType = Type.GetTypeFromProgID("Inventor.Application");
                _invApp = Activator.CreateInstance(invType) as Inventor.Application;
                _invApp.Visible = true;
            }

            AssemblyDocument assemblyDocument = (Inventor.AssemblyDocument)_invApp.ActiveDocument;

            if(assemblyDocument != null)
            {

                AssemblyComponentDefinition assemblyCompDef = assemblyDocument.ComponentDefinition;

                UserParameters assemblyUserParameters = assemblyCompDef.Parameters.UserParameters;

                assemblyUserParameters[4].Value = 5;

                assemblyDocument.Update();
            }
            else
            {
                MessageBox.Show("Open the crate assembly file");
            }

        }
    }
}
