using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;

using System.Collections.Generic;
using System.Linq;
using System.Text; 

namespace ConsoleApp1
{

    class Program
    { 

        static void Main(string[] args)
        {
            string presentationFilePath = @"C:\Users\Ralph.hachache\Downloads\Input.pptx";
            Process g = new Process();
            g.ChangePackage(presentationFilePath);  
        } 
         
    }
}

