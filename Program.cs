using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using AutoItX3Lib;

namespace Autoit
{

    public class NotepadAuto
    {

        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private static extern int FindWindow(string sClass, string sWindow);

        static void Main(string[] args)
        {
           
            AutoItX3 auto = new AutoItX3();

            //Asks for user input of desired path for saving the file
            //if the directory dosn"t exist then it continues to ask for 
            //path till a existing directory is given.
            Boolean check =true;
            string path1 = "";
            while (check)
            {
                Console.WriteLine("Enter a desired path for save please(Ex: C:\\Users\\person\\Downloads\\): ");
                path1 = Console.ReadLine();
                if(Directory.Exists(path1))
                    check = false;
                else
                    Console.WriteLine("Please enter a valid path.(Ex: C:\\Users\\person\\Downloads\\)");
            }

            // var to check if notepad app is open. 
            int hwnd = 0;
            int hwnd2 = 0;

            //boolean check for if a notepade proccess is not currently exist. 
            Boolean b = true;

            //grabs handles for notpad if it is already opened.
            hwnd = FindWindow(null, "*Untitled - Notepad");
            hwnd2 = FindWindow(null, "Untitled - Notepad");

            while (b)
            {
                //Check if note is open
                if (hwnd > 0 || hwnd2 > 0)
                {
                    Console.WriteLine("Notepad application handle has been aquired.");

                    //Clicks on File menu and new
                    auto.WinMenuSelectItem("*Untitled - Notepad", "", "&File", "&New");
                    Console.WriteLine("file and new buttons have been clicked.");
                   
                    //types hello world in text field
                    auto.Send("Hello World");
                    Console.WriteLine("Hello world message sent to the ,txt file in notepad.");

                    //Clicks on File menu and save as
                    auto.WinMenuSelectItem("*Untitled - Notepad", "", "&File", "&Save");
                    Console.WriteLine("Save button has been clicked.");

                    // waiting for Save As menu
                    auto.WinWaitActive("Save As");

                    //Code for Clicking through Save As
                    auto.ControlFocus("Save As", "", "Edit1");

                    //Saves the file with a specific path.
                    auto.ControlSend("Save As", "", "Edit1", path1+ "Untitled");
                    auto.ControlClick("Save As", "&Save", "Button2");
                    Console.WriteLine("File has been saved to "+ path1);

                    //waits for pop of dupelicate files.
                    auto.WinWaitActive("Confirm Save As");

                    //Code for replacing a copy of the file.
                    auto.ControlClick("Confirm Save As", "&Yes", "Button1");
                    Console.WriteLine("Copy file has replaced *Untitled.txt.");

                    //stops looping.
                    b = false;

                    //Check if file was saved in its given path.
                    Console.WriteLine(File.Exists(path1 + "Untitled.txt") ? "File saved in desired path." : "File is not saved or not saved in desired path.");

                    //close notepad.
                    Process[] proc = Process.GetProcessesByName("Notepad");
                    proc[0].Kill();
                }
                //if notepad isnt open, then it starts. 
                else
                {
                    Console.WriteLine("Notepad has not been launched. Starting Notepad.");
                    
                    //starts notepad
                    auto.Run("notepad.exe");
                    
                    //pause execution of script until window opens
                    auto.WinWaitActive("Untitled - Notepad");
                    
                    //grabs notepad handler # so we can proceede with the program and end the loop.
                    hwnd = FindWindow(null, "Untitled - Notepad");
                }
            }
        }
    }
}
