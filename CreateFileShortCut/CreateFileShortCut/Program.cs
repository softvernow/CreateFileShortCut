using IWshRuntimeLibrary;

namespace CreateFileShortCut
{
    class Program
    {
        static void Main(string[] args)
        {
            //Initiate instance of WshShell using IWshRuntimeLibrary.
            WshShell wsh = new WshShell();

            //on line below you will enter the path where you want to create the shortcut, for example, if you want to create the shortcut on your desktop.
            IWshRuntimeLibrary.IWshShortcut shortcut = wsh.CreateShortcut(@"C:\Users\USER_NAME\Desktop\your_shortcut_name.lnk") as IWshRuntimeLibrary.IWshShortcut;

            //if you need your shortcut to start with arguments use line below to enter arguments. 
            //shortcut.Arguments = "";

            //on line below you will enter the target of original file. 
            shortcut.TargetPath = @"C:\Program Files\YOUR_PROGRAM_FOLDER\your_program_name.exe";
            
            //on line below you will enter the description of your shortcut. 
            shortcut.Description = "Description of Shortcut";

            //on line below you will enter the Working Directory of shortcut - Directory of your program. 
            shortcut.WorkingDirectory = @"C:\Program Files\YOUR_PROGRAM_FOLDER\";

            //if you want to another ICON for shortcut, on line below you will enter the path of the icon you want to use. 
            shortcut.IconLocation = @"C:\Program Files\YOUR_PROGRAM_FOLDER\your_icon.ico";
            shortcut.Save();


        }
    }
}
