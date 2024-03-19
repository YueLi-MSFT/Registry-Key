This is a Windows Batch script that is used to add or delete registry keys in the Windows Registry. Here's how to use it:

1. Open a command prompt window.
2. Navigate to the directory where the `sideload.bat` file is located.
3. Run the script with one of the following commands:

   - To add a registry key:
     ```
     sideload.bat add [value1] [value2]
     ```
     Replace `[value1]` and `[value2]` with the values you want to set for the registry keys `Microsoft.Office.OEP.MosProviderEnabled` and `Microsoft.Office.OEP.MosManifest` respectively.

   - To delete the registry key:
     ```
     sideload.bat delete
     ```
4. The script will then add or delete the registry keys as per your command and print a success or failure message. It will then pause, keeping the command prompt window open until you press any key.

Please note that modifying the Windows Registry can have significant impacts on your system's operation and should only be done if you are sure of what you are doing. Always back up your registry before making any changes.
