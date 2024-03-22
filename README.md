# Office Registry Modifier

This is a Windows Batch script that is used to add or delete specific registry keys in the Windows Registry related to Microsoft Office.

## Usage

1. Open a command prompt window.
2. Navigate to the directory where the `office_registry_modifier.bat` file is located.
3. Run the script with one of the following commands:

   - To add a registry key:
     ```
     office_registry_modifier.bat add [value1] [value2]
     ```
     Replace `[value1]` and `[value2]` with the values you want to set for the registry keys `Microsoft.Office.OEP.MosProviderEnabled` and `Microsoft.Office.OEP.MosManifest` respectively. The values can be `true` or `false`.

   - To delete the registry key:
     ```
     office_registry_modifier.bat delete
     ```
4. The script will then add or delete the registry keys as per your command and print a success or failure message. It will then pause, keeping the command prompt window open until you press any key.

## Warning

Modifying the Windows Registry can have significant impacts on your system's operation and should only be done if you are sure of what you are doing. Always back up your registry before making any changes.