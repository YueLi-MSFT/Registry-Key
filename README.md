# Office Registry Modifier

This is a Windows Batch script that is used to add or delete specific registry keys in the Windows Registry related to Microsoft Office.

## Usage

1. Open a command prompt window.
2. Navigate to the directory where the `office_registry_modifier.bat` file is located.
3. Run the script with one of the following commands:

   - To add the registry keys:
     ```
     office_registry_modifier.bat add
     ```
   - To delete the registry keys:
     ```
     office_registry_modifier.bat delete
     ```
4. The script will add or delete the following keys:
   - `Microsoft.Office.OEP.MosProviderEnabled` with value `true`
   - `Microsoft.Office.OEP.MosManifest` with value `true`
   - `OEP.CG_MosPopulateContentEnabled` with value `true`
   - `OEP.EnableOsfMosAppFlyout` with value `true`
   - `OEP.MosPopulateContentEnabled` with value `true`
   - `OEP.ChangeGate.DedupeXmlAddInForMosExtension` with value `false`

5. The script will then print a success or failure message. It will then pause, keeping the command prompt window open until you press any key.

## Warning

Modifying the Windows Registry can have significant impacts on your system's operation and should only be done if you are sure of what you are doing. Always back up your registry before making any changes.