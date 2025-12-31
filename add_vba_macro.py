"""
Script to add VBA macro to OPC_TEST.xlsm
Requires: pip install pywin32

This script uses the Windows COM interface to add VBA code to an Excel workbook.
"""

import os
import sys

def add_macro_to_workbook():
    try:
        import win32com.client as win32
    except ImportError:
        print("Error: pywin32 is required. Install with: pip install pywin32")
        sys.exit(1)
    
    # Paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    workbook_path = os.path.join(script_dir, "OPC_TEST.xlsm")
    macro_path = os.path.join(script_dir, "ProcessEmailAttachments.bas")
    
    # Check files exist
    if not os.path.exists(workbook_path):
        print(f"Error: Workbook not found at {workbook_path}")
        sys.exit(1)
    
    if not os.path.exists(macro_path):
        print(f"Error: Macro file not found at {macro_path}")
        sys.exit(1)
    
    # Read the VBA code from the .bas file
    with open(macro_path, 'r', encoding='utf-8') as f:
        vba_code = f.read()
    
    # Remove the Attribute lines that are added when exporting (if present)
    # These are auto-generated when importing
    lines = vba_code.split('\n')
    cleaned_lines = [line for line in lines if not line.startswith('Attribute ')]
    vba_code = '\n'.join(cleaned_lines)
    
    print("Opening Excel...")
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        print(f"Opening workbook: {workbook_path}")
        workbook = excel.Workbooks.Open(workbook_path)
        
        # Access the VBA project
        vba_project = workbook.VBProject
        
        # Check if module already exists and remove it
        module_name = "ProcessEmailAttachments"
        for component in vba_project.VBComponents:
            if component.Name == module_name:
                print(f"Removing existing module: {module_name}")
                vba_project.VBComponents.Remove(component)
                break
        
        # Add new module
        print(f"Adding module: {module_name}")
        new_module = vba_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        new_module.Name = module_name
        
        # Add the code to the module
        new_module.CodeModule.AddFromString(vba_code)
        
        # Save the workbook
        print("Saving workbook...")
        workbook.Save()
        
        print("=" * 50)
        print("SUCCESS! Macro added to OPC_TEST.xlsm")
        print("=" * 50)
        print("\nTo run the macro:")
        print("1. Open OPC_TEST.xlsm in Excel")
        print("2. Select a cell in the row you want to process")
        print("3. Press Alt+F8, select 'ProcessEmailAttachments', click Run")
        print("\nNote: You may need to enable macros and add references:")
        print("  - Microsoft Outlook Object Library")
        print("  - (Optional) Adobe Acrobat Type Library")
        
    except Exception as e:
        print(f"Error: {e}")
        print("\nIf you get a 'Programmatic access to Visual Basic Project' error:")
        print("1. Open Excel")
        print("2. Go to File > Options > Trust Center > Trust Center Settings")
        print("3. Click 'Macro Settings'")
        print("4. Check 'Trust access to the VBA project object model'")
        print("5. Click OK and try again")
        raise
    finally:
        workbook.Close(SaveChanges=True)
        excel.Quit()

if __name__ == "__main__":
    add_macro_to_workbook()
