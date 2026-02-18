import wx
import list_documents
from pathlib import Path
from copy_renamed_files import copy_renamed_files

class GenerateIndexPanel(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)

        self.description = wx.StaticText(
            self,
            label=(
                "Generate an index file for a selected folder.\n"
                "Choose the folder, select where to save the output file, "
                "and enable any optional settings below."
            ),
        )
        self.description.Wrap(400)

        self.checkbox = wx.CheckBox(self, label="Include additional details for email")
        self.checkbox.SetValue(True)

        self.select_folder_btn = wx.Button(self, label="Select Folder")
        self.save_as_btn = wx.Button(self, label="Generate index")

        self.folder_label = wx.StaticText(self, label="No folder selected")
        self.save_label = wx.StaticText(self, label="No file selected")

        # Description sizer
        desc_sizer = wx.BoxSizer(wx.HORIZONTAL)
        desc_sizer.Add(self.description, 1, wx.EXPAND)

        # Main layout
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(desc_sizer, 0, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(self.checkbox, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        main_sizer.Add(self.select_folder_btn, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        main_sizer.Add(self.folder_label, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        main_sizer.Add(self.save_as_btn, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        main_sizer.Add(self.save_label, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        self.SetSizer(main_sizer)
        # Events
        self.select_folder_btn.Bind(wx.EVT_BUTTON, self.on_select_folder)
        self.save_as_btn.Bind(wx.EVT_BUTTON, self.on_save_as)

        self.Layout()

    def on_select_folder(self, event):
        with wx.DirDialog(self, "Select a folder") as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.folder_label.SetLabel("Folder to index is: " + dlg.GetPath())
                
                parent_frame = self.GetTopLevelParent()
                parent_frame.folder_path = dlg.GetPath()
                
                

    def on_save_as(self, event):
        
        parent_frame = self.GetTopLevelParent()
        
        resolved_folder_path = Path(parent_frame.folder_path).resolve()

        if parent_frame.folder_path == "" or not resolved_folder_path.is_dir():
            with wx.MessageDialog(self, "Please select a valid folder before generating the index.", "Error", wx.OK | wx.ICON_ERROR) as error_dlg:
                error_dlg.ShowModal()

        else:
        
            with wx.FileDialog(
                self,
                "Save As",
                wildcard="Excel file (*.xlsx)|*.xlsx",
                defaultFile="index.xlsx",
                style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
            ) as dlg:
                current_path = Path(parent_frame.folder_path)
                
                new_path = str(current_path.parent)
                
                dlg.SetDirectory(new_path)

                if dlg.ShowModal() == wx.ID_OK:
                    self.save_label.SetLabel("Wait...")
                    list_documents.export_folder_contents_to_excel(parent_frame.folder_path, dlg.GetPath(), self.checkbox.IsChecked(), True)
                    self.save_label.SetLabel("Index created successfully at " + dlg.GetPath())
                    parent_frame.index_path = dlg.GetPath()
                    parent_frame.index_set = True


class RenameFilesPanel(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)
        self.description = wx.StaticText(
            self,
            label=(
                "Rename files based on a selected index file.\n"
                "Currently the program just creates a duplicate folder which has been renamed according to the index file. In the future, the program will rename the files in place.\n"
            ),
        )
        self.description.Wrap(400)
    
        self.select_file_btn = wx.Button(self, label="Select index")
        self.select_folder_btn = wx.Button(self, label="Create new folder with renamed files")

        self.file_label = wx.StaticText(self, label="No file selected")
        self.folder_label = wx.StaticText(self, label="No folder selected")

        desc_sizer = wx.BoxSizer(wx.HORIZONTAL)
        desc_sizer.Add(self.description, 1, wx.EXPAND)

        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(desc_sizer, 0, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(self.select_file_btn, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        main_sizer.Add(self.file_label, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        main_sizer.Add(self.select_folder_btn, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        main_sizer.Add(self.folder_label, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)
        parent_frame = self.GetTopLevelParent()

        if parent_frame.index_set:

            self.file_label.SetLabel("Index selected: " + parent_frame.index_path)
            
        else:

            self.file_label.SetLabel("No index currently selected")

        self.SetSizer(main_sizer)

        self.select_file_btn.Bind(wx.EVT_BUTTON, self.on_select_file)
        self.select_folder_btn.Bind(wx.EVT_BUTTON, self.on_select_folder)
        self.Layout()

    def on_select_file(self, event):
                
        with wx.FileDialog(
            self,
            "Select index",
            wildcard="Excel file (*.xlsx)|*.xlsx",
            defaultFile="index.xlsx",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
        ) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.file_label.SetLabel("Index selected:" + dlg.GetPath())
                parent_frame = self.GetTopLevelParent()
                parent_frame.index_path = dlg.GetPath()

    def on_select_folder(self, event):
        parent_frame = self.GetTopLevelParent()

        resolved_index_path = Path(parent_frame.index_path).resolve()


        if parent_frame.index_path == "" or not resolved_index_path.is_file():
            with wx.MessageDialog(self, "Please select a valid index before renaming.", "Error", wx.OK | wx.ICON_ERROR) as error_dlg:
                error_dlg.ShowModal()

        else:
             
            if parent_frame.index_set:

                resolved_folder_path = Path(parent_frame.folder_path).resolve()
                new_folder_name = resolved_folder_path.name + "_renamed"
                new_folder_path = resolved_folder_path.parent / new_folder_name
        
                with wx.MessageDialog(self, "Folder created successfully at: " + str(new_folder_path), "Success", wx.OK) as success_dlg:
                    copy_renamed_files(resolved_index_path, resolved_folder_path, new_folder_path)
                    success_dlg.ShowModal()

            else:

                with wx.DirDialog(self, "Select the root folder that correponds to the index") as dlg:
                    if dlg.ShowModal() == wx.ID_OK:

                        resolved_folder_path = Path(dlg.GetPath()).resolve()
                        new_folder_name = resolved_folder_path.name + "_renamed"
                        new_folder_path = resolved_folder_path.parent / new_folder_name
                        self.folder_label.SetLabel(str(new_folder_path))
                        parent_frame.index_set = True
                        copy_renamed_files(resolved_index_path, resolved_folder_path, new_folder_path)

            #with wx.DirDialog(self, "Select an output folder") as dlg:
            #    if dlg.ShowModal() == wx.ID_OK:
            #        self.folder_label.SetLabel("Folder successfully copied to " +dlg.GetPath())


class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Index generator", size=(450, 350))

        notebook = wx.Notebook(self)

        

        self.folder_path = ""
        self.index_path = ""
        self.index_set = False

        notebook.AddPage(GenerateIndexPanel(notebook), "Generate index")
        notebook.AddPage(RenameFilesPanel(notebook), "Rename files")

        close_btn = wx.Button(self, wx.ID_CLOSE, "Close")
        close_btn.Bind(wx.EVT_BUTTON, self.on_close)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        button_sizer.AddStretchSpacer()
        button_sizer.Add(close_btn, 0, wx.ALL, 10)

        frame_sizer = wx.BoxSizer(wx.VERTICAL)
        frame_sizer.Add(notebook, 1, wx.EXPAND)
        frame_sizer.Add(button_sizer, 0, wx.EXPAND)

        self.SetSizer(frame_sizer)


    def on_close(self, event):
        self.Close()


class MyApp(wx.App):
    def OnInit(self):
        frame = MainFrame()
        frame.Show()
        return True


if __name__ == "__main__":
    app = MyApp()
    app.MainLoop()