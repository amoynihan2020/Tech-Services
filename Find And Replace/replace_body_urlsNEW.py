import pandas as pd
import time
import wx
import os


class App:

    def __init__(self):
        self.find_list = []
        self.sanitize_list = []

    @staticmethod
    def get_df(file):
        if file.endswith('.csv', -4):
            df = pd.read_csv(file)
        elif file.endswith('.xlsx', -5) or file.endswith('.xls', -4):
            df = pd.read_excel(file)
        elif file.endswith('.json', -5):
            
            df = pd.read_json(file)
        else:
            raise Exception('Unsupported file type')
        return df

    @staticmethod
    def find_replace(body_df, replace_df):
     
        body = body_df['body']
        title= body_df['title']
        content_type = body_df['content_type']
        content_id = body_df['id']
       
        category= body_df['category']
        published = body_df['published']
        keywords = body_df['keywords']
        
        
        
        
        source = body_df['source']
      
        frm.SetStatusText(f"Finding and Replacing")
        
        for index, row in replace_df.iterrows():
            
            body = body.str.replace(row['find'], row['replace'], case=False, regex=False)
            title = title.str.replace(row['find'], row['replace'], case=False, regex = False)
            
            
            
        body_df['body'] = body
        body_df['title'] = title
        
       
        s1 = body_df['body']
        s2 = body_df['title']
        
        final_df = pd.concat([s2,s1, content_type,content_id,category, published,keywords,source],axis=1)
        final_df.reset_index()
       
        return final_df

    # def sanitize_quotes(self, df):
    #     body = df['body']
    #     frm.SetStatusText(f"Sanitizing Quotes")
    #     body = body.str.replace(r'(\"{2,4}(?!(href=)))(?!\s)|(?<!(href=))\"{2,4}(?=\s)', '"', case=False)
    #     df.loc['body'] = body
    #     return df


#file1 isJSON
#file2 is the All Content report 
    def main(self, file1, file2):
        
   
        try:
            #JSON
            find_df = self.get_df(file1)
            find_df = find_df.filter(['content_type', 'id', 'body','title', 'category', 'published','keywords', 'source'])
            find_df = find_df.dropna(subset=['body'])
           
            findCopy= find_df.copy()
            for index, rows in findCopy.iterrows():
                categoryNoBrack = str(rows['category']).strip('[]')
                keyWordsNoBrack = str(rows['keywords']).strip('[]')
                
                
                keyWordsNoBrack =  keyWordsNoBrack.replace("'", "")
                categoryNoBrack = categoryNoBrack.replace("'", "")
                
                findCopy['category'].set_value(index, categoryNoBrack)
                findCopy['keywords'].set_value(index, keyWordsNoBrack)
                 
#            for item in tryDf.iteritems():
#               # print(str(item[1]).strip('[]'))
#                noBrack.append(str(item[1]).strip('[]'))
#            # find_df = self.sanitize_quotes(find_df)
#            x = pd.Series(noBrack)
#            print(x)
            
            find_df = findCopy
            replace_df = self.get_df(file2)
            export_df = self.find_replace(find_df, replace_df)
            export_df.to_csv('find_and_replace_import_' + str(time.time()).replace('.', '') + '.csv', index=False)
            export_df.to_json('find_and_replace_json' + str(time.time()).replace('.','') + '.json', orient= 'records', lines = True)
            return 'Find and replace complete'
        except Exception as e:
            
            return str(e)


class GUI(wx.Frame):

    def __init__(self, *args, **kwargs):
        kwargs["style"] = (wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER)
        self.app = App()
        super(GUI, self).__init__(*args, **kwargs)
        self.pnl = wx.Panel(self, size=(50, 50))
        self.pnl.SetBackgroundColour("#e6fffe")
        self.path = os.getcwd()
        self.vbox = wx.BoxSizer(wx.HORIZONTAL)
        self.InitUI()
        self.mainPage()
        self.vbox.SetSizeHints(self)
        self.SetSizer(self.vbox)

    def mainPage(self):
        self.btn = wx.Button(self.pnl, -1, "Select United Report")
        self.vbox.Add(self.btn, 0, wx.ALIGN_CENTER)
        self.btn.Bind(wx.EVT_BUTTON, self.OnClicked)
        self.btn2 = wx.Button(self.pnl, -1, "Select Replacement Map File")
        self.vbox.Add(self.btn2, 0, wx.ALIGN_CENTER)
        self.btn2.Bind(wx.EVT_BUTTON, self.OnClicked2)
        self.btn3 = wx.Button(self.pnl, -1, "Select Output Directory")
        self.vbox.Add(self.btn3, 0, wx.ALIGN_CENTER)
        self.btn3.Bind(wx.EVT_BUTTON, self.OutputFolder)
        self.btn4 = wx.Button(self.pnl, -1, "Run Find and Replace")
        self.vbox.Add(self.btn4, 0, wx.ALIGN_CENTER)
        self.btn4.Bind(wx.EVT_BUTTON, self.runReplace)

    def InitUI(self):
        self.makeMenuBar()
        self.CreateStatusBar()
        self.SetStatusText('Waiting')
        self.Centre()

    def makeMenuBar(self):
        fileMenu = wx.Menu()
        exitItem = fileMenu.Append(wx.ID_EXIT)
        helpMenu = wx.Menu()
        aboutItem = helpMenu.Append(wx.ID_ABOUT)
        menuBar = wx.MenuBar()
        menuBar.Append(fileMenu, "&File")
        menuBar.Append(helpMenu, "&About")
        self.SetMenuBar(menuBar)
        self.Bind(wx.EVT_MENU, self.OnExit, exitItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

    def OnExit(self, event):
        self.Close(True)

    def OnAbout(self, event):
        wx.MessageBox(message="Find and Replace Utility Vers 1.0 March 2019\nby Chris Clunie for SilverCloud Inc.",
                      caption="About",
                      style=wx.OK|wx.ICON_INFORMATION)

    def OnClicked(self, event):
        dlg = wx.FileDialog(self, "Choose source file", wildcard="JSON files (*.json)|*.json",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        dlg.ShowModal()
        self.SetStatusText(f'Path: {dlg.GetPath()}')
        self.file = dlg.GetPath()

    def OnClicked2(self, event):
        dlg = wx.FileDialog(self, "Choose source file", wildcard="All files (*.csv; *.xls; *.xlsx; *.json)|\
        *.csv; *.xls; *.xlsx; *.json",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        dlg.ShowModal()
        self.SetStatusText(f'Path: {dlg.GetPath()}')
        self.file2 = dlg.GetPath()

    def OutputFolder(self, event):
        dlg = wx.DirDialog(None, "Choose output directory", "",
                           wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        dlg.ShowModal()
        self.SetStatusText(f'Save As: {dlg.GetPath()}')
        self.path = dlg.GetPath()

    def runReplace(self, event):
        self.SetStatusText('Generating Files')
        os.chdir(self.path)
        mainapp = App()
        msg = mainapp.main(self.file, self.file2)
        self.SetStatusText(msg)


if __name__ == '__main__':

    app = wx.App()
    frm = GUI(None, title='Bulk Find and Replace')
    frm.SetIcon(wx.Icon('MigrateCloud.png', wx.BITMAP_TYPE_PNG))
    frm.Show()
    app.MainLoop()
