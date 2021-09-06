Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.ComponentModel.Composition.Hosting
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO.File
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports FastColoredTextBoxNS
Imports SS.Ynote.Classic.Core
Imports SS.Ynote.Classic.Core.Extensibility
Imports SS.Ynote.Classic.Core.Project
Imports SS.Ynote.Classic.Core.Search
Imports SS.Ynote.Classic.Core.Settings
Imports SS.Ynote.Classic.Core.Syntax
Imports SS.Ynote.Classic.Extensibility
Imports SS.Ynote.Classic.Extensibility.Packages
Imports SS.Ynote.Classic.UI
Imports SS.Ynote.Classic
Imports WeifenLuo.WinFormsUI.Docking
Imports AutocompleteItem = AutocompleteMenuNS.AutocompleteItem
Imports Timer = System.Windows.Forms.Timer

Public Class Form1
    Inherits Form
    Implements IYnote

#Region "Private Fields"

    Private toolBar As ToolStrip
    Private projectPanel As ProjectPanel
    Private input As InputWindow
    Private _incrementalSearcher As IncrementalSearcher
    Private _mru As Queue(Of String)
    Private _projs As IList(Of String)

#End Region

#Region "Properties"

    Private ReadOnly Property ActiveEditor As Editor
        Get
            Return TryCast(dock.ActiveDocument, Editor)
        End Get
    End Property

    Public Property Panel As DockPanel Implements IYnote.Panel

    '   Public Property Menu As MainMenu Implements IYnote.MainMenu

#End Region

#Region "Constructor"

    Public Sub New()

        Dim CachePath As String = IO.Path.Combine(IO.Path.GetDirectoryName(Application.ExecutablePath), "Ynote Classic")

        If IO.Directory.Exists(CachePath) Then
            GlobalSettings.SettingsDir = CachePath & "\"
        End If

        Globals.Settings = GlobalSettings.Load(GlobalSettings.SettingsDir & "User.ynotesettings")

        InitializeComponent()
        InitSettings()
        Panel = dock
        LoadPlugins()
        If Globals.Settings.ShowStatusBar Then InitTimer()
        Globals.Ynote = Me

        Dim arguments As String() = Environment.GetCommandLineArgs()
        Dim FilesOpen As Integer = 0

        For Each argument As String In arguments
            If String.IsNullOrEmpty(argument) Then
                OpenFile(argument)
                FilesOpen += 1
            End If
        Next

        If FilesOpen = 0 Then
            CreateNewDoc()
        End If

    End Sub

#End Region

#Region "Methods"

#Region "FILE/IO"

    Public Sub CreateNewDoc() Implements IYnote.CreateNewDoc
        Dim edit = New Editor()
        edit.Text = "untitled"
        edit.Show(Panel)
    End Sub

    Public Sub OpenFileHandler(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim MenuItem As String = sender.Text.ToString()
        If IO.File.Exists(MenuItem) Then
            OpenFile(MenuItem)
        End If
    End Sub

    Public Sub OpenFile(ByVal file As String) Implements IYnote.OpenFile
        If IO.File.Exists(file) Then
            Dim edit = OpenEditor(file)
            If edit IsNot Nothing Then edit.Show(dock, DockState.Document)
        End If
    End Sub

    Public Sub SaveEditor(ByVal edit As Editor) Implements IYnote.SaveEditor
        SaveEditor(edit, Encoding.GetEncoding(Globals.Settings.DefaultEncoding))
    End Sub

    Private Function OpenEditor(ByVal file As String) As Editor
        If Not IO.File.Exists(file) Then
            Dim result As DialogResult = MessageBox.Show("Cannot find File " & file & vbLf & "Would you like to create it ?", "Ynote Classic", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                IO.File.Create(file).Dispose()
                Return OpenEditor(file)
            End If

            Return Nothing
        End If

        For Each content As DockContent In dock.Contents
            If content.Name = file Then Return TryCast(content, Editor)
        Next

        Dim edit = New Editor()
        edit.Name = file
        If FileTypes.FileTypesDictionary Is Nothing Then FileTypes.BuildDictionary()
        Dim lang = FileTypes.GetLanguage(FileTypes.FileTypesDictionary, IO.Path.GetExtension(file))
        edit.Text = IO.Path.GetFileName(file)
        edit.Tb.Language = lang
        edit.HighlightSyntax(lang)
        Dim encoding As Encoding = If(EncodingDetector.DetectTextFileEncoding(file), Encoding.GetEncoding(Globals.Settings.DefaultEncoding))
        Dim info = New IO.FileInfo(file)

        If info.Length > 1024 * 5 * 1024 Then
            edit.Tb.OpenBindingFile(file, encoding)
        Else
            edit.Tb.OpenFile(file, encoding)
        End If

        edit.Tb.[ReadOnly] = info.IsReadOnly
        edit.RebuildAutocompleteMenu()
        Return edit
    End Function

    Private Sub OpenFileAsync(ByVal name As String)
        Me.BeginInvoke(Sub()
                           OpenFile(name)
                       End Sub)

        ' BeginInvoke(CType((Function() OpenFile(name)), MethodInvoker))
    End Sub

    Private Shared Function BuildDialogFilter(ByVal lang As String, ByVal dlg As FileDialog) As String
            Dim builder = New StringBuilder()
            builder.Append("All Files (*.*)|*.*|Text Files (*.txt)|*.txt")
            If FileTypes.FileTypesDictionary Is Nothing Then FileTypes.BuildDictionary()

            For i As Integer = 0 To FileTypes.FileTypesDictionary.Count() - 1
                Dim keyval = "*" & FileTypes.FileTypesDictionary.Keys.ElementAt(i).ElementAt(0)
                builder.AppendFormat("|{0} Files ({1})|{1}", FileTypes.FileTypesDictionary.Values.ElementAt(i), keyval)
                If lang = FileTypes.FileTypesDictionary.Values.ElementAt(i) Then dlg.FilterIndex = i + 3
            Next

            Return builder.ToString()
        End Function

    Private Sub SaveEditor(ByVal edit As Editor, ByVal encoding As Encoding)
        Try
            Dim fileName As String

            If Not edit.IsSaved Then

                Using s = New SaveFileDialog()
                    s.Title = "Save " & edit.Text
                    s.Filter = BuildDialogFilter(edit.Tb.Language, s)
                    If s.ShowDialog() <> DialogResult.OK Then Return
                    fileName = s.FileName
                End Using
            Else
                fileName = edit.Name
            End If

            If Globals.Settings.UseTabs Then
                Dim tabSpaces As String = New String(" "c, edit.Tb.TabLength)
                Dim tx = edit.Tb.Text
                Dim text As String = Regex.Replace(tx, tabSpaces, vbTab)
                IO.File.WriteAllText(fileName, text)
            Else
                edit.Tb.SaveToFile(fileName, encoding)
            End If

            edit.Text = IO.Path.GetFileName(fileName)
            edit.Name = fileName
            Trace("Saved File to " & fileName, 100000)
        Catch ex As Exception
            MessageBox.Show("Error Saving File !!" & vbLf & ex.Message, Nothing, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub RevertFile()
        If ActiveEditor Is Nothing OrElse Not ActiveEditor.IsSaved Then Return
        ActiveEditor.Tb.OpenFile(ActiveEditor.Name)
        ActiveEditor.Text = IO.Path.GetFileName(ActiveEditor.Name)
    End Sub

#End Region

#Region "Recent Handlers"

    Private Sub SaveRecentFiles()
        If _mru Is Nothing Then LoadRecentList()

        While _mru.Count > Convert.ToInt32(Globals.Settings.RecentFileNumber)
            _mru.Dequeue()
        End While

        Using stringToWrite = New IO.StreamWriter(GlobalSettings.SettingsDir & "Recent.info")

            For Each item In _mru
                stringToWrite.WriteLine(item)
            Next

            stringToWrite.Flush()
            stringToWrite.Close()
        End Using
    End Sub

    Private Sub LoadRecentList()
        Dim rfPath = GlobalSettings.SettingsDir & "Recent.info"
        If Not IO.Directory.Exists(GlobalSettings.SettingsDir) Then IO.Directory.CreateDirectory(GlobalSettings.SettingsDir)
        If Not IO.File.Exists(rfPath) Then IO.File.WriteAllText(rfPath, String.Empty)
        _mru = New Queue(Of String)()

        Using listToRead = New IO.StreamReader(rfPath)
            Dim line As String = listToRead.ReadLine()

            Do While line IsNot Nothing
                _mru.Enqueue(line)
                line = listToRead.ReadLine()
            Loop

            listToRead.Close()
        End Using

        For Each item In _mru
            recentfilesmenu.MenuItems.Add(item, New EventHandler(AddressOf OpenFileHandler))
        Next

    End Sub

    Private Sub AddRecentFile(ByVal name As String)
        If _mru Is Nothing Then LoadRecentList()

        If Not _mru.Contains(name) Then
            Dim fileRecent = New MenuItem(name)
            _mru.Enqueue(name)
            AddHandler fileRecent.Click, AddressOf OpenFileHandler
            recentfilesmenu.MenuItems.Add(fileRecent)
        End If
    End Sub

    Private Sub LoadRecentProjects()
        _projs = New List(Of String)()
        Dim file As String = GlobalSettings.SettingsDir & "Projects.ynote"
        If IO.File.Exists(file) Then
            For Each line In IO.File.ReadAllLines(file)
                _projs.Add(line)
            Next
        End If
    End Sub

    Private Sub SaveRecentProjects()
        Dim file As String = GlobalSettings.SettingsDir & "Projects.ynote"
        IO.File.WriteAllLines(file, _projs.ToArray())
    End Sub

#End Region

#Region "Menu Builders"

    Private Sub BuildLangMenu()
        For Each m In SyntaxHighlighter.Scopes.[Select](Function(lang) New MenuItem(lang))
            AddHandler m.Click, AddressOf LangMenuItemClicked
            milanguage.MenuItems.Add(m)
        Next

        milanguage.GetMenuByName("Text").Checked = True
    End Sub

    Private Sub LangMenuItemClicked(ByVal sender As Object, ByVal e As EventArgs)
        Dim item = TryCast(sender, MenuItem)
        If item Is Nothing Then Return

        For Each t As MenuItem In item.Parent.MenuItems
            t.Checked = False
        Next

        item.Checked = True
        If ActiveEditor Is Nothing Then Return
        Dim lang = item.Text
        ActiveEditor.HighlightSyntax(lang)
        ActiveEditor.Tb.Language = lang
        langmenu.Text = item.Text
    End Sub

    Private Shared Function TrimPunctuation(ByVal value As String) As String
        Dim removeFromStart = 0

        For Each t In value

            If Char.IsPunctuation(t) Then
                removeFromStart += 1
            Else
                Exit For
            End If
        Next

        Dim removeFromEnd = 0

        For i = value.Length - 1 To 0

            If Char.IsPunctuation(value(i)) Then
                removeFromEnd += 1
            Else
                Exit For
            End If
        Next

        If removeFromStart = 0 AndAlso removeFromEnd = 0 Then
            Return value
        End If

        If removeFromStart = value.Length AndAlso removeFromEnd = value.Length Then
            Return String.Empty
        End If

        Return value.Substring(removeFromStart, value.Length - removeFromEnd - removeFromStart)
    End Function

#End Region

#Region "MISC"

    Public Sub Trace(ByVal message As String, ByVal timeout As Integer) Implements IYnote.Trace
        infotimer.Stop()
        ThreadPool.QueueUserWorkItem(Function()
                                         mistats.Text = message
                                         status.Invalidate()
                                         Thread.Sleep(timeout)
                                     End Function)
        infotimer.Start()
    End Sub

    Private Shared Function ConvertToText(ByVal rtf As String) As String
        Using rtb = New RichTextBox()
            rtb.Rtf = rtf
            Return rtb.Text
        End Using
    End Function

    Public Sub AskInput(ByVal caption As String, ByVal EnterInput As InputWindow.GotInputEventHandler) Implements IYnote.AskInput
        If input Is Nothing Then
            input = New InputWindow()
            input.Dock = DockStyle.Bottom
            Me.Controls.Add(input)
            input.BringToFront()
            dock.BringToFront()
        End If

        If Not input.Visible Then
            input.Visible = True
            input.BringToFront()
        End If

        input.InitInput(caption, EnterInput)
        input.Focus()
    End Sub

    Private Sub InitSettings()
        If Not Globals.Settings.ShowMenuBar Then ToggleMenu(False)
        dock.DocumentStyle = Globals.Settings.DocumentStyle
        dock.DocumentTabStripLocation = Globals.Settings.TabLocation
        mihiddenchars.Checked = Globals.Settings.HiddenChars
        status.Visible = (statusbarmenuitem.Checked = Globals.Settings.ShowStatusBar)

        If Globals.Settings.ShowToolBar Then
            toolBar = New ToolStrip()
            toolBar.RenderMode = ToolStripRenderMode.System
            toolBar.Dock = DockStyle.Top

            If YnoteToolbar.ToolBarExists() Then
                YnoteToolbar.AddItems(toolBar)
                Controls.Add(toolBar)
                mitoolbar.Checked = True
            Else
                MessageBox.Show("Can't Find ToolBar File. Please Download a Tool bar Package to use it")
            End If
        End If
    End Sub

    Private Sub ToggleMenu(ByVal visible As Boolean)
        For Each menu As MenuItem In menu.MenuItems
            menu.Visible = visible
        Next
    End Sub

    Private infotimer As Timer

    Private Sub InitTimer()
        infotimer = New Timer()
        infotimer.Interval = 500
        AddHandler infotimer.Tick, AddressOf UpdateDocumentInfo
        infotimer.Start()
    End Sub

    Private Sub UpdateDocumentInfo()
        ThreadPool.QueueUserWorkItem(Function()

                                         Try
                                             If Not (TypeOf dock.ActiveDocument Is Editor) OrElse ActiveEditor Is Nothing Then

                                             Else
                                                 If ActiveEditor.Tb.Selection.IsEmpty Then
                                                     Dim nCol As Integer = ActiveEditor.Tb.Selection.Start.iChar + 1
                                                     Dim line As Integer = ActiveEditor.Tb.Selection.Start.iLine + 1
                                                     mistats.Text = String.Format("Line {0}, Column {1}", line, nCol)
                                                 Else
                                                     mistats.Text = String.Format("{0} Characters Selected", ActiveEditor.Tb.SelectedText.Length)
                                                 End If
                                             End If

                                         Catch
                                         End Try
                                     End Function)
    End Sub

    Private Shared Function SortByLength(ByVal e As IEnumerable(Of String)) As IEnumerable(Of String)
        Dim sorted = From s In e Order By s.Length Select s
        Return sorted
    End Function

#End Region

#Region "Plugins"

    Private Sub LoadPlugins()
        Dim directory As String = GlobalSettings.SettingsDir & "\Plugins"
        If Not IO.Directory.Exists(directory) Then IO.Directory.CreateDirectory(directory)

        Using dircatalog = New DirectoryCatalog(directory, "*.dll")

            Using container = New CompositionContainer(dircatalog)
                Dim plugins As IEnumerable(Of IYnotePlugin) = container.GetExportedValues(Of IYnotePlugin)()

                For Each plugin As IYnotePlugin In plugins
                    plugin.Main(Me)
                Next
            End Using
        End Using
    End Sub

#End Region

#Region "Layouts"

    Private Function GetContentFromPersistString(ByVal persistString As String) As IDockContent
        Dim parsedStrings As String() = persistString.Split({","c})

        If parsedStrings(0) = GetType(ProjectPanel).ToString() Then
            Dim projp = New ProjectPanel()
            If parsedStrings(1) <> "ProjectPanel" Then projp.OpenProject(YnoteProject.Load(parsedStrings(1)))
            Return projp
        End If

        If parsedStrings(0) = GetType(Editor).ToString() Then
            If parsedStrings(1) = "Editor" Then Return Nothing
            Return OpenEditor(parsedStrings(1))
        End If

        Return Nothing
    End Function

#End Region

#End Region

#Region "Overrides"

    Protected Overrides Sub OnClosing(ByVal e As CancelEventArgs)
        SaveRecentFiles()
        GlobalSettings.Save(Globals.Settings, GlobalSettings.SettingsDir & "User.ynotesettings")
        If _projs IsNot Nothing Then SaveRecentProjects()
        If Globals.ActiveProject IsNot Nothing AndAlso Globals.ActiveProject.IsSaved Then dock.SaveAsXml(Globals.ActiveProject.LayoutFile)
        MyBase.OnClosing(e)
    End Sub

    Protected Overrides Sub OnResize(ByVal e As EventArgs)
        If Globals.Settings.MinimizeToTray Then
            Dim nicon = New NotifyIcon()
            nicon.Icon = Icon
            AddHandler nicon.DoubleClick, Function()
                                              Show()
                                              WindowState = FormWindowState.Normal
                                          End Function

            nicon.BalloonTipIcon = ToolTipIcon.Info
            nicon.BalloonTipTitle = "Ynote Classic"
            nicon.BalloonTipText = "Ynote Classic has minimized to the System Tray"

            If FormWindowState.Minimized = WindowState Then
                nicon.Visible = True
                nicon.ShowBalloonTip(500)
                Hide()
            ElseIf FormWindowState.Normal = WindowState Then
                nicon.Visible = False
            End If
        End If

        MyBase.OnResize(e)
    End Sub

#End Region

#Region "Events"

    Private path As String

    Public Function GetActiveProject() As YnoteProject Implements IYnote.GetActiveProject
        Return Globals.ActiveProject
    End Function

    Private Sub NewMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles NewMenuItem.Click
        CreateNewDoc()
    End Sub

    Private Sub OpenMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OpenMenuItem.Click
        Using dialog = New OpenFileDialog()
            dialog.Filter = "All Files (*.*)|*.*"
            dialog.Multiselect = True
            Dim res = dialog.ShowDialog() = DialogResult.OK
            If Not res Then Return

            For Each file In dialog.FileNames
                OpenFileAsync(file)
                AddRecentFile(file)
            Next
        End Using
    End Sub

    Private Sub UndoMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles UndoMenuItem.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Undo()
    End Sub

    Private Sub RedoMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles RedoMenuItem.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Redo()
    End Sub

    Private Sub CutMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CutMenuItem.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Cut()
    End Sub

    Private Sub CopyMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CopyMenuItem.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Copy()
    End Sub

    Private Sub PasteMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PasteMenuItem.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Paste()
    End Sub

    Private Sub mifind_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mifind.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.ShowFindDialog()
    End Sub

    Private Sub replacemenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles replacemenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.ShowReplaceDialog()
    End Sub

    Private Sub increaseindent_Click(ByVal sender As Object, ByVal e As EventArgs) Handles increaseindent.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.IncreaseIndent()
    End Sub

    Private Sub decreaseindent_Click(ByVal sender As Object, ByVal e As EventArgs) Handles decreaseindent.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.DecreaseIndent()
    End Sub

    Private Sub doindent_Click(ByVal sender As Object, ByVal e As EventArgs) Handles doindent.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.DoAutoIndent()
    End Sub

    Private Sub gotofirstlinemenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles gotofirstlinemenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.GoHome()
    End Sub

    Private Sub gotoendmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles gotoendmenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.GoEnd()
    End Sub

    Private Sub navforwardmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles navforwardmenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.NavigateForward()
    End Sub

    Private Sub navbackwardmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles navbackwardmenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.NavigateBackward()
    End Sub

    Private Sub selectallmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles selectallmenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.SelectAll()
    End Sub

    Private Sub foldallmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles foldallmenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.CollapseAllFoldingBlocks()
    End Sub

    Private Sub unfoldmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles unfoldmenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.ExpandAllFoldingBlocks()
    End Sub

    Private Sub foldselected_Click(ByVal sender As Object, ByVal e As EventArgs) Handles foldselected.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.CollapseBlock(ActiveEditor.Tb.Selection.Start.iLine, ActiveEditor.Tb.Selection.[End].iLine)
    End Sub

    Private Sub unfoldselected_Click(ByVal sender As Object, ByVal e As EventArgs) Handles unfoldselected.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.CollapseBlock(ActiveEditor.Tb.Selection.Start.iLine, ActiveEditor.Tb.Selection.[End].iLine)
    End Sub

    Private Sub datetime_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles datetime.Click
        Dim dateEx = DateTime.Now
        Dim time = New TimeSpan(36, 0, 0, 0)
        Dim combined = dateEx.Add(time).ToString("DD/MM/YYYY")
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.InsertText(combined)
    End Sub

    Private Sub fileastext_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles fileastext.Click
        Using openfile = New OpenFileDialog()
            openfile.Filter = "All Files (*.*)|*.*|Text Documents(*.txt)|*.txt"
            openfile.ShowDialog()
            If openfile.FileName <> "" AndAlso ActiveEditor IsNot Nothing Then ActiveEditor.Tb.InsertText(IO.File.ReadAllText(openfile.FileName))
        End Using
    End Sub

    Private Sub filenamemenuitem_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles filenamemenuitem.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.InsertText(ActiveEditor.Text)
    End Sub

    Private Sub fullfilenamemenuitem_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles fullfilenamemenuitem.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.InsertText(ActiveEditor.Name)
    End Sub

    Private Sub mifindinfiles_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mifindinfiles.Click
        Using findinfiles = New FindInFiles()
            findinfiles.StartPosition = FormStartPosition.CenterParent
            findinfiles.ShowDialog(Me)
        End Using
    End Sub

    Private Sub mifindinproject_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mifindinproject.Click
        Using findinfiles = New FindInFiles()
            If Globals.ActiveProject Is Nothing Then Return
            findinfiles.Directory = Globals.ActiveProject.Path
            findinfiles.StartPosition = FormStartPosition.CenterParent
            findinfiles.ShowDialog(Me)
        End Using
    End Sub

    Private Sub mifindchar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mifindchar.Click
        MessageBox.Show("Press Alt+F+{char} you want to find", "Ynote Classic", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub movelineup_Click(ByVal sender As Object, ByVal e As EventArgs) Handles movelineup.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.MoveSelectedLinesUp()
    End Sub

    Private Sub movelinedown_Click(ByVal sender As Object, ByVal e As EventArgs) Handles movelinedown.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.MoveSelectedLinesDown()
    End Sub

    Private Sub duplicatelinemenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles duplicatelinemenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.DuplicateLine(ActiveEditor.Tb.Selection.Start.iLine)
    End Sub

    Private Sub removeemptylines_Click(ByVal sender As Object, ByVal e As EventArgs) Handles removeemptylines.Click
        If ActiveEditor IsNot Nothing Then
            Dim iLines = ActiveEditor.Tb.FindLines("^\s*$", RegexOptions.None)
            ActiveEditor.Tb.RemoveLines(iLines)
        End If
    End Sub

    Private Sub caseuppermenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles caseuppermenu.Click
        If ActiveEditor IsNot Nothing Then
            Dim upper = ActiveEditor.Tb.SelectedText.ToUpper()
            ActiveEditor.Tb.SelectedText = upper
        End If
    End Sub

    Private Sub caselowermenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles caselowermenu.Click
        If ActiveEditor Is Nothing Then Return
        Dim lower = ActiveEditor.Tb.SelectedText.ToLower()
        ActiveEditor.Tb.SelectedText = lower
    End Sub

    Private Sub casetitlemenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles casetitlemenu.Click
        Dim cultureinfo = Thread.CurrentThread.CurrentCulture
        Dim info = cultureinfo.TextInfo
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.SelectedText = info.ToTitleCase(ActiveEditor.Tb.SelectedText)
    End Sub

    Private Sub swapcase_Click(ByVal sender As Object, ByVal e As EventArgs) Handles swapcase.Click
        If ActiveEditor Is Nothing Then Return
        Dim input = ActiveEditor.Tb.SelectedText
        Dim reversedCase = New String(input.[Select](Function(c) If(Char.IsLetter(c), (If(Char.IsUpper(c), Char.ToLower(c), Char.ToUpper(c))), c)).ToArray())
        ActiveEditor.Tb.SelectedText = reversedCase
    End Sub

    Private Sub Addbookmarkmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Addbookmarkmenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Bookmarks.Add(ActiveEditor.Tb.Selection.Start.iLine)
    End Sub

    Private Sub removebookmarkmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles removebookmarkmenu.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Bookmarks.Remove(ActiveEditor.Tb.Selection.Start.iLine)
    End Sub

    Private Sub navigatethroughbookmarks_Click(ByVal sender As Object, ByVal e As EventArgs) Handles navigatethroughbookmarks.Click
        MessageBox.Show("Press Ctrl + Shift + N To Navigate through bookmarks")
    End Sub

    Private Sub rtfExport_Click(ByVal sender As Object, ByVal e As EventArgs) Handles rtfExport.Click
        Using rtfs = New SaveFileDialog()
            rtfs.Filter = "Rich Text Documents (*.rtf)|*.rtf"
            rtfs.ShowDialog()

            If rtfs.FileName <> "" Then
                If ActiveEditor IsNot Nothing Then IO.File.WriteAllText(rtfs.FileName, ActiveEditor.Tb.Rtf)
            End If
        End Using
    End Sub

    Private Sub htmlexport_Click(ByVal sender As Object, ByVal e As EventArgs) Handles htmlexport.Click
        Using htmls = New SaveFileDialog()
            htmls.FileName = "HTML Web Page (*.htm), (*.html)|*.htm|Shtml Page (*.shtml)|*.shtml"
            Dim result = htmls.ShowDialog() = DialogResult.OK

            If result Then
                If ActiveEditor IsNot Nothing Then IO.File.WriteAllText(htmls.FileName, ActiveEditor.Tb.Html)
            End If
        End Using
    End Sub

    Private Sub pngexport_Click(ByVal sender As Object, ByVal e As EventArgs) Handles pngexport.Click
        If ActiveEditor Is Nothing Then Return
        Dim bmp = New Bitmap(ActiveEditor.Tb.Width, ActiveEditor.Tb.Height)
        ActiveEditor.Tb.DrawToBitmap(bmp, New Rectangle(0, 0, ActiveEditor.Tb.Width, ActiveEditor.Tb.Height))

        Using pngs = New SaveFileDialog()
            pngs.Filter = "Portable Network Graphics (*.png)|*.png|JPEG (*.jpg)|*.jpg"
            pngs.ShowDialog()
            If Not String.IsNullOrEmpty(pngs.FileName) Then bmp.Save(pngs.FileName)
        End Using
    End Sub

    Private Sub fromrtf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fromrtf.Click
        Using o = New OpenFileDialog With {
            .Filter = "RTF Files (*.rtf)|*.rtf"
        }
            o.ShowDialog()
            If o.FileName = "" Then Return
            Dim edit = New Editor()
            edit.Tb.Text = ConvertToText(IO.File.ReadAllText(o.FileName))
            edit.Name = o.FileName
            edit.Text = IO.Path.GetFileName(o.FileName)
            edit.Show(dock, DockState.Document)
        End Using
    End Sub

    Private Sub miproperties_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miproperties.Click
        If ActiveEditor IsNot Nothing AndAlso ActiveEditor.IsSaved Then
            NativeMethods.ShowFileProperties(ActiveEditor.Name)
        Else
            MessageBox.Show("File Not Saved!", "Ynote Classic")
        End If
    End Sub

    Private Sub ExitMenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitMenu.Click
        Application.[Exit]()
    End Sub

    Private Sub mizoomin_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mizoomin.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Zoom += 10
    End Sub

    Private Sub mizoomout_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mizoomout.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Zoom -= 10
    End Sub

    Private Sub mirestoredefault_Click(ByVal sender As Object, ByVal e As EventArgs) ' Handles mirestoredefault.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Zoom = 100
    End Sub

    Private Sub mitransparent_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mitransparent.Click
        mitransparent.Checked = Not mitransparent.Checked
        Opacity = If(mitransparent.Checked, 0.7, 1)
    End Sub

    Private Sub mifullscreen_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mifullscreen.Click
        mifullscreen.Checked = Not mifullscreen.Checked

        If Not mifullscreen.Checked Then
            FormBorderStyle = FormBorderStyle.Sizable
            WindowState = FormWindowState.Normal
        Else
            FormBorderStyle = FormBorderStyle.None
            WindowState = FormWindowState.Maximized
        End If
    End Sub

    Private Sub wordwrapmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles wordwrapmenu.Click
        wordwrapmenu.Checked = Not wordwrapmenu.Checked

        For Each document In dock.Documents
            If TypeOf document Is Editor Then
                Dim documentEx As Editor = TryCast(document, Editor)
                documentEx.Tb.WordWrap = wordwrapmenu.Checked
            End If
        Next

        Globals.Settings.WordWrap = wordwrapmenu.Checked
    End Sub

    Private Sub aboutmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles aboutmenu.Click
        Using ab = New About()
            ab.ShowDialog()
        End Using
    End Sub

    Private Sub mifb_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mifb.Click
        Process.Start("http://twitter.com/ynoteclassic")
    End Sub

    Private Sub miwikimenu_Click(ByVal sender As Object, ByVal e As EventArgs) ' Handles miwikimenu.Click
        Process.Start("http://ynoteclassic.codeplex.com/documentation")
    End Sub

    Private Sub zoom_DropDownItemClicked(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles zoom.DropDownItemClicked
        Dim i = e.ClickedItem.Text.ToInt()
        If ActiveEditor Is Nothing Then Return
        ActiveEditor.Tb.Zoom = i
        Globals.Settings.Zoom = i
    End Sub

    Private Sub pluginmanagermenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles pluginmanagermenu.Click
        Using manager = New PackageManager With {
            .StartPosition = FormStartPosition.CenterParent
        }
            manager.ShowDialog(Me)
        End Using
    End Sub

    Private Sub savemenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles savemenu.Click
        Me.BeginInvoke(Sub()
                           SaveEditor(ActiveEditor)
                       End Sub)
        '  BeginInvoke(CType((Function() SaveEditor(ActiveEditor)), MethodInvoker))
    End Sub

    Private Sub misaveas_Click(ByVal sender As Object, ByVal e As EventArgs) Handles misaveas.Click
        If ActiveEditor Is Nothing Then Return

        Using sf = New SaveFileDialog()
            sf.Filter = BuildDialogFilter(ActiveEditor.Tb.Language, sf)
            If sf.ShowDialog() <> DialogResult.OK Then Return
            ActiveEditor.Tb.SaveToFile(sf.FileName, Encoding.GetEncoding(Globals.Settings.DefaultEncoding))
            ActiveEditor.Text = IO.Path.GetFileName(sf.FileName)
            ActiveEditor.Name = sf.FileName
        End Using
    End Sub

    Private Sub misaveall_Click(ByVal sender As Object, ByVal e As EventArgs) Handles misaveall.Click
        For Each doc As Editor In dock.Documents
            Me.BeginInvoke(Sub()
                               SaveEditor(doc)
                           End Sub)

            '  BeginInvoke(CType((Function() SaveEditor(doc)), MethodInvoker))
        Next
    End Sub

    Private Sub OptionsMenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OptionsMenu.Click
        Dim optionsdialog = New Options()
        optionsdialog.Show()
    End Sub

    Private Sub mimacrorecord_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles mimacrorecord.Click
        ActiveEditor.Tb.MacrosManager.IsRecording = Not ActiveEditor.Tb.MacrosManager.IsRecording
    End Sub

    Private Sub miExecmacro_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles miExecmacro.Click
        If ActiveEditor.Tb.MacrosManager.Macros IsNot Nothing Then ActiveEditor.Tb.MacrosManager.ExecuteMacros()
    End Sub

    Private Sub misavemacro_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles misavemacro.Click
        Using sf = New SaveFileDialog()
            sf.Filter = "Ynote Macro File(*.ynotemacro)|*.ynotemacro"
            sf.InitialDirectory = GlobalSettings.SettingsDir & "User\"
            If sf.ShowDialog() <> DialogResult.OK Then Return

            If Not ActiveEditor.Tb.MacrosManager.MacroIsEmpty Then
                IO.File.WriteAllText(sf.FileName, ActiveEditor.Tb.MacrosManager.Macros)
            Else
                MessageBox.Show("Macro Is Empty!")
            End If
        End Using
    End Sub

    Private Sub miclearmacro_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles miclearmacro.Click
        ActiveEditor.Tb.MacrosManager.ClearMacros()
    End Sub

    Private Sub menuItem30_Click(ByVal sender As Object, ByVal e As EventArgs) ' Handles menuItem30.Click
        Try

            If ActiveEditor IsNot Nothing AndAlso ActiveEditor.IsSaved Then
                SaveEditor(ActiveEditor)

                If IO.Path.GetExtension(ActiveEditor.Name) = ".ys" Then
                    YnoteScript.RunScript(Me, ActiveEditor.Name)
                Else
                    Process.Start(ActiveEditor.Name)
                End If
            Else
                MessageBox.Show("File Not Saved!")
            End If

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message)
        End Try
    End Sub

    Private Sub mitoolbar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mitoolbar.Click
        If toolBar IsNot Nothing AndAlso toolBar.Items.Count > 0 Then
            toolBar.Visible = mitoolbar.Checked
            mitoolbar.Checked = Not mitoolbar.Checked
            Globals.Settings.ShowToolBar = mitoolbar.Checked
        Else

            If Not YnoteToolbar.ToolBarExists() Then
                MessageBox.Show("You do not have any toolbars installed." & vbLf & "Please Download a ToolBar Package")
            Else
                toolBar = New ToolStrip()
                toolBar.RenderMode = ToolStripRenderMode.System
                toolBar.Dock = DockStyle.Top

                If YnoteToolbar.ToolBarExists() Then
                    YnoteToolbar.AddItems(toolBar)
                    Controls.Add(toolBar)
                    mitoolbar.Checked = True
                Else
                    MessageBox.Show("Can't Find ToolBar File. Please Download a Tool bar Package to use it")
                End If
            End If
        End If
    End Sub

    Private Sub statusbarmenuitem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles statusbarmenuitem.Click
        status.Visible = Not statusbarmenuitem.Checked
        statusbarmenuitem.Checked = Not statusbarmenuitem.Checked
        Globals.Settings.ShowStatusBar = statusbarmenuitem.Checked
    End Sub

    Private Sub mincrementalsearch_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mincrementalsearch.Click
        If ActiveEditor IsNot Nothing Then
            If _incrementalSearcher Is Nothing Then
                _incrementalSearcher = New IncrementalSearcher With {
                    .Dock = DockStyle.Bottom
                }
                Controls.Add(_incrementalSearcher)
                _incrementalSearcher.Tb = ActiveEditor.Tb
                _incrementalSearcher.tbFind.Text = ActiveEditor.Tb.SelectedText
                _incrementalSearcher.FocusTextBox()
                _incrementalSearcher.BringToFront()
                dock.BringToFront()
            Else

                If _incrementalSearcher.Visible Then
                    _incrementalSearcher.[Exit]()
                Else
                    _incrementalSearcher.Tb = ActiveEditor.Tb
                    _incrementalSearcher.tbFind.Text = ActiveEditor.Tb.SelectedText
                    _incrementalSearcher.Visible = True
                    _incrementalSearcher.FocusTextBox()
                End If
            End If
        End If
    End Sub

    Private Sub mirevert_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles mirevert.Click
        RevertFile()
    End Sub

    Private Sub miprint_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miprint.Click
        ActiveEditor.Tb.Print(New PrintDialogSettings With {
            .ShowPrintDialog = True
        })
    End Sub

    Private Sub reopenclosedtab_Click(ByVal sender As Object, ByVal e As EventArgs) Handles reopenclosedtab.Click
        Try
            If _mru Is Nothing Then LoadRecentList()
            Dim recentlyclosed = _mru.Last()
            If recentlyclosed IsNot Nothing Then OpenFile(recentlyclosed)
        Catch
            recentfilesmenu.PerformSelect()
        End Try
    End Sub

    Private Sub migotoline_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles migotoline.Click
        ActiveEditor.Tb.ShowGoToDialog()
    End Sub

    Private Sub colorschemeitem_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles colorschemeitem.Click
        Dim m = TryCast(sender, MenuItem)

        For Each item As MenuItem In colorschememenu.MenuItems
            item.Checked = False
        Next

        If m IsNot Nothing Then
            m.Checked = True
            Globals.Settings.ThemeFile = m.Name

            For Each content In dock.Documents.OfType(Of Editor)()
                content.RePaintTheme()
            Next
        End If

    End Sub

    Private Sub colorschememenu_Select(ByVal sender As Object, ByVal e As EventArgs) Handles colorschememenu.Select
        If colorschememenu.MenuItems.Count <> 0 Then Return

        For Each menuitem In IO.Directory.GetFiles(GlobalSettings.SettingsDir, "*.ynotetheme", IO.SearchOption.AllDirectories).[Select](Function(file) New MenuItem With {
            .Text = IO.Path.GetFileNameWithoutExtension(file),
            .Name = file
        })
            AddHandler menuitem.Click, AddressOf colorschemeitem_Click
            colorschememenu.MenuItems.Add(menuitem)
        Next
    End Sub

    Private Sub menuItem29_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem29.Click
        If ActiveEditor IsNot Nothing Then
            ActiveEditor.Tb.SelectedText = String.Join(" ", ActiveEditor.Tb.SelectedText.Split({Environment.NewLine}, StringSplitOptions.None))
        End If
    End Sub

    Private Sub milanguage_Select(ByVal sender As Object, ByVal e As EventArgs) Handles milanguage.Select
        If milanguage Is Nothing OrElse milanguage.MenuItems.Count <> 0 Then

        Else
            BuildLangMenu()

            For Each menu As MenuItem In milanguage.MenuItems
                menu.Checked = False
            Next

            If ActiveEditor IsNot Nothing Then
                milanguage.GetMenuByName(ActiveEditor.Tb.Language).Checked = True
            End If

        End If

    End Sub

    Private Sub menuItem65_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem65.Click
        Dim fctb = ActiveEditor.Tb
        Dim lines = fctb.SelectedText.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
        Array.Sort(lines)
        Dim formedtext = "" 'lines.Aggregate(Of String, String)(Nothing, Function(current, str) current + (str & vbCrLf))
        fctb.SelectedText = formedtext
    End Sub

    Private Sub menuItem69_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem69.Click
        ActiveEditor.Tb.SortLines()
    End Sub

    Private Sub gotobookmark_Click(ByVal sender As Object, ByVal e As EventArgs) Handles gotobookmark.Click
        Dim items = New List(Of AutocompleteItem)()

        For Each bookmark In ActiveEditor.Tb.Bookmarks
            items.Add(New FuzzyAutoCompleteItem(bookmark.LineIndex + 1 & ":" + ActiveEditor.Tb(bookmark.LineIndex).Text))
        Next

        Dim bookmarkwindow = New CommandWindow(items)
        AddHandler bookmarkwindow.ProcessCommand, AddressOf bookmarkwindow_ProcessCommand
        bookmarkwindow.ShowDialog(Me)
    End Sub

    Private Sub bookmarkwindow_ProcessCommand(ByVal sender As Object, ByVal e As CommandWindowEventArgs) 'Handles bookmarkwindow.ProcessCommand
        Dim markCommand = YnoteCommand.FromString(e.Text)
        Dim index As Integer = markCommand.Key.ToInt() - 1

        For Each bookmark In ActiveEditor.Tb.Bookmarks
            If bookmark.LineIndex = index Then bookmark.DoVisible()
        Next
    End Sub

    Private Sub menuItem71_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem71.Click
        Process.Start("charmap.exe")
    End Sub

    Private Sub Split(ByVal iLine As Integer, ByVal tb As FastColoredTextBox, ByVal s As String)
        Dim sep As String() = s.Split(":"c)
        tb.Selection.Start = New Place(0, iLine)
        tb.Selection.Expand()
        Dim seperator As String

        If sep.Length = 3 AndAlso sep(1) = String.Empty Then
            seperator = ":"
        Else
            seperator = sep(1)
        End If

        Dim lines = tb.SelectedText.Split(seperator.ToCharArray())
        tb.ClearSelected()

        For Each line As String In lines
            tb.InsertText(line & vbLf)
        Next

        tb.ClearCurrentLine()
    End Sub

    Private Sub splitlinemenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles splitlinemenu.Click
        '  AskInput("Seperator:", Function(o, args) Split(ActiveEditor.Tb.Selection.Start.iLine, ActiveEditor.Tb, args.InputValue))
    End Sub

    Private Sub InsertCharacters(ByVal times As Integer, ByVal character As String)
        While times > 0
            ActiveEditor.Tb.InsertText(character)
            times -= 1
        End While
    End Sub

    Private Sub emptycolumns_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles emptycolumns.Click
        'AskInput("Number:", Function(o, args) InsertCharacters(args.GetFormattedInput().ToInt(), " "))
    End Sub

    Private Sub emptylines_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles emptylines._Click
        ' AskInput("Number:", Function(o, args) InsertCharacters(args.GetFormattedInput().ToInt(), vbLf))
    End Sub

    Private Sub mimultimacro_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles mimultimacro.Click
        If ActiveEditor Is Nothing Then Return
        Dim macrodlg = New MacroExecDialog(ActiveEditor.Tb)
        macrodlg.ShowDialog(Me)
    End Sub

    Private Sub commanderMenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles commandermenu.Click
        Using console = New Commander()
            console.LangMenu = langmenu
            console.ShowDialog(Me)
        End Using
    End Sub

    Private Sub menuItem75_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem75.Click
        If ActiveEditor.Tb.SelectedText <> "" Then
            Dim trimmed = ""
            Dim lines = ActiveEditor.Tb.SelectedText.Split({Environment.NewLine}, StringSplitOptions.None)
            trimmed = If(lines.Length <> 1, lines.Aggregate(trimmed, Function(current, line) current + (line.TrimEnd() + Environment.NewLine)), ActiveEditor.Tb.SelectedText.TrimEnd())
            ActiveEditor.Tb.SelectedText = trimmed
        Else
            MessageBox.Show("Noting Selected to Perform Function", "Ynote Classic")
        End If
    End Sub

    Private Sub menuItem76_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem76.Click
        If ActiveEditor.Tb.SelectedText <> "" Then
            Dim trimmed = String.Empty
            Dim lines = ActiveEditor.Tb.SelectedText.Split({Environment.NewLine}, StringSplitOptions.None)
            trimmed = If(lines.Length <> 1, lines.Aggregate(trimmed, Function(current, line) current + (line.TrimStart() + Environment.NewLine)), ActiveEditor.Tb.SelectedText.TrimStart())
            ActiveEditor.Tb.SelectedText = trimmed
        Else
            MessageBox.Show("Noting Selected to Perform Function", "Ynote Classic")
        End If
    End Sub

    Private Sub menuItem79_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem79.Click
        If Not String.IsNullOrEmpty(ActiveEditor.Tb.SelectedText) Then
            Dim trimmed = ""
            Dim lines = ActiveEditor.Tb.SelectedText.Split({Environment.NewLine}, StringSplitOptions.None)
            trimmed = If(lines.Length <> 1, lines.Aggregate(trimmed, Function(current, line) current + (line.Trim() + Environment.NewLine)), ActiveEditor.Tb.SelectedText.Trim())
            ActiveEditor.Tb.SelectedText = trimmed
        Else
            MessageBox.Show("Noting Selected to Perform Function", "Ynote Classic")
        End If
    End Sub

    Private Sub menuItem78_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem78.Click
        If Not String.IsNullOrEmpty(ActiveEditor.Tb.SelectedText) Then
            ActiveEditor.Tb.SelectedText = ActiveEditor.Tb.SelectedText.Replace(vbCrLf, " ")
        Else
            MessageBox.Show("Nothing Selected to Perfrom Function", "Ynote Classic")
        End If
    End Sub

    Private Sub menuItem80_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem80.Click
        If ActiveEditor.Tb.SelectedText <> "" Then
            ActiveEditor.Tb.SelectedText = ActiveEditor.Tb.SelectedText.Replace(" ", vbCrLf)
        Else
            MessageBox.Show("Nothing Selected to Perfrom Function", "Ynote Classic")
        End If
    End Sub

    Private Sub menuItem77_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem77.Click
        If ActiveEditor.Tb.SelectedText <> "" Then
            ActiveEditor.Tb.SelectedText = ActiveEditor.Tb.SelectedText.Replace(vbCrLf, String.Empty)
        Else
            MessageBox.Show("Nothing Selected to Perfrom Function", "Ynote Classic")
        End If
    End Sub

    Private Sub menuItem84_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles menuItem84.Click
        Dim items = New List(Of AutocompleteItem)()

        For Each doc In dock.Documents
            items.Add(New FuzzyAutoCompleteItem((TryCast(doc, DockContent)).Text))
        Next

        Dim fileswitcher = New CommandWindow(items)
        AddHandler fileswitcher.ProcessCommand, AddressOf fileswitcher_ProcessCommand
        fileswitcher.ShowDialog(Me)
    End Sub

    Private Sub fileswitcher_ProcessCommand(ByVal sender As Object, ByVal e As CommandWindowEventArgs)
        For Each content As DockContent In dock.Documents

            If content.Text = e.Text Then
                content.Show(dock)
            End If
        Next
    End Sub

    Private Sub mifindNext_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mifindnext.Click
        MessageBox.Show("Press F3 to Find Next")
    End Sub

    Private Sub mitrimpunctuation_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles mitrimpunctuation.Click
        If ActiveEditor.Tb.SelectedText <> "" Then
            Dim trimmed = ""
            Dim lines = ActiveEditor.Tb.SelectedText.Split({Environment.NewLine}, StringSplitOptions.None)
            trimmed = If(lines.Length <> 1, lines.Aggregate(trimmed, Function(current, line) current + (TrimPunctuation(line) + Environment.NewLine)), TrimPunctuation(ActiveEditor.Tb.SelectedText))
            ActiveEditor.Tb.SelectedText = trimmed
        Else
            MessageBox.Show("Noting Selected to Perform Function", "Ynote Classic")
        End If
    End Sub

    Private Sub mihiddenchars_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mihiddenchars.Click
        mihiddenchars.Checked = Not mihiddenchars.Checked
        Globals.Settings.HiddenChars = mihiddenchars.Checked
    End Sub

    Private Sub removelinemenu_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles removelinemenu.Click
        ActiveEditor.Tb.ClearCurrentLine()
    End Sub

    Private Sub macroitem_click(ByVal sender As Object, ByVal e As EventArgs) ' Handles macroitem.click
        Dim item = TryCast(sender, MenuItem)
        If item IsNot Nothing Then ActiveEditor.Tb.MacrosManager.ExecuteMacros(item.Name)
    End Sub

    Private Sub scriptitem_clicked(ByVal sender As Object, ByVal e As EventArgs) ' Handles scriptitem.clicked
        Dim item = TryCast(sender, MenuItem)
        If item IsNot Nothing Then YnoteScript.RunScript(Me, item.Name)
    End Sub

    Private Sub menuItem100_Click(ByVal sender As Object, ByVal e As EventArgs) ' Handles menuItem100.Click
        Dim fctb = ActiveEditor.Tb
        Dim form = New HotkeysEditorForm(fctb.HotkeysMapping)

        If form.ShowDialog() = DialogResult.OK Then
            fctb.HotkeysMapping = form.GetHotkeys()
            IO.File.WriteAllText(GlobalSettings.SettingsDir & "Editor.ynotekeys", form.GetHotkeys().ToString())
        End If
    End Sub

    Private Sub langmenu_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles langmenu.Click
        If langmenu.HasDropDownItems Then
            Dim item = langmenu.DropDownItems.Cast(Of ToolStripMenuItem)().FirstOrDefault(Function(c) ActiveEditor IsNot Nothing AndAlso c.Text = ActiveEditor.Tb.Language.ToString())
            If item IsNot Nothing Then item.Checked = True
        End If
    End Sub

    Private Sub langmenu_MouseEnter(ByVal sender As Object, ByVal e As EventArgs) Handles langmenu.MouseEnter
        If langmenu.DropDownItems.Count <> 0 Then Return
        If SyntaxHighlighter.Scopes Is Nothing Then Return

        For Each m In SyntaxHighlighter.Scopes.[Select](Function(lang) New ToolStripMenuItem(lang.ToString()))
            AddHandler m.Click, AddressOf langitem_Click
            langmenu.DropDownItems.Add(m)
        Next
    End Sub

    Private Sub langitem_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles langitem.Click
        Dim item = TryCast(sender, ToolStripMenuItem)
        Dim lang = item.Text
        ActiveEditor.Tb.Language = lang
        ActiveEditor.HighlightSyntax(ActiveEditor.Tb.Language)
        langmenu.Text = item.Text
    End Sub

    Private Sub mirunscripts_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles mirunscripts_Click
        If ActiveEditor Is Nothing Then Return

        If ActiveEditor.IsSaved Then
            MessageBox.Show("Use Ctrl+Shift+P and type Run:")
        Else
            MessageBox.Show("Please Save the Current Document before proceeding!", "RunScript Executor")
        End If
    End Sub

    Private Sub dock_ActiveDocumentChanged(ByVal sender As Object, ByVal e As EventArgs) Handles dock.ActiveDocumentChanged
        If ActiveEditor Is Nothing Then Return
        langmenu.Text = ActiveEditor.Tb.Language
        If _incrementalSearcher IsNot Nothing AndAlso _incrementalSearcher.Visible Then _incrementalSearcher.Tb = ActiveEditor.Tb
    End Sub

    Private Sub migoleftbracket_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles migoleftbracket.Click
        ActiveEditor.Tb.GoLeftBracket()
    End Sub

    Private Sub migorightbracket_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles migorightbracket.Click
        ActiveEditor.Tb.GoRightBracket()
    End Sub

    Private Sub misortlength_Click(ByVal sender As Object, ByVal e As EventArgs) Handles misortlength.Click
        If ActiveEditor IsNot Nothing Then
            Dim lines = ActiveEditor.Tb.SelectedText.Split({Environment.NewLine}, StringSplitOptions.None)
            Dim results = SortByLength(lines)
            ' ActiveEditor.Tb.SelectedText = results.Aggregate(Of String, String)(Nothing, Function(current, result) current + (result & vbCrLf))
        End If
    End Sub

    Private Sub mimarkRed_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles mimarkRed.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Selection.SetStyle(New TextStyle(Nothing, New SolidBrush(Color.FromArgb(180, Color.Red)), FontStyle.Regular))
    End Sub

    Private Sub mimarkblue_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles mimarkblue.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Selection.SetStyle(New TextStyle(Nothing, New SolidBrush(Color.FromArgb(180, Color.Blue)), FontStyle.Regular))
    End Sub

    Private Sub mimarkgray_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles mimarkgray.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Selection.SetStyle(New TextStyle(Nothing, New SolidBrush(Color.FromArgb(180, Color.Gainsboro)), FontStyle.Regular))
    End Sub

    Private Sub mimarkgreen_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles mimarkgreen.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Selection.SetStyle(New TextStyle(Nothing, New SolidBrush(Color.FromArgb(180, Color.Green)), FontStyle.Regular))
    End Sub

    Private Sub mimarkyellow_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mimarkyellow.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Selection.SetStyle(New TextStyle(Nothing, New SolidBrush(Color.FromArgb(180, Color.Yellow)), FontStyle.Regular))
    End Sub

    Private Sub miclearMarked_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miclearmarked.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.Selection.ClearStyle(StyleIndex.All)
    End Sub

    Private Sub misnippets_Click(ByVal sender As Object, ByVal e As EventArgs) Handles misnippets.Click
        ActiveEditor.ForceAutoComplete()
    End Sub

    Private Sub miupdates_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miupdates.Click
        Try
            Process.Start(Application.StartupPath & "\Config\sup.exe", Application.StartupPath & "\Config\AppFiles.xml https://raw.githubusercontent.com/samarjeet27/ynoteclassic/master/Updates.xml" + Application.ProductVersion)
        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, String.Empty, MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub micopyhtml_Click(ByVal sender As Object, ByVal e As EventArgs) Handles micopyhtml.Click
        If ActiveEditor.Tb.Html IsNot Nothing Then Clipboard.SetText(ActiveEditor.Tb.Html)
    End Sub

    Private Sub micopyrtf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles micopyrtf.Click
        If ActiveEditor IsNot Nothing Then Clipboard.SetText(ActiveEditor.Tb.Rtf)
    End Sub

    Private Sub minewscript_Click(ByVal sender As Object, ByVal e As EventArgs) Handles minewscript.Click
        CreateNewDoc()
        ActiveEditor.Text = "Script"
        ActiveEditor.Tb.Text = "using SS.Ynote.Classic;" & vbCrLf & vbCrLf & "static void Main(IYnote ynote)" & vbCrLf & "{" & vbCrLf & "// your code" & vbCrLf & "}"
        ActiveEditor.HighlightSyntax("CSharp")
        ActiveEditor.Tb.Language = "CSharp"
        ActiveEditor.Tb.DoAutoIndent()
    End Sub

    Private Sub minewsnippet_Click(ByVal sender As Object, ByVal e As EventArgs) Handles minewsnippet.Click
        CreateNewDoc()
        ActiveEditor.Text = "Snippet"
        ActiveEditor.Tb.Text = "<?xml version=""1.0""?>" & vbLf & "<YnoteSnippet>" & vbLf & "<description></description>" & vbLf & "<content><!-- content of your snippet --></content>" & vbLf & "<tabTrigger><!-- text to trigger the snippet --></tabTrigger>" & vbLf & "<scope><!-- the scope of the snippet --></scope>" & vbLf & "</YnoteSnippet>"
        ActiveEditor.HighlightSyntax("Xml")
        ActiveEditor.Tb.Language = "Xml"
        ActiveEditor.Tb.DoAutoIndent()
    End Sub

    Private Sub midocinfo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles midocinfo.Click
        If Not (TypeOf dock.ActiveDocument Is Editor) OrElse ActiveEditor Is Nothing Then

        Else
            Dim allwords = Regex.Matches(ActiveEditor.Tb.Text, "[\S]+")
            Dim selectionWords = Regex.Matches(ActiveEditor.Tb.SelectedText, "[\S]+")
            Dim message = String.Empty
            Dim startline = ActiveEditor.Tb.Selection.Start.iLine
            Dim endlines = ActiveEditor.Tb.Selection.[End].iLine
            Dim sellines = (endlines - startline) + 1

            If ActiveEditor.IsSaved Then
                Dim enc = If(EncodingDetector.DetectTextFileEncoding(ActiveEditor.Name), Encoding.[Default])
                message = String.Format("Encoding : {4}" & vbCrLf & "Words : {0}" & vbCrLf & "Selected Words : {1}" & vbCrLf & "Lines : {2}" & vbCrLf & "Selected Lines : {5}" & vbCrLf & "Column : {3}", allwords.Count, selectionWords.Count, ActiveEditor.Tb.LinesCount, ActiveEditor.Tb.Selection.Start.iChar + 1, enc.EncodingName, sellines)
            Else
                message = String.Format("Words : {0}" & vbCrLf & "Selected Words : {1}" & vbCrLf & "Lines : {2}" & vbCrLf & "Selected Lines : {4}" & vbCrLf & "Column : {3}", allwords.Count, selectionWords.Count, ActiveEditor.Tb.LinesCount, ActiveEditor.Tb.Selection.Start.iChar + 1, sellines)
            End If

            MessageBox.Show(message, "Document Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub miclose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miclose.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Close()
    End Sub

    Private Sub miprojpage_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miprojpage.Click
        Process.Start("http://ynoteclassic.codeplex.com")
    End Sub

    Private Sub mibugreport_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mibugreport.Click
        Process.Start("http://ynoteclassic.codeplex.com/workitem/list/basic")
    End Sub

    Private Sub miforum_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miforum.Click
        Process.Start("http://ynoteclassic.codeplex.com/discussions")
    End Sub

    Private Sub filemenu_Select(ByVal sender As Object, ByVal e As EventArgs) Handles filemenu.Select
        If recentfilesmenu.MenuItems.Count <> 0 Then Exit Sub
        If _mru Is Nothing Then LoadRecentList()

        For Each r In _mru
            AddRecentFile(r)
        Next
    End Sub

    Private Sub macrosmenu_Select(ByVal sender As Object, ByVal e As EventArgs) Handles macrosmenu.Select
        mimacros.MenuItems.Clear()
        Dim files As String() = IO.Directory.GetFiles(GlobalSettings.SettingsDir, "*.ynotemacro", IO.SearchOption.AllDirectories)

        For Each item In files.[Select](Function(file) New MenuItem(IO.Path.GetFileNameWithoutExtension(file), New EventHandler(AddressOf macroitem_click)) With {
            .Name = file
        })
            mimacros.MenuItems.Add(item)
        Next

        If miscripts.MenuItems.Count <> 0 Then Return
        Dim scripts As String() = IO.Directory.GetFiles(GlobalSettings.SettingsDir, "*.ys", IO.SearchOption.AllDirectories)

        For Each item In scripts.[Select](Function(file) New MenuItem(IO.Path.GetFileNameWithoutExtension(file), New EventHandler(AddressOf scriptitem_clicked)) With {
            .Name = file
        })
            miscripts.MenuItems.Add(item)
        Next
    End Sub

    Private Sub miprotectfile_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miprotectfile.Click
        If ActiveEditor Is Nothing OrElse Not ActiveEditor.IsSaved Then
            MessageBox.Show("Error : Document not saved !", Nothing, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        Using dlg = New PasswordDialog()
            Dim result = dlg.ShowDialog(Me)

            If result = DialogResult.OK Then
                Dim bytes = Encryption.Encrypt(IO.File.ReadAllBytes(ActiveEditor.Name), dlg.Password)
                If bytes IsNot Nothing Then IO.File.WriteAllBytes(ActiveEditor.Name, bytes)
            End If
        End Using

        RevertFile()
    End Sub

    Private Sub midecryptfile_Click(ByVal sender As Object, ByVal e As EventArgs) Handles midecryptfile.Click
        If ActiveEditor Is Nothing OrElse Not ActiveEditor.IsSaved Then
            MessageBox.Show("Error : Document not saved !", Nothing, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        Using dlg = New PasswordDialog()
            Dim result = dlg.ShowDialog(Me)

            If result = DialogResult.OK Then
                Dim bytes = Encryption.Decrypt(IO.File.ReadAllBytes(ActiveEditor.Name), dlg.Password)
                If bytes IsNot Nothing Then IO.File.WriteAllBytes(ActiveEditor.Name, bytes)
            End If
        End Using

        RevertFile()
    End Sub

    Private Sub minewsyntax_Click(ByVal sender As Object, ByVal e As EventArgs) Handles minewsyntax.Click
        CreateNewDoc()
        ActiveEditor.Text = "SyntaxFile"
        ActiveEditor.Tb.Text = "<?xml version=""1.0""?>" & vbCrLf & vbTab & "<YnoteSyntax>" & vbCrLf & vbTab & vbTab & "<Syntax CommentPrefix="""" Extensions=""""/>" & vbCrLf & vbTab & vbTab & "<Brackets Left="""" Right=""""/>" & vbCrLf & vbTab & vbTab & "<Rule Type="""" Regex=""""/>" & vbCrLf & vbTab & vbTab & "<Folding Start="""" End=""""/>" & vbCrLf & vbTab & "</YnoteSyntax>"
        ActiveEditor.HighlightSyntax("Xml")
        ActiveEditor.Tb.Language = "Xml"
    End Sub

    Private Sub miinscliphis_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miinscliphis.Click
        Dim lst = New List(Of AutocompleteItem)()

        For Each doc As Editor In dock.Documents.OfType(Of Editor)()

            For Each historyitem In doc.Tb.ClipboardHistory
                lst.Add(New AutocompleteItem(historyitem))
            Next
        Next

        Using cmw = New CommandWindow(lst)
            AddHandler cmw.ProcessCommand, Function(o, args)
                                               ActiveEditor.Tb.InsertText(args.Text)
                                           End Function
            cmw.ShowDialog(Me)
        End Using
    End Sub

    Private Sub misplitbelow_Click(ByVal sender As Object, ByVal e As EventArgs) Handles misplitbelow.Click
        Dim splitedit = New Editor With {
            .Name = ActiveEditor.Name,
            .Text = "[Split] " & ActiveEditor.Text
        }
        splitedit.Tb.SourceTextBox = ActiveEditor.Tb
        splitedit.Tb.[ReadOnly] = True
        splitedit.Show(ActiveEditor.Pane, DockAlignment.Bottom, 0.5)
    End Sub

    Private Sub misplitbeside_Click(ByVal sender As Object, ByVal e As EventArgs) Handles misplitbeside.Click
        Dim splitedit = New Editor With {
            .Name = ActiveEditor.Name,
            .Text = "[Split] " & ActiveEditor.Text
        }
        splitedit.Tb.SourceTextBox = ActiveEditor.Tb
        splitedit.Tb.[ReadOnly] = True
        splitedit.Show(ActiveEditor.Pane, DockAlignment.Right, 0.5)
    End Sub

    Private Sub menuItem4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles menuItem4.Click
        Dim splitedit = New Editor With {
            .Name = ActiveEditor.Name,
            .Text = "[Split] " & ActiveEditor.Text
        }
        splitedit.Tb.SourceTextBox = ActiveEditor.Tb
        splitedit.Tb.[ReadOnly] = True
        AddHandler ActiveEditor.Tb.VisibleRangeChangedDelayed, Function(o, args)
                                                                   UpdateScroll(splitedit.Tb, ActiveEditor.Tb.VerticalScroll.Value, ActiveEditor.Tb.Selection.Start.iLine)
                                                               End Function
        splitedit.Show(ActiveEditor.Pane, DockAlignment.Right, 0.5)
    End Sub

    Private Sub UpdateScroll(ByVal tb As FastColoredTextBox, ByVal vPos As Integer, ByVal curLine As Integer)
        If vPos <= tb.VerticalScroll.Maximum Then
            tb.VerticalScroll.Value = vPos
            tb.UpdateScrollbars()
        End If

        If curLine < tb.LinesCount Then tb.Selection = New Range(tb, 0, curLine, 0, curLine)
    End Sub

    Private Sub mimap_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mimap.Click
        mimap.Checked = Not mimap.Checked

        For Each content As Editor In dock.Documents.OfType(Of Editor)()
            content.ShowDocumentMap = mimap.Checked
        Next
    End Sub

    Private Sub viewmenu_Select(ByVal sender As Object, ByVal e As EventArgs) Handles viewmenu.Select
        If ActiveEditor IsNot Nothing Then
            mimap.Checked = Globals.Settings.ShowDocumentMap
            midistractionfree.Checked = ActiveEditor.DistractionFree
            wordwrapmenu.Checked = Globals.Settings.WordWrap
            mimenu.Checked = Globals.Settings.ShowMenuBar
            mihiddenchars.Checked = Globals.Settings.HiddenChars
        End If
    End Sub

    Private Sub distractionfree_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles distractionfree.Click
        If Panel.ActiveDocument Is Nothing OrElse Not (TypeOf Panel.ActiveDocument Is Editor) Then Exit Sub

        If ActiveEditor.DistractionFree Then
            FormBorderStyle = FormBorderStyle.Sizable
            WindowState = FormWindowState.Normal
        Else
            FormBorderStyle = FormBorderStyle.None
            WindowState = FormWindowState.Maximized
        End If

        For Each edit As DockContent In dock.Contents
            If TypeOf edit Is Editor Then
                TryCast(edit, Editor).ToggleDistrationFreeMode()
            End If
        Next

        Globals.DistractionFree = ActiveEditor.DistractionFree
    End Sub

    Private Sub commentmenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles commentmenu.Click
        ActiveEditor.Tb.CommentSelected()
    End Sub

    Private Sub mimenu_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mimenu.Click
        mimenu.Checked = Not mimenu.Checked
        ToggleMenu(mimenu.Checked)
        Globals.Settings.ShowMenuBar = mimenu.Checked
    End Sub

    Private Sub migotosymbol_Click(ByVal sender As Object, ByVal e As EventArgs) Handles migotosymbol.Click
        If ActiveEditor Is Nothing Then Return
        Dim symbols = SymbolList.GetPlaces(ActiveEditor)
        Dim cwin = New CommandWindow(symbols)
        AddHandler cwin.ProcessCommand, AddressOf cwin_ProcessCommand
        cwin.Tag = symbols
        cwin.ShowDialog(Me)
    End Sub

    Private Sub cwin_ProcessCommand(ByVal sender As Object, ByVal e As CommandWindowEventArgs) 'Handles cwin.ProcessCommand
        For Each item In TryCast((TryCast(sender, CommandWindow)).Tag, IEnumerable(Of AutocompleteItem))

            If item.Text = e.Text Then
                ActiveEditor.Tb.SelectionStart = CInt((item.Tag))
                ActiveEditor.Tb.DoSelectionVisible()
            End If
        Next
    End Sub

    Private Sub miclearbookmarks_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miclearbookmarks.Click
        For Each bookmark In ActiveEditor.Tb.Bookmarks.ToArray()
            ActiveEditor.Tb.UnbookmarkLine(bookmark.LineIndex)
        Next
    End Sub

    Private Sub mireindentline_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mireindentline.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.DoAutoIndent(ActiveEditor.Tb.Selection.Start.iLine)
    End Sub

    Private Sub miselectedfile_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miselectedfile.Click
        Try

            If ActiveEditor.IsSaved Then
                Dim dir = IO.Path.GetDirectoryName(ActiveEditor.Name)
                Dim filename As String = ActiveEditor.Tb.SelectedText

                If IO.Path.IsPathRooted(filename) Then
                    OpenFile(ActiveEditor.Tb.SelectedText)
                Else

                    For Each file In IO.Directory.GetFiles(dir)
                        If IO.Path.GetFileName(file) = filename OrElse IO.Path.GetFileNameWithoutExtension(file) = filename Then OpenFile(file)
                    Next
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, Nothing, MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub migotofileinproject_Click(ByVal sender As Object, ByVal e As EventArgs) Handles migotofileinproject.Click
        If Globals.ActiveProject IsNot Nothing Then
            path = Globals.ActiveProject.Path
        Else

            If ActiveEditor.IsSaved Then
                path = IO.Path.GetDirectoryName(ActiveEditor.Name)
            Else
                MessageBox.Show("Please Open a Project First!")
                Exit Sub
            End If
        End If

        Dim autocompletelist = New List(Of AutocompleteItem)()
        Dim files As String() = IO.Directory.GetFiles(path, "*.*", IO.SearchOption.TopDirectoryOnly)

        For Each file In files
            autocompletelist.Add(New FuzzyAutoCompleteItem(IO.Path.GetFileName(file)))
        Next

        Dim commandwindow = New CommandWindow(autocompletelist)
        AddHandler commandwindow.ProcessCommand, AddressOf commandwindow_ProcessCommand
        commandwindow.ShowDialog(Me)
    End Sub

    Private Sub commandwindow_ProcessCommand(ByVal sender As Object, ByVal e As CommandWindowEventArgs) 'Handles commandwindow.ProcessCommand
        Dim files = IO.Directory.GetFiles(path, "*.*", IO.SearchOption.AllDirectories)

        For Each file In files

            If IO.Path.GetFileName(file) = e.Text Then
                OpenFile(file)
                Return
            End If
        Next
    End Sub

    Private Sub miopenproject_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miopenproject.Click
        Using dlg = New OpenFileDialog()
            dlg.Filter = "Ynote Project Files (*.ynoteproject)|*.ynoteproject|All Files (*.*)|*.*"

            If dlg.ShowDialog() = DialogResult.OK Then
                OpenProject(YnoteProject.Load(dlg.FileName))
                If Not _projs.Contains(dlg.FileName) Then _projs.Add(dlg.FileName)
            End If
        End Using
    End Sub

    Private Sub OpenProject(ByVal project As YnoteProject)
        If Globals.ActiveProject Is project OrElse project Is Nothing Then Exit Sub
        If projectPanel Is Nothing Then projectPanel = New ProjectPanel()

        If IO.File.Exists(project.LayoutFile) Then

            If dock.Contents.Count <> 0 Then
                Dim docs = dock.Contents.ToArray()

                For i As Integer = 0 To docs.Length - 1
                    Dim document = docs(i)
                    document.DockHandler.Close()
                Next
            End If

            ' dock.LoadFromXml(project.LayoutFile, GetContentFromPersistString)
        Else
            projectPanel.OpenProject(project)
            projectPanel.Show(dock, DockState.DockLeft)
        End If

        Text = project.Name & " - Ynote Classic"
    End Sub

    Private Sub micloseproj_Click(ByVal sender As Object, ByVal e As EventArgs) Handles micloseproj.Click
        Dim proj As ProjectPanel = Nothing

        For Each window In dock.Contents
            If TypeOf window Is ProjectPanel Then proj = TryCast(window, ProjectPanel)
        Next

        If proj IsNot Nothing Then
            proj.CloseProject()
            Globals.ActiveProject = Nothing
            Text = "Ynote Classic"
        End If
    End Sub

    Private Sub mieditproj_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mieditproj.Click
        If Globals.ActiveProject Is Nothing OrElse String.IsNullOrEmpty(Globals.ActiveProject.FilePath) Then Return
        OpenFile(Globals.ActiveProject.FilePath)
    End Sub

    Private Sub misaveproj_Click(ByVal sender As Object, ByVal e As EventArgs) Handles misaveproj.Click
        Dim proj = Globals.ActiveProject
        If proj Is Nothing Then Return

        If Not proj.IsSaved Then

            Using dlg = New SaveFileDialog()
                dlg.Filter = "Ynote Project Files (*.ynoteproject)|*.ynoteproject"
                Dim result = dlg.ShowDialog()

                If result = DialogResult.OK Then
                    proj.Name = IO.Path.GetFileNameWithoutExtension(dlg.FileName)
                    proj.Save(dlg.FileName)
                    proj.FilePath = dlg.FileName

                    For Each content As DockContent In dock.Contents
                        If TypeOf content Is ProjectPanel Then TryCast(content, ProjectPanel).OpenProject(proj)
                    Next

                    If Not _projs.Contains(proj.FilePath) Then _projs.Add(proj.FilePath)
                End If
            End Using
        End If
    End Sub

    Private Sub miaddtoproj_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miaddtoproj.Click
        Using browser = New FolderBrowserDialogEx()
            browser.ShowEditBox = True
            browser.ShowFullPathInEditBox = True

            If browser.ShowDialog() = DialogResult.OK Then
                Dim proj = New YnoteProject()
                proj.Path = browser.SelectedPath
                OpenProject(proj)
            End If
        End Using
    End Sub

    Private Sub miproject_Select(ByVal sender As Object, ByVal e As EventArgs) Handles miproject.Select
        If _projs Is Nothing Then LoadRecentProjects()

        If miopenrecent.MenuItems.Count = 0 Then

            For Each item In _projs
                miopenrecent.MenuItems.Add(New MenuItem(IO.Path.GetFileNameWithoutExtension(item), AddressOf openrecentproj_Click) With {
                    .Name = item
                })
            Next
        End If
    End Sub

    Private Sub openrecentproj_Click(ByVal sender As Object, ByVal e As EventArgs) 'Handles openrecentproj.Click
        Dim menu = TryCast(sender, MenuItem)
        OpenProject(YnoteProject.Load(menu.Name))
    End Sub

    Private Sub miswitchproj_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miswitchproj.Click
        If miopenrecent.MenuItems.Count = 0 Then miproject_Select(Nothing, Nothing)
        Dim completemenu = New List(Of AutocompleteItem)()

        For Each item As MenuItem In miopenrecent.MenuItems
            completemenu.Add(New FuzzyAutoCompleteItem(item.Text))
        Next

        Dim cwindow = New CommandWindow(completemenu)
        AddHandler cwindow.ProcessCommand, AddressOf cwindow_ProcessCommand
        cwindow.ShowDialog(Me)
    End Sub

    Private Sub cwindow_ProcessCommand(ByVal sender As Object, ByVal e As CommandWindowEventArgs) 'Handles cwindow.ProcessCommand
        For Each item As MenuItem In miopenrecent.MenuItems

            If item.Text = e.Text Then
                Dim proj = YnoteProject.Load(item.Name)
                OpenProject(proj)
            End If
        Next
    End Sub

    Private Sub mirefreshproj_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mirefreshproj.Click
        If Globals.ActiveProject IsNot Nothing Then OpenProject(Globals.ActiveProject)
    End Sub

    Private Sub minewview_Click(ByVal sender As Object, ByVal e As EventArgs) Handles minewview.Click
        If ActiveEditor.IsSaved Then
            Dim edit = New Editor()
            edit.Text = ActiveEditor.Text
            edit.Name = ActiveEditor.Name
            edit.Tb.SourceTextBox = ActiveEditor.Tb
            edit.Show(dock)
        Else
            MessageBox.Show("File Not Saved!")
        End If
    End Sub

    Private Sub mishellcmd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mishellcmd.Click
        AskInput("Shell:", New InputWindow.GotInputEventHandler(Sub(o, args)
                                                                    ExecShellCommand(args.InputValue)
                                                                End Sub))
    End Sub

    Private Sub ExecShellCommand(ByVal command As String)
        Dim splits = command.Split(":"c)
        If splits(0) <> "Shell" Then Return
        command = splits(1)
        Dim output As BuildOutput
        Dim contents = dock.Contents.OfType(Of BuildOutput)()

        If contents.Any() Then
            output = contents.First()
        Else
            output = New BuildOutput()
            output.Show(dock, DockState.DockBottom)
        End If

        Dim info = New ProcessStartInfo("cmd.exe", "/k " & command)
        info.CreateNoWindow = True
        info.WindowStyle = ProcessWindowStyle.Hidden
        info.RedirectStandardOutput = True
        info.RedirectStandardError = True
        info.UseShellExecute = False

        If Globals.ActiveProject Is Nothing Then
            If ActiveEditor IsNot Nothing AndAlso ActiveEditor.IsSaved Then info.WorkingDirectory = IO.Path.GetDirectoryName(ActiveEditor.Name)
        Else
            info.WorkingDirectory = Globals.ActiveProject.Path
        End If

        Dim proc = Process.Start(info)
        Dim out_msg As String
        Dim error_msg As String

        Using reader = proc.StandardOutput
            out_msg = reader.ReadToEnd()
        End Using

        Using erReader = proc.StandardError
            error_msg = erReader.ReadToEnd()
        End Using

        If out_msg <> String.Empty Then
            output.AddOutput(out_msg)
        Else
            output.AddOutput(error_msg)
        End If
    End Sub

    Private Sub migoogle_Click(ByVal sender As Object, ByVal e As EventArgs) Handles migoogle.Click
        AskInput("Google:", Function(o, args) Process.Start("http://www.google.com/search?q=" & args.GetFormattedInput()))
    End Sub

    Private Sub miwikipedia_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miwikipedia.Click
        AskInput("Wikipedia:", Function(o, args) Process.Start("http://en.wikipedia.org/w/index.php?search=" & args.GetFormattedInput()))
    End Sub

    Private Sub miselectfindNext_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miselectfindNext.Click
        If ActiveEditor IsNot Nothing Then ActiveEditor.Tb.SelectAndFindNext()
    End Sub

    Private Sub miscriptconsole_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miscriptconsole.Click
        AskInput("Function:", AddressOf EvaluateFunction)
    End Sub

    Private Sub EvaluateFunction(ByVal sender As Object, ByVal args As InputEventArgs)
        Dim s As String = args.InputValue.Substring(9, args.InputValue.Length - "Function:".Length)
        If Not s.Contains(";") Then s += ";"
        Dim code As String = String.Format("//css_include {0}\Installed\YnoteCommons.ysr
        using System.Windows.Forms;
        using System.IO;
        
        public static void Main(IYnote yn){{
            var ynote = new YnoteCommons(yn);{1}}}", GlobalSettings.SettingsDir, s)
        YnoteScript.RunCode(code)
    End Sub

    Private Sub miuserkeys_Click(ByVal sender As Object, ByVal e As EventArgs) Handles miuserkeys.Click
        Dim file As String = IO.Path.Combine(GlobalSettings.SettingsDir, "User.ynotekeys")
        Dim info As IO.FileInfo = New IO.FileInfo(file)

        If info.Exists Then
            OpenFileAsync(file)
        Else
            Dim result = MessageBox.Show("User.ynotekeys doesn't exist." & vbLf & " Create it ?", "Ynote Clasic", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                IO.File.Create(file).Close()
                OpenFileAsync(file)
            End If
        End If
    End Sub

    Private Sub mifindprev_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mifindprev.Click
        If ActiveEditor Is Nothing Then Exit Sub
        ActiveEditor.Tb.findForm.FindPrevious(ActiveEditor.Tb.findForm.Text)
    End Sub

    Private Sub milasttab_Click(ByVal sender As Object, ByVal e As EventArgs) Handles milasttab.Click
        dock.Contents(dock.Contents.Count - 1).DockHandler.Show(dock, DockState.Document)
    End Sub

#End Region

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

End Class
