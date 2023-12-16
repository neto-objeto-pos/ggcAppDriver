'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Table Look up 
'
' Copyright 2012 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 04/16/2013 11:43 am ]
'      Start translating of this object.
'      This object is based from the Lookup originally written in vb6.
'  XerSys [ 05/29/2013 08:50 am ]
'       Continue creating this object.
'  XerSys [ 06/13/2013 01:36 pm ]
'       Add auto load feature for automatic returning record if only 1 passed the criteria
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports System.Environment
Imports MySql.Data.MySqlClient
Imports ADODB
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Threading
Imports System.Timers
Imports System

Public Class Lookup
    Private p_oAppDriver As GRider
    Private p_oDT As DataTable
    Private p_oDR As DataRow
    Private p_oDA As MySqlDataAdapter

    Private p_sSQLSource As String
    Private p_sCondition As String

    Private p_sColName As String
    Private p_sColTitle As String
    Private p_sColFormat As String
    Private p_sFieldName As String
    Private p_nDefCol As Integer

    Private p_oFieldName() As String
    Private p_oColName() As String
    Private p_oColFormat() As String
    Private p_oColAlign() As DataGridViewContentAlignment
    Private p_oRow As Object
    Private p_nIndex As Integer = 1

    Private p_bDisplay As Boolean = False
    Private p_bListed As Boolean = False
    Private p_bLoading As Boolean = False
    Private p_bCancel As Boolean = True
    Private p_bSearch As Boolean = False
    
    Private p_bCancelAsync As Boolean = False
    Private p_bClosing As Boolean = False
    Private p_bAutoLoad As Boolean = False
    Private p_bShowCondition As Boolean = True

    Private p_nStart As Integer = 0
    Private p_nPageRow As Integer = 10

    Public Delegate Function fillListCallback() As Integer

    Private Declare Function keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, _
                              ByVal dwFlags As Int32, ByVal dwExtraInfo As Int32) As Boolean

    Private Declare Function GetLastInputInfo Lib "user32.dll" (ByRef plii As LASTINPUTINFO) As Boolean

    Private Structure LASTINPUTINFO
        Public cbSize As Int32
        Public dwTime As Int32
    End Structure

    Public Property AppDriver As GRider
        Get
            Return p_oAppDriver
        End Get
        Set(ByVal value As GRider)
            p_oAppDriver = value
        End Set
    End Property

    Public Property QuerySource As String
        Get
            Return p_sSQLSource
        End Get
        Set(ByVal value As String)
            p_sSQLSource = value
        End Set
    End Property

    Public Property FieldName As String
        Get
            Return p_sFieldName
        End Get
        Set(ByVal value As String)
            p_sFieldName = value
        End Set
    End Property

    Public Property ColumnDesc As String
        Get
            Return p_sColName
        End Get
        Set(ByVal value As String)
            p_sColName = value
        End Set
    End Property

    Public Property ColHeader As String
        Get
            Return p_sColTitle
        End Get
        Set(ByVal value As String)
            p_sColTitle = value
        End Set
    End Property

    Public Property ColFormat As String
        Get
            Return p_sColFormat
        End Get
        Set(ByVal value As String)
            p_sColFormat = value
        End Set
    End Property

    Public Property Condition As String
        Get
            Return p_sCondition
        End Get
        Set(ByVal value As String)
            p_sCondition = value
        End Set
    End Property

    Public Property DefaultField As Integer
        Get
            Return p_nDefCol
        End Get
        Set(ByVal value As Integer)
            p_nDefCol = value
        End Set
    End Property

    Public ReadOnly Property SelectedRow As DataRow
        Get
            Return getSelectedRow()
        End Get
    End Property

    Public Property AutoLoad As Boolean
        Get
            Return p_bAutoLoad
        End Get
        Set(ByVal value As Boolean)
            p_bAutoLoad = value
        End Set
    End Property

    Public Property ShowCondition As Boolean
        Get
            Return p_bShowCondition
        End Get
        Set(ByVal value As Boolean)
            p_bShowCondition = value
        End Set
    End Property

    Public ReadOnly Property Cancelled As Boolean
        Get
            Return p_bCancel
        End Get
    End Property

    Private Sub setDefaults()
        Dim lnCtr As Integer
        Dim lnCol As Integer
        Dim loTemp() As String

        ' initialize sql adapter
        p_oDA = New MySqlDataAdapter(AddCondition(p_sSQLSource, "0 = 1"), p_oAppDriver.Connection)
        Debug.Print("setDefaults")

        ' initialize fields
        p_oDT = New DataTable
        Try
            p_oDA.Fill(p_nStart, 1, p_oDT)
        Catch ex2 As MySqlException
            MsgBox(ex2.Message)
        Catch ex1 As Exception
            ' this was just a trial by error solution because everytime i re-run the lookup
            '   an error occured stating that the datareader is null. But if i continue 
            '   executing the code, it runs. Hope it won't create havok in memory handling 
            '   of vb.net if it will be used by other client.
            p_oDA.Fill(p_nStart, 1, p_oDT)
            'MsgBox(ex1.Message)
        End Try

        ' bases of grid for column intended to be displayed
        If p_sColName = "" Then
            lnCol = p_oDT.Columns.Count - 1
            p_oColName = getFieldName(lnCol)
        Else
            p_oColName = Split(p_sColName, "»")
            lnCol = UBound(p_oColName)
        End If

        If p_sFieldName = "" Then
            p_oFieldName = p_oColName
        Else
            p_oFieldName = Split(p_sFieldName, "»")
        End If

        If p_sColTitle = "" Then
            loTemp = p_oColName
        Else
            loTemp = Split(p_sColTitle, "»")
        End If

        If p_sColFormat = "" Then
            p_oColFormat = getColFormat(lnCol)
        Else
            p_oColFormat = Split(p_sColFormat, "»")
        End If

        p_oColAlign = getColAlign(lnCol)

        ' fill the header
        DataGridView1.RowHeadersVisible = False
        ' format datagridview
        With DataGridView1
            .Anchor = Windows.Forms.AnchorStyles.Top Or Windows.Forms.AnchorStyles.Bottom Or Windows.Forms.AnchorStyles.Left Or Windows.Forms.AnchorStyles.Right

            For lnCtr = 0 To lnCol
                .Columns.Add(p_oFieldName(lnCtr), loTemp(lnCtr))
                .Columns(lnCtr).AutoSizeMode = Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
                .Columns(lnCtr).DefaultCellStyle.Alignment = p_oColAlign(lnCtr)
                ComboBox1.Items.Add(loTemp(lnCtr))
            Next
        End With

        ' set default search column
        'kalyptus (2013-06-31)
        'originally (If lnCol < 0 then)
        If lnCol < 1 Then
            p_nDefCol = 0
        Else
            If p_nDefCol > lnCol Then
                p_nDefCol = 1
            End If
        End If

        p_nIndex = p_nDefCol
        ComboBox1.SelectedIndex = p_nIndex

        ' check if condition must be displayed
        If p_bShowCondition Then txtSearch.Text = p_sCondition
    End Sub

    Private Function isParamOK() As Boolean
        If p_sSQLSource = "" Then
            Return False
        End If

        If p_oAppDriver Is Nothing Then
            Return False
        End If
        Return True
    End Function

    Private Function getFieldName(ByVal fnIndex As Integer) As Object
        Dim lasFieldName() As String
        Dim lnCtr As Integer

        ReDim lasFieldName(p_oDT.Columns.Count - 1)
        For lnCtr = 0 To UBound(lasFieldName)
            lasFieldName(lnCtr) = p_oDT.Columns(lnCtr).ColumnName
        Next
        Return lasFieldName
    End Function

    Private Function getColFormat(ByVal fnIndex As Integer) As Object
        Dim lasColFormat(fnIndex) As String
        Dim lnCtr As Integer

        For lnCtr = 0 To fnIndex
            Debug.Print(p_oDT.Columns(lnCtr).DataType.ToString)
            lasColFormat(lnCtr) = getDefaultFormat(p_oDT.Columns(lnCtr).DataType.ToString)
        Next
        Return lasColFormat
    End Function

    Private Function getColAlign(ByVal fnIndex As Integer) As Object
        Dim lasColAlign(fnIndex) As DataGridViewContentAlignment
        Dim lnCtr As Integer

        For lnCtr = 0 To fnIndex
            ' single object formating 
            Select Case p_oDT.Columns(lnCtr).DataType.ToString().ToLower
                Case "system.datetime"
                    lasColAlign(lnCtr) = DataGridViewContentAlignment.MiddleLeft
                Case "system.decimal", "system.double"
                    lasColAlign(lnCtr) = DataGridViewContentAlignment.MiddleRight
                Case "system.byte", "system.int16", "system.int32", "system.int64", "system.single"
                    lasColAlign(lnCtr) = DataGridViewContentAlignment.MiddleRight
                Case Else
                    lasColAlign(lnCtr) = DataGridViewContentAlignment.MiddleLeft
            End Select
        Next
        Return lasColAlign
    End Function

    Private Function setSQLAdapter() As Boolean
        Dim lsSQL As String

        'kalyptus - 2013.07.03
        'add testing of p_bshowcondition. Load p_sCondition if false
        If p_bShowCondition Then
            lsSQL = AddCondition(p_sSQLSource, p_oFieldName(p_nIndex) & " LIKE " & strParm(Trim(txtSearch.Text) & "%"))
        Else
            lsSQL = AddCondition(p_sSQLSource, p_sCondition)
        End If

        lsSQL = setOrderBy(lsSQL)
        'p_oDA.SelectCommand.CommandText = lsSQL 
        p_oDA.SelectCommand = New MySqlCommand(lsSQL, p_oAppDriver.Connection)
        Return True
    End Function

    Private Function setOrderBy(ByVal fsSQL As String) As String
        Dim lsSQL As String
        Dim lsOrder As String
        Dim lnCtr As Integer
        Dim lxOrder As String = " ORDER BY "

        lsOrder = ""
        lsSQL = fsSQL

        lnCtr = InStr(fsSQL, lxOrder, CompareMethod.Text)
        If lnCtr > 0 Then
            lsSQL = fsSQL.Substring(0, lnCtr)
            lsOrder = Trim(fsSQL.Substring(lnCtr + lxOrder.Length - 1))
        End If

        If lsOrder = "" Then
            lsOrder = p_oFieldName(p_nIndex)
        Else
            lsOrder = p_oFieldName(p_nIndex) & ", " & lsOrder
        End If
        Return lsSQL & lxOrder & lsOrder
    End Function

    Private Sub clearGrid()
        DataGridView1.Rows.Clear()
        p_bListed = False
    End Sub

    Private Sub loadInitList()
        ' every time the txt Search changes values, retrieve on page of data for initial display
        p_nStart = 0
        p_oDT.Clear()
        Call clearGrid()

        Debug.Print(Now)
        p_bListed = fillList() > 0
        Debug.Print(Now)
    End Sub

    Private Function fillList() As Integer
        Dim loDTTemp As DataTable
        Dim loDRow As DataRow
        Dim lnCtr As Integer
        Dim lnRow As Integer
        Dim lnCol As Integer

        ' check if closing
        If p_bClosing Or p_bCancelAsync Then Return 0
        ' fetch record
        If Not setSQLAdapter() Then
            MsgBox("Unable to set Query Statement!" & vbCrLf & _
                "Please inform SEG/SSG about this mattter!", vbCritical, "Warning")
            Return 0
        End If

        '' verify if pending cancel for backgroundworker exists
        'Do While BackgroundWorker1.CancellationPending
        '    BackgroundWorker1.CancelAsync()
        '    Debug.Print(Now)
        '    Thread.Sleep(5)
        'Loop

        loDTTemp = New DataTable
        Debug.Print(p_oDA.SelectCommand.CommandText)

        Try
            If p_bClosing Or p_bCancelAsync Then Return 0
            p_oDA.Fill(p_nStart, p_nPageRow, loDTTemp)
        Catch ex As Exception
            Return 0
        End Try

        ' check if client wants to auto load lookup result
        If p_bAutoLoad Then
            If loDTTemp.Rows.Count = 1 Then
                p_oDR = loDTTemp.Rows(0)

                p_bCancel = False
                p_bClosing = True
                Me.Hide()
            End If
        End If

        Dim loRow As DataGridViewRow

        If Me.DataGridView1.InvokeRequired Then
            If p_bClosing Or p_bCancelAsync Then Return 0

            Dim loTemp As New fillListCallback(AddressOf fillList)

            Try
                lnRow = Me.Invoke(loTemp)
            Catch ex As Exception
                Return 0
            End Try
            Return lnRow
        End If

        DataGridView1.AllowUserToAddRows = True
        For lnCtr = 0 To loDTTemp.Rows.Count - 1
            loDRow = p_oDT.NewRow
            loDRow = loDTTemp.Rows(lnCtr)
            p_oDT.ImportRow(loDRow)

            'Debug.Print(loDRow(1).ToString & " " & lnRow)
            lnRow = DataGridView1.Rows.Count - 1
            loRow = DataGridView1.Rows(lnRow).Clone
            For lnCol = 0 To UBound(p_oColName)
                If p_oColFormat(lnCol) = "" Then ' string
                    loRow.Cells(lnCol).Value = String.Format(p_oColFormat(lnCol), loDRow.Item(p_oColName(lnCol)))
                Else
                    Debug.Print(p_oColFormat(lnCol) & " " & String.Format(p_oColFormat(lnCol), loDRow.Item(p_oColName(lnCol))))
                    loRow.Cells(lnCol).Value = String.Format(p_oColFormat(lnCol), loDRow.Item(p_oColName(lnCol)))
                End If
            Next

            DataGridView1.Rows.Add(loRow)
        Next

        'kalyptus - 2013.07.03
        'reset condition
        'set value to txtSearch
        'If Not p_bShowCondition Then
        '   If loDTTemp.Rows.Count > 0 Then
        '        txtSearch.Text = loDTTemp.Rows(0).Item(p_nIndex)
        '        p_bShowCondition = True
        '    End If
        'End If

        ' remove the blank row at the end of the grid
        DataGridView1.AllowUserToAddRows = False
        Return loDTTemp.Rows.Count
    End Function

    'Private Sub addRow2Grid(ByVal foDRow As DataRow)
    '    ' check if query was cancelled
    '    If p_bCancelAsync Then
    '        Exit Sub
    '    End If

    '    ' if the thread accessing the control is different from the thread that created it,
    '    '   it must be asyncrhorously proccess using invoke method
    '    If Me.DataGridView1.InvokeRequired Then
    '        'Dim loTemp As New addRow2GridCallback(AddressOf addRow2Grid)
    '        'Me.Invoke(loTemp, New Object() {foDRow})
    '    Else
    '        DataGridView1.AllowUserToAddRows = True

    '        Dim lnRow As Integer = DataGridView1.Rows.Count - 1
    '        Dim lnCol As Integer
    '        Dim loRow As DataGridViewRow = DataGridView1.Rows(lnRow).Clone

    '        Debug.Print(foDRow(1).ToString & " " & lnRow)
    '        For lnCol = 0 To UBound(p_oColName)
    '            loRow.Cells(lnCol).Value = foDRow.Item(lnCol).ToString()
    '        Next

    '        DataGridView1.Rows.Add(loRow)
    '        ' remove the blank row at the end of the grid
    '        DataGridView1.AllowUserToAddRows = False
    '    End If
    'End Sub

    'Private Function colName2Array(ByVal foValue As DataRow)
    '    Dim lnCol As Integer = foValue.Table.Columns.Count - 1
    '    Dim loColName(lnCol) As String

    '    For lnCol = 0 To UBound(loColName)
    '        loColName(lnCol) = foValue.Table.Columns(lnCol).ColumnName
    '    Next
    '    Return loColName
    'End Function

    'Private Function colFormat2Array(ByVal foValue As DataRow) As Object
    '    Dim lnCol As Integer = foValue.Table.Columns.Count - 1
    '    Dim loColFormat(lnCol) As String

    '    For lnCol = 0 To UBound(loColFormat)
    '        loColFormat(lnCol) = getDefaultFormat(foValue.Table.Columns(lnCol).DataType.ToString)
    '    Next
    '    Return loColFormat
    'End Function

    Private Sub sortList()
        DataGridView1.Sort(DataGridView1.Columns(p_nIndex), ListSortDirection.Ascending)
    End Sub

    Private Function getDefaultFormat(ByVal fsType As String) As String
        Dim lsDefault As String

        ' single object formating 
        Select Case fsType.ToLower
            Case "system.datetime"
                lsDefault = xsDATE_LONG
            Case "system.decimal", "system.double"
                lsDefault = xsDECIMAL
            Case "system.byte", "system.int16", "system.int32", "system.int64", "system.single"
                lsDefault = xsINTEGER
            Case Else
                lsDefault = "G"
        End Select
        Return "{0:" & lsDefault & "}"
    End Function

    Private Function getSelectedRow() As DataRow
        If p_bCancel Then Return Nothing

        If p_bLoading Then
            Call cancelQuery()
        End If

        Dim lnRow As Integer

        For lnRow = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(lnRow).Selected Then
                Return p_oDT.Rows(lnRow)
            End If
        Next
        Return Nothing
    End Function

    Private Sub cancelQuery()
        If p_bLoading Then
            BackgroundWorker1.CancelAsync()
            p_oDA.SelectCommand.Cancel()

            p_bCancelAsync = True
            Thread.Sleep(50)
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        p_nIndex = ComboBox1.SelectedIndex
        Call sortList()
        txtSearch.Text = ""
    End Sub

    Private Sub txtSearch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearch.GotFocus
        p_bSearch = True
    End Sub

    Private Sub txtSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyDown
        Select Case e.KeyCode
            Case Windows.Forms.Keys.Down, Windows.Forms.Keys.Up, Windows.Forms.Keys.PageDown, Windows.Forms.Keys.PageUp
                DataGridView1.Focus()
                ' now move the grid 
                keybd_event(e.KeyCode, 0, 0, 0)
                keybd_event(e.KeyCode, 0, 2, 0) ' release keystroke
            Case Windows.Forms.Keys.Return
        End Select
    End Sub

    Private Sub Lookup_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If Not p_bListed Then Call loadInitList()
    End Sub

    Private Sub txtSearch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearch.LostFocus
        p_bSearch = False
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim lnRow As Integer

        p_bLoading = True
        lnRow = DataGridView1.Rows.Count

        If lnRow < p_nPageRow Then
            ' Initially retrieve record is less than maximum visible row
            Exit Sub
        End If

        p_nStart += lnRow 
        Do Until p_bCancelAsync
            lnRow = fillList()
            Debug.Print(lnRow)
            If lnRow < p_nPageRow Then
                Exit Do
            End If
            p_nStart += p_nPageRow

            BackgroundWorker1.ReportProgress(p_nStart / p_nPageRow * 100)
            'Thread.Sleep(50)
        Loop

        If p_bCancelAsync Then
            p_bCancelAsync = False
        End If

        p_bDisplay = True
        p_bListed = True
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        p_bLoading = False
    End Sub

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        BackgroundWorker1.WorkerReportsProgress = True
        BackgroundWorker1.WorkerSupportsCancellation = True
        Call PreventFlicker()
    End Sub

    'reduces flickering on controls
    Private Sub PreventFlicker()
        With Me
            .SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            .SetStyle(ControlStyles.UserPaint, True)
            .SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            .UpdateStyles()
        End With
    End Sub

    Public Sub New(ByRef foAppDriver As GRider _
                   , ByVal fsSource As String _
                   , ByVal fsSearch As String _
                   , Optional ByVal fsColName As String = "" _
                   , Optional ByVal fsColTitle As String = "" _
                   , Optional ByVal fsColFormat As String = "" _
                   , Optional ByVal fsFieldName As String = "" _
                   , Optional ByVal fnDefCol As Integer = 1 _
                   , Optional ByVal fbAutoLoad As Boolean = False)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        BackgroundWorker1.WorkerReportsProgress = True
        BackgroundWorker1.WorkerSupportsCancellation = True

        p_oAppDriver = foAppDriver
        p_sSQLSource = fsSource
        p_sCondition = fsSearch
        p_sColName = fsColName
        p_sColTitle = fsColTitle
        p_sColFormat = fsColFormat
        p_sFieldName = fsFieldName
        p_nDefCol = fnDefCol
        p_bAutoLoad = fbAutoLoad
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        If p_bLoading Then
            Call cancelQuery()
        End If
        p_bClosing = True
        p_bCancel = True
        Me.Hide()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex = -1 Then
            If e.ColumnIndex <> p_nIndex Then
                ComboBox1.SelectedIndex = e.ColumnIndex
                p_nIndex = e.ColumnIndex
            Else
                txtSearch.Text = ""
            End If
        Else
            cmdLoad.PerformClick()
        End If
    End Sub

    Private Sub DataGridView1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.GotFocus
        If Not BackgroundWorker1.IsBusy Then
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If p_bSearch Then
            Dim loLastInput As LASTINPUTINFO

            loLastInput.cbSize = Marshal.SizeOf(loLastInput)
            loLastInput.dwTime = 0

            If GetLastInputInfo(loLastInput) Then
                Dim lnInterval As Single = Environment.TickCount - loLastInput.dwTime

                If lnInterval > 500 Then
                    Call loadInitList()
                End If
            End If
        End If
    End Sub

    Private Sub Lookup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not isParamOK() Then
            Me.Close()
        End If

        If Not p_bDisplay Then
            Call setDefaults()
        End If
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        If BackgroundWorker1.IsBusy Then
            BackgroundWorker1.CancelAsync()
            p_bCancelAsync = True
        End If
    End Sub

    Private Sub txtSearch_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearch.Validated
        Call loadInitList()
    End Sub

    Private Sub DataGridView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        Select Case e.KeyCode
            Case Windows.Forms.Keys.Return
                cmdLoad.PerformClick()
        End Select
    End Sub

    Private Sub DataGridView1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Resize
        If Not p_bListed Then Exit Sub

        Dim lnCtr As Integer

        lnCtr = DataGridView1.Height / (DataGridView1.Rows(0).Height + DataGridView1.Rows(0).DividerHeight)

        p_nPageRow = lnCtr - 1
    End Sub

    Private Sub cmdLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLoad.Click
        If p_bLoading Then
            Call cancelQuery()
        End If

        p_oDR = getSelectedRow()

        p_bCancel = False
        p_bClosing = True
        Me.Hide()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()

        p_oDT.Dispose()
        p_oDA.Dispose()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
