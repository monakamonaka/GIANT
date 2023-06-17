
Imports System.Data.Entity
Imports System.Data.SqlClient
Imports System.Data.SQLite
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.ApplicationServices
Imports Microsoft.Web.WebView2.Core
'Imports System.Net.Mime.MediaTypeNames
Imports Microsoft.Web.WebView2.WinForms
Imports stdole
Imports WindowsApplication1.Form1

Public Class Form1

    Public Sub New()
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()
        ' InitializeComponent() 呼び出しの後で初期化を追加します。
    End Sub
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '(とりあえず動けばよいという事で、性能は考慮してない)不足データは追々足す
    '-----------------
    '研究所

    Public Structure Lab_Type
        Public Db_Type As String '利用するDB名　’MSSQL　SQLite

        Public Owner As String 'OWNER
        Public OwnerID As Long  'OWNER

        Public XOwner As String 'カレントOWNER
        Public XOwnerID As Long  'カレントOWNER


        Public ProjID As Long  'カレントのRepoID
        Public ProjName As String  'カレントのRepo名

        Public DeskName As String  'カレントのデスク名
        Public DeskID As Long  'カレントのデスクID

        Public DeskSx As Integer  'カレントのデスクのシフト
        Public DeskSy As Integer  'カレントのデスクのシフト
        Public DeskZm As Integer  'カレントのデスクのシフト

        ' Public DeskF As Frame_Type  'カレントのデスク状態

        Public MapF As Frame_Type  'カレントのデスク状態


        Public TabID As Long  'カレントのタブID
        '   Public TabP As Panel 'カレントのパネル

        Public BoxID As Long  'カレントのBOXID
        Public Box As Panel  'カレントのBASE
        Public BluePrint As String  'カレントのBluePrint
        Public Line As Panel  'カレントのラインパネル
        Public LineID As Long  'カレントのラインID
        Public DbPath As String

        Public DBACS_Type As Dictionary(Of String, String)
        Public DBACS_Field As Dictionary(Of String, String)

        Public ProjAT As ProjAT_Type
        Public Proj() As Proj_Type


        Public Person() As Person_Type

        Public DeskAT As DeskAT_Type 'Desk
        Public Desk() As Desk_Type

        Public Base() As Base_Type
        Public Pbox() As Pbox_Type
        Public Tbox() As Tbox_Type

        Public StackAT As ElementAT_Type 'ノードの格納域RepoEleTbl_Type 
        'Public Stuck() As Element_Type

        Public ElementAT As ElementAT_Type 'ノードの格納域RepoEleTbl_Type 
        Public Element() As Element_Type

        'Public RoleAT As RoleAT_Type
        Public Role() As Role_Type



        Public Image_List As Image_List_Type
        Public ToolTip As ToolTip
        Public Flg_BoxMouse As String 'BOX上のマウスの状態フラグBOXに入ると１　マウスが動くと０
        Public Flg_DeskMove As Integer 'デスクを動かすと１、動いてなければ０
        ' Public Columns As Columns_Type
    End Structure
    Public Structure ProjAT_Type
        Public View As ListView_Type 'リスト表示
        Public Menu As ContextMenuStrip
    End Structure
    Public Structure Proj_Type
        Public ProjID As Long
        Public OwnerID As Long
        Public Owner As String
        Public Name As String
        Public Title As String
        Public Summary As String
        Public Hypothesis As String
        Public Purpose As String
        Public Problem As String

        Public Member() As Person_Type
        Public Role() As String

        Public PID() As Long  '引用元のProj番号

    End Structure
    '人事情報

    Public Structure PersonAT_Type
        Public DBACS As DBACS_Type     'データDBアクセスタイプ
        ' Public DBACS_ProjPerson As DBACS_Type


    End Structure
    Public Structure Person_Type
        Public ID As Long
        Public Name As String
        Public Prof As String
        Public Mail As String
        Public Tel As String
        Public Address As String
        Public Image As String 'URLファイルパスが本来だがこうしておく

    End Structure
    Public Structure RoleAT_Type
        Public DBACS As DBACS_Type     'データDBアクセスタイプ

    End Structure
    Public Structure Role_Type
        Public ID As Long
        Public LID As Integer
        Public PID As Integer
        Public Role As String

    End Structure
    'Role
    '筆頭著者	First Author	
    '第二執筆者	Second Author	
    '第三執筆者	Third Author	
    '執筆者	Author	
    '最終著者	Last Author	
    '責任著者	Corresponding Author	
    '研究責任者	Principal Investigator	
    '協力者	Collaborator	
    '研究責任者	Research Director
    '研究所長 Director Of Laboratory
    '研究所長 lab chief


    'Public Structure RepoDB_Type
    '    Public Db As DBACS_Type     'データDBアクセスタイプ


    'End Structure
    ''Base情報
    '--------------------------------
    Public Structure BaseAT_Type
        Public DBACS As DBACS_Type     'データDBアクセスタイプ

    End Structure

    ' 'Base情報
    ' --------------
    Public Structure Base_Type
        Public BaseID As Long
        Public PageID As Integer
        Public PageName As String
        Public Type As String
        Public Name As String

        '位置サイズ情報
        Public Style As Integer 'Fill 
        Public X As Integer
        Public Y As Integer
        Public W As Integer
        Public H As Integer
        Public LV As Integer

    End Structure

    Public Structure Frame_Type
        Public ID As Long
        Public Name As String

        '位置サイズ情報
        Public Style As Integer 'Fill 
        Public X As Integer
        Public Y As Integer
        Public W As Integer
        Public H As Integer

        Public FX As Integer
        Public FY As Integer
        Public FW As Integer
        Public FH As Integer


        Public LV As Integer
        Public ZM As Integer
    End Structure
    ''エレメント情報
    '--------------------------------
    Public Structure ElementAT_Type

        Public View As ListView_Type

        Public Image As ImageList
        Public ImageID As Integer

    End Structure

    ' 'エレメント情報
    '--------------------------------
    Public Structure Element_Type
        Public ID As Long
        ' Public PageID As Integer
        '  Public Pagename As String
        Public Name As String
        Public Type As String
        Public Path As String

        Public Note As String
    End Structure
    Public Structure Stack_Type
        Public ID As Long
        Public PageID As Integer
        Public Pagename As String
        Public Name As String
        Public Type As String

        '位置サイズ情報
        Public State As Integer 'Fill 1 ,else 0
        Public X As Integer
        Public Y As Integer
        Public W As Integer
        Public H As Integer
        Public LV As Integer

        Public Box As Box_Type
        Public BluePrint As String
        Public Path As String

        Public Note As String
    End Structure

    Public Structure Image_List_Type
        Public Key() As String
        Public Image As ImageList
        Public Sub Int(ByVal Hmax)
            ReDim Preserve Key(Hmax)
            Image = New ImageList
        End Sub

    End Structure
    Public Structure RepoEleTbl_Type    'レポジトリとエレメントの対応テーブル
        Public ID As Long
        Public PrID As Long
        Public ElID As Long

    End Structure

    ''作業部屋情報
    ''--------------------------------
    Public Structure DeskAT_Type
        'Public DBACS As DBACS_Type     'データDBアクセスタイプ

        Public View As TabControl '表示タイプ

    End Structure

    '作業机情報
    '--------------------------------
    Public Structure Desk_Type

        Public ID As Long
        Public ProjID As Long
        Public OwnerID As Long
        Public State As Integer
        Public DeskNo As Integer

        Public Type As String 'Desk名　
        Public Name As String 'Desk名　
        Public Text As String 'Desk名　
        Public ShiftX As Integer 'DeskSIFT
        Public ShiftY As Integer 'DeskSIFT
        Public ZM As Integer 'DeskSIFT

        Public Tab As System.Windows.Forms.TabPage    'LabのDeskTop
    End Structure


    Public Structure ListView_Type
        Public Name As String
        'コントロール
        Public ListView As System.Windows.Forms.ListView     'LabのList作業の対象のオブジェクト
        'カラム情報
        Public Columns As Columns_Type

    End Structure


    Public Structure RepoBlockType
        Public Name As String
        Public NodeList() As Long
    End Structure
    Public Structure Columns_Type
        Public Type As String
        Public DBKeys As String 'DBの戻り値
        Public Name() As String
        Public Width() As String　'文字列にしてあるので適用時に数値化すること
    End Structure

    Public Structure Box_Type
        'Public Name As String
        'Public Type As String

        Public Code As String
        Public Dataset As DataSet
        Public WID As Long
        Public OwnerID As Long
        Public Base As Panel
        Public Split() As SplitContainer
        Public Pict() As PictureBox
        Public Text() As TextBox
        Public WebB() As WebView2
        Public Draw() As AxMSINKAUTLib.AxInkPicture
        Public CogT() As AxINKEDLib.AxInkEdit
        Public Combo() As ComboBox
        Public Tab() As TabControl
        Public TabPage() As TabPage
        Public List() As ListView
        Public Btn() As Button



    End Structure


    Public Structure Pbox_Type
        Public ID As Long
        Public BaseID As Long
        Public BoxType As String
        Public BoxPath As String
        Public Style As Integer
        Public PictureBox As PictureBox
        ' Public DbAcsKey As String 後で書き直す
        Public DBACS As DBACS_Type
    End Structure

    Public Structure Tbox_Type
        Public ID As Long
        Public BaseID As Long
        Public BoxType As String
        Public BoxPath As String
        Public Style As Integer
        Public TextBox As TextBox
        ' Public DbAcsKey As String 後で書き直す
        Public DBACS As DBACS_Type
    End Structure


    Public Structure Node_Type
        Public ID As Long
        Public State As Integer     'State
        Public NodeType As String
        Public Name As String
        Public text As String
        Public Owner As String
        Public OwnerID As Long
        Public PID As Integer     'ノードの 親ID
        Public RID As Integer     'レポジトリID

        'Public DT() As DicTbl_Type    'DicTbl DType単位
        'Public DS() As DicTbl_Type    'DicTbl Suggest情報 


        ' コピーを作成するメソッド
        Public Function Clone() As Node_Type
            Dim cloned As Node_Type = CType(MemberwiseClone(), Node_Type)
            ' 参照型フィールドの複製を作成する
            cloned.Name = CType(MyClass.Name.Clone, String)
            If MyClass.text IsNot Nothing Then cloned.text = CType(MyClass.text.Clone, String)
            'cloned.Path = CType(MyClass.Path.Clone, String)
            'If Not MyClass.DT Is Nothing Then cloned.DT = CType(MyClass.DT.Clone(), DicTbl_Type())
            'If Not MyClass.DS Is Nothing Then cloned.DS = CType(MyClass.DS.Clone(), DicTbl_Type())
            Return cloned
        End Function
    End Structure



    Public Structure DBACS_Type

        Public DB_Path As String    'DBへのパス
        Public DB_Name As String    'DBの名前

        Public TBL_Name As String
        Public TBL_Type As String
        Public TBL_Field As String


        ' コピーを作成するメソッド
        Public Function Clone() As DBACS_Type
            Dim cloned As DBACS_Type = CType(MemberwiseClone(), DBACS_Type)
            ' 参照型フィールドの複製を作成する
            cloned.DB_Path = CType(MyClass.DB_Path.Clone, String)
            cloned.DB_Name = CType(MyClass.DB_Name.Clone, String)

            cloned.TBL_Type = CType(MyClass.TBL_Type.Clone, String)
            cloned.TBL_Field = CType(MyClass.TBL_Field.Clone, String)
            cloned.TBL_Name = CType(MyClass.TBL_Name.Clone, String)

            Return cloned
        End Function

    End Structure










    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■




    Public Structure Deco_Type

        Public BackColor As String
        Public ForeColor As String
    End Structure



    Public Class ListViewItemComparer
        Implements IComparer
        Private _column As Integer

        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim TmpR As Integer
            Dim Text1, Text2 As String
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)
            If itemx Is Nothing Or itemy Is Nothing Then Return 0
            If itemx.SubItems(_column).Text = "" Then Return 0
            If itemy.SubItems(_column).Text = "" Then Return 0

            'xとyを文字列として比較する
            Text1 = itemx.SubItems(_column).Text
            Text2 = itemy.SubItems(_column).Text

            TmpR = String.Compare(Text1, Text2)
            Return TmpR

        End Function
    End Class
    Public Class ListViewItemComparer_F
        Implements IComparer
        Private _column As Integer

        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim TmpR As Integer
            Dim Text1, Text2 As String
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)
            If itemx Is Nothing Or itemy Is Nothing Then Return 0

            Try
                If itemx.SubItems(_column).Text = "" Then Return 0
                If itemy.SubItems(_column).Text = "" Then Return 0
                'xとyを文字列として比較する
                Text1 = itemx.SubItems(_column).Text
                Text2 = itemy.SubItems(_column).Text

                TmpR = String.Compare(Text1, Text2)
                Return TmpR

            Catch


            End Try
            Return 0
        End Function
    End Class

    Public Class ListViewItemComparer_L

        Implements IComparer
        Private _column As Integer

        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim TmpR As Integer
            Dim Text1, Text2 As String
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)
            If itemx Is Nothing Or itemy Is Nothing Then Return 0
            If itemx.SubItems(_column).Text = "" Then Return 0
            If itemy.SubItems(_column).Text = "" Then Return 0

            'xとyを文字列として比較する
            Text1 = itemx.SubItems(_column).Text
            Text2 = itemy.SubItems(_column).Text

            TmpR = String.Compare(Text1, Text2)
            Return -TmpR

        End Function
    End Class

    Public Class ListViewItemComparer_BF
        Implements IComparer
        Private _column As Integer

        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim TmpR As Integer
            Dim Text1, Text2 As String
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)
            If itemx Is Nothing Or itemy Is Nothing Then Return 0
            If itemx.SubItems(_column).Text = "" Then Return 0
            If itemy.SubItems(_column).Text = "" Then Return 0

            'xとyを文字列として比較する
            Text1 = StrReverse(itemx.SubItems(_column).Text)
            Text2 = StrReverse(itemy.SubItems(_column).Text)

            TmpR = String.Compare(Text1, Text2)
            Return TmpR

        End Function
    End Class

    Public Class ListViewItemComparer_BL
        Implements IComparer
        Private _column As Integer

        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim TmpR As Integer
            Dim Text1, Text2 As String
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)
            If itemx Is Nothing Or itemy Is Nothing Then Return 0
            If itemx.SubItems(_column).Text = "" Then Return 0
            If itemy.SubItems(_column).Text = "" Then Return 0

            'xとyを文字列として比較する
            Text1 = StrReverse(itemx.SubItems(_column).Text)
            Text2 = StrReverse(itemy.SubItems(_column).Text)

            TmpR = String.Compare(Text1, Text2)
            Return -TmpR

        End Function
    End Class



    'グローバル
    '----------------------------------------------------
    '----------------------------------------------------
    '----------------------------------------------------

    Public gLab As Lab_Type         'Lab　グローバル

    Public gX_owner As String             '現在の作業者
    Public gX_ownerID As Long

    Public gOwner As String             'LABの所有者
    Public gOwnerID As Long
    Public gDeco As Deco_Type          'デザインカラー
    Public gApp_Path As String          'EXE位置
    Public gIcon_Path As String         'Iconデータへのパス  
    Public gWORK_Path As String         '各種WORKデータへのパス  
    Public gENV_Path As String         '各種設定データへのパス  
    Public gLOG_Path As String         'ログ出力先のパス  
    Public gDB_Path As String
    Public gDB_Acs As String        'DBへのアクセスキー
    Public gImg_List As New Image_List_Type     'イメージリスト

    Public gSelectNode As TreeNode      '現在フォーカスしてるRepoNode 
    Public gView(5) As Integer  '画面表示制御

    Public gProjID As String         'カレントのレポジトリID  
    Public gBox() As Box_Type        'ラッパーデータ  



    Public CatName() As String
    '----------------------------------------------------
    '----------------------------------------------------
    '----------------------------------------------------
    '----------------------------------------------------

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '●●
        ' gLab.DBname = "GIANT_"
        gLab.Db_Type = "MSSQL"
        gLab.Db_Type = "MSSQLW"
        gLab.Db_Type = "SQLite"

        Lab_Set_Env(gLab) '環境設定
        Handler_Set() ' 'イベントハンドラを追加する

        'OWNERの設定
        '-----------------------------------
        gX_ownerID = Owner_Get_ID()

        If gX_ownerID = 0 Then
            'ユーザ仮登録
            '  gX_ownerID = DBAcs_Insert(gLab, "Person", "Name,Tel,Mail,Prof", "'Unknown','Tel','Mail','Prof'")  '   ユーザー登録
            gX_ownerID = 1
            Owner_Save_ID(gX_ownerID)
            gLab.OwnerID = gX_ownerID
            'プロジェクト新設
            Proj_Make_New(gLab)

        Else
            gLab.OwnerID = gX_ownerID
            Owner_Set_Env(gLab, gX_ownerID)
            Owner_Show_Data(gLab)
        End If

        '前回までの作業の再開
        '----------------------
        Lab_Start(gLab)


        'らっぱーの設定 現在は全部出すが本来はユーザ依存
        Box_Show_List()
        ' Box_Set_Env()





        'Proj_Show_ListByCmd(gLab, " OwnerID = " & gLab.OwnerID)
        ''WebView21.CoreWebView2.Navigate("https://app.slack.com/client/T056HSRJQ3D/D056WKNQYAV?geocode=ja-jp")
        '' WebView2.CoreWebView2.Navigate("https://freecalend.com/open/mem179529/index")
        'InitializeAsync()

    End Sub
    Function Person_Set_NewOwner() As Long

        Lab_Start(gLab)



    End Function


    Private Sub Proj_Show_ListByCmd(ByRef Lab As Lab_Type, Cmd As String)
        '研究リストを表示
        '------------------
        Dim DS As DataSet
        With Lab
            DS = DBAcs_Get_DataS(Lab, "Proj", .ProjAT.View.Columns.DBKeys, Cmd)

            Button3.Text = "Hits　　" & DS.Tables(0).Rows.Count.ToString
            Proj_Show_List(.ProjAT.View.ListView, .ProjAT.View.Columns, DS, True)
        End With

    End Sub

    Private Sub Lab_Start(ByRef Lab As Lab_Type)
        Proj_Show_ListByCmd(gLab, " OwnerID = " & gX_ownerID) 'オーナーのレポ表示
        '    ListView1.Items(ListView1.Items.Count - 1).Selected = True '(本来なら終了する前のカレントノードだが今はさぼる)

        'gSelectNode = Lab.Desk(1).Node_Opened
        ''Proj_Set_Env(Lab) 'Proj環境設定
        'Proj_Read_Dat(Lab, gOwnerID) 'Projの読み込み
        ''Desk_Set_Env(Lab) 'Desk環境設定
        ''Desk_Read_NodeData(Lab) 'Desk読み込み
        ''Desk_Show_NodeAll(Lab) 'Desk表示
        'Element_Set_Env(Lab) 'ElementDesk環境設定


    End Sub


    Sub Handler_Set()

        'イベントハンドラを追加する
        'Treeview1
        'AddHandler TreeView1.ItemDrag, AddressOf TreeView1_ItemDrag
        'AddHandler TreeView1.DragOver, AddressOf TreeView1_DragOver
        'AddHandler TreeView1.DragDrop, AddressOf TreeView1_DragDrop

        'AddHandler Panel1.DragOver, AddressOf Panel1_DragOver
        ' AddHandler Panel1.DragDrop, AddressOf Panel1_DragDrop

        ''ListView1
        AddHandler ListView1.ItemDrag, AddressOf ListView1_ItemDrag
        'AddHandler ListView1.DragOver, AddressOf ListView1_DragOver
        'AddHandler ListView1.DragDrop, AddressOf ListView1_DragDrop

        ''ListView2
        AddHandler ListView2.ItemDrag, AddressOf ListView2_ItemDrag
        AddHandler ListView2.DragOver, AddressOf ListView2_DragOver
        AddHandler ListView2.DragDrop, AddressOf ListView2_DragDrop

        ''ListView3
        AddHandler ListView3.ItemDrag, AddressOf ListView2_ItemDrag
        AddHandler ListView3.DragOver, AddressOf ListView2_DragOver
        AddHandler ListView3.DragDrop, AddressOf ListView2_DragDrop
        AddHandler ListView5.ItemDrag, AddressOf ListView5_ItemDrag
        'AddHandler ListView6.ItemDrag, AddressOf ListView6_ItemDrag

        'AddHandler TabControl3.SelectionChanged, AddressOf TabControl1_SelectionChanged

        'ColumnClickイベントハンドラの追加
        AddHandler ListView1.ColumnClick, AddressOf ListView1_ColumnClick
        AddHandler ListView2.ColumnClick, AddressOf ListView1_ColumnClick

        'TabControl1.SelectTab(1)


        'AddHandler RichTextBox1.LostFocus, AddressOf RichTextBox1_LostFocus
        'AddHandler RichTextBox2.LostFocus, AddressOf RichTextBox2_LostFocus


        ' AddHandler TabControl1.SelectionChanged, AddressOf TabControl1_SelectionChanged

        'AddHandler PictureBox2.Paint, New PaintEventHandler(AddressOf PictureBox2_Paint)
        'New PaintEventHandler()
        '  ProgressBar1.Controls.Add(Info2)

        'Label1の位置をPictureBox1内の位置に変更する
        'Info2.Top = Info2.Top - ProgressBar1.Top
        'Info2.Left = Info2.Left - ProgressBar1.Left

        'Info2.Top = ProgressBar1.Top + 10
        'Info2.Left = ProgressBar1.Left

    End Sub
    Private Sub Proj_Save_Text(Lab As Lab_Type, RID As Long, Type As String, Text As String)

        'With Lab.ProjAT.DBACS
        '    DBAcs_Update_Data_ByKey(.Db, .Db.TBL_Name, RID, Type & " = '" & Text & "'")
        'End With

    End Sub


    Private Sub RichTextBox1_LostFocus(ByVal sender As Object, ByVal e As EventArgs)
        Proj_Save_Text(gLab, gProjID, "Note", RichTextBox1.Text)


    End Sub



    'Private Sub TabControl1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs)
    '    '選択されたタブの番号を取得
    '    Dim selectedIndex As Integer = TabControl1.SelectedIndex + 1

    '    '選択されたタブの情報を表示
    '    TextBlock1.Text = selectedIndex.ToString() + "番目のタブが選択されました"
    'End Sub


    Private Sub ListView1_ColumnClick(ByVal sender As Object, ByVal e As ColumnClickEventArgs)
        'ListViewItemSorterを指定する

        If e.Column <> 1 And e.Column <> 2 Then Return

        Dim Key, Type As String

        With ListView1



            Key = Mid(.Columns(e.Column).Text, 2)
            Type = Mid(.Columns(e.Column).Text, 1, 1)
            If System.Windows.Forms.Control.ModifierKeys = Keys.Control Then

                Select Case Type
                    Case "▲"
                        Type = "▼"
                    Case "▼"
                        Type = "▲"
                    Case "▽"
                        Type = "▲"
                    Case "△"
                        Type = "▼"
                    Case Else
                        Type = "▲"
                        Key = .Columns(e.Column).Text
                End Select
            Else
                Select Case Type
                    Case "▲"
                        Type = "▽"
                    Case "▼"
                        Type = "▲"
                    Case "▽"
                        Type = "△"
                    Case "△"
                        Type = "▽"
                    Case Else
                        Type = "△"
                        Key = .Columns(e.Column).Text
                End Select
            End If

            .Columns(e.Column).Text = Type & Key

            '並び替える（ListViewItemSorterを設定するとSortが自動的に呼び出される）
            Select Case Type
                Case "▽"
                    .ListViewItemSorter = New ListViewItemComparer_F(e.Column)

                Case "△"
                    .ListViewItemSorter = New ListViewItemComparer_L(e.Column)
                Case "▲"
                    .ListViewItemSorter = New ListViewItemComparer_BF(e.Column)

                Case "▼"
                    .ListViewItemSorter = New ListViewItemComparer_BL(e.Column)

            End Select

        End With

    End Sub

    'Private Sub ListView3_ColumnClick(ByVal sender As Object, ByVal e As ColumnClickEventArgs)
    '    'ListViewItemSorterを指定する

    '    If e.Column <> 1 And e.Column <> 2 Then Return
    '    Dim Key, Type As String

    '    With ListView3
    '        Key = Mid(.Columns(e.Column).Text, 2)
    '        Type = Mid(.Columns(e.Column).Text, 1, 1)
    '        If System.Windows.Forms.Control.ModifierKeys = Keys.Control Then

    '            Select Case Type
    '                Case "▲"
    '                    Type = "▼"
    '                Case "▼"
    '                    Type = "▲"
    '                Case "▽"
    '                    Type = "▲"
    '                Case "△"
    '                    Type = "▼"
    '                Case Else
    '                    Type = "▲"
    '                    Key = .Columns(e.Column).Text
    '            End Select
    '        Else
    '            Select Case Type
    '                Case "▲"
    '                    Type = "▽"
    '                Case "▼"
    '                    Type = "▲"
    '                Case "▽"
    '                    Type = "△"
    '                Case "△"
    '                    Type = "▽"
    '                Case Else
    '                    Type = "△"
    '                    Key = .Columns(e.Column).Text
    '            End Select
    '        End If

    '        .Columns(e.Column).Text = Type & Key

    '        '並び替える（ListViewItemSorterを設定するとSortが自動的に呼び出される）
    '        Select Case Type
    '            Case "▽"
    '                .ListViewItemSorter = New ListViewItemComparer_F(e.Column)

    '            Case "△"
    '                .ListViewItemSorter = New ListViewItemComparer_L(e.Column)
    '            Case "▲"
    '                .ListViewItemSorter = New ListViewItemComparer_BF(e.Column)

    '            Case "▼"
    '                .ListViewItemSorter = New ListViewItemComparer_BL(e.Column)

    '        End Select

    '    End With


    'End Sub


    'Private Sub TreeView1_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
    '    'ノードのドラッグを開始
    '    '--------------------------------
    '    Dim tv As TreeView = CType(sender, TreeView)
    '    tv.SelectedNode = CType(e.Item, TreeNode)
    '    tv.Focus()

    '    'ノードのドラッグを開始する
    '    Dim dde As DragDropEffects = tv.DoDragDrop(e.Item, DragDropEffects.All)

    '    '移動した時は、ドラッグしたノードを削除する
    '    If (dde And DragDropEffects.Move) = DragDropEffects.Move Then
    '        tv.Nodes.Remove(CType(e.Item, TreeNode))
    '    End If
    'End Sub

    'Private Sub TreeView1_DragOver(ByVal sender As Object, ByVal e As DragEventArgs)
    '    'ドラッグしている時
    '    '--------------------------------
    '    TreeView1.Focus()
    '    TreeView_DropOvwer_Check(sender, e)

    'End Sub

    'Private Sub TreeView1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
    '    'ドロップされたとき
    '    '------------------------------------
    '    Proj_Drop_At_TreeView(sender, e)

    'End Sub

    'Private Sub TreeView2_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
    '    'ノードのドラッグを開始
    '    '--------------------------------
    '    Dim tv As TreeView = CType(sender, TreeView)
    '    tv.SelectedNode = CType(e.Item, TreeNode)
    '    tv.Focus()

    '    'ノードのドラッグを開始する
    '    Dim dde As DragDropEffects = tv.DoDragDrop(e.Item, DragDropEffects.All)

    '    '移動した時は、ドラッグしたノードを削除する
    '    If (dde And DragDropEffects.Move) = DragDropEffects.Move Then
    '        tv.Nodes.Remove(CType(e.Item, TreeNode))
    '    End If
    'End Sub



    'Private Sub TreeView2_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
    '    'ドロップされたとき
    '    '------------------------------------
    '    Proj_Drop_At_TreeView(sender, e)

    'End Sub

    Private Sub ListView6_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
        'ノードのドラッグを開始
        '--------------------------------
        Dim Lv As ListView = CType(sender, ListView)
        'Lv.SelectedNode = CType(e.Item, ListView)
        'Lv.Focus()

        'ノードのドラッグを開始する
        '  Dim dde As DragDropEffects = Lv.DoDragDrop(e.Item, DragDropEffects.All)

        ' 複数項目選択？
        If Lv.SelectedItems.Count >= 2 Then
            DoDragDrop(Lv.SelectedItems, DragDropEffects.All)
        Else
            DoDragDrop(e.Item, DragDropEffects.All)
        End If


    End Sub
    Private Sub ListView5_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
        'ノードのドラッグを開始
        '--------------------------------
        Dim Lv As ListView = CType(sender, ListView)
        'Lv.SelectedNode = CType(e.Item, ListView)
        'Lv.Focus()

        'ノードのドラッグを開始する
        '  Dim dde As DragDropEffects = Lv.DoDragDrop(e.Item, DragDropEffects.All)

        ' 複数項目選択？
        If Lv.SelectedItems.Count >= 2 Then
            DoDragDrop(Lv.SelectedItems, DragDropEffects.All)
        Else
            DoDragDrop(e.Item, DragDropEffects.All)
        End If



        ''移動した時は、ドラッグしたノードを削除する
        'If (dde And DragDropEffects.Move) = DragDropEffects.Move Then
        '    Lv.Nodes.Remove(CType(e.Item, TreeNode))
        'End If
    End Sub

    Private Sub ListView1_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
        'ノードのドラッグを開始
        '--------------------------------
        Dim Lv As ListView = CType(sender, ListView)
        'Lv.SelectedNode = CType(e.Item, ListView)
        'Lv.Focus()

        'ノードのドラッグを開始する
        '  Dim dde As DragDropEffects = Lv.DoDragDrop(e.Item, DragDropEffects.All)

        ' 複数項目選択？
        If Lv.SelectedItems.Count >= 2 Then
            DoDragDrop(Lv.SelectedItems, DragDropEffects.All)
        Else
            DoDragDrop(e.Item, DragDropEffects.All)
        End If



        ''移動した時は、ドラッグしたノードを削除する
        'If (dde And DragDropEffects.Move) = DragDropEffects.Move Then
        '    Lv.Nodes.Remove(CType(e.Item, TreeNode))
        'End If
    End Sub

    Private Sub ListView1_DragOver(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドラッグしている時
        '--------------------------------
        ListView2.Focus()
        '   ListView_DropOvwer_Check(sender, e)
    End Sub


    Private Sub ListView1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドロップされたとき
        '------------------------------------
        Proj_Drop_At_ListView(sender, e)

    End Sub


    Private Sub ListView2_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
        'ノードのドラッグを開始
        '--------------------------------
        Dim Lv As ListView = CType(sender, ListView)
        'Lv.SelectedNode = CType(e.Item, ListView)
        'Lv.Focus()

        'ノードのドラッグを開始する
        Dim dde As DragDropEffects = Lv.DoDragDrop(e.Item, DragDropEffects.All)

        ''移動した時は、ドラッグしたノードを削除する
        'If (dde And DragDropEffects.Move) = DragDropEffects.Move Then
        '    Lv.Nodes.Remove(CType(e.Item, TreeNode))
        'End If
    End Sub
    Private Sub Panel1_DragOver(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドラッグしている時
        '--------------------------------
        ListView2.Focus()
        e.Effect = DragDropEffects.Move

    End Sub

    'Private Sub Panel1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
    '    'ドロップされたとき
    '    '------------------------------------
    '    Dim Dat As VariantType
    '    Dim Drops() As String
    '    Dim Element() As Element_Type
    '    Dim cnt As Integer
    '    cnt = 0
    '    ReDim Preserve Element(cnt)

    '    ' FILE
    '    Drops = e.Data.GetData(DataFormats.FileDrop, False)
    '    If Not (Drops Is Nothing) Then

    '        For L = 0 To Drops.Length - 1
    '            Element(cnt).Path = Drops(L)
    '            Element(cnt).Name = System.IO.Path.GetFileNameWithoutExtension(Drops(L))
    '            Element(cnt).Type = "File"
    '            cnt = cnt + 1
    '            ReDim Preserve Element(cnt)
    '        Next

    '    End If

    '    '' Html
    '    'Drops(0) = e.Data.GetData(DataFormats.Html, False)
    '    'If Not (Drops Is Nothing) Then
    '    '    For L = 0 To Drops.Length - 1
    '    '        With Element(cnt)
    '    '            .Path = Drops(L)
    '    '            .Type = "Html"
    '    '            .Name = Mid(.Path, 1, 20) 'タイトルから取るべきだが
    '    '        End With

    '    '        cnt = cnt + 1
    '    '        ReDim Preserve Element(cnt)
    '    '    Next
    '    'End If


    '    ' TEXT
    '    ReDim Drops(0)
    '    Drops(0) = e.Data.GetData(DataFormats.Text, False)
    '    If Not (Drops(0) = "") Then
    '        For L = 0 To Drops.Length - 1
    '            With Element(cnt)
    '                .Path = Drops(L)
    '                If InStr(.Path, "http") > 0 Or InStr(.Path, ".htm") > 0 Then
    '                    .Type = "Url"
    '                    .Name = Mid(.Path, 1, 20) 'タイトルから取るべきだが
    '                Else
    '                    .Type = "Text"
    '                    .Name = Mid(.Path, 1, 20)
    '                End If
    '            End With

    '            cnt = cnt + 1
    '            ReDim Preserve Element(cnt)
    '        Next
    '    End If

    '    ' Bitmap
    '    Dim img As System.Drawing.Image
    '    img = e.Data.GetData(DataFormats.Bitmap, False)
    '    If Not (img Is Nothing) Then
    '        For L = 0 To Drops.Length - 1
    '            With Element(cnt)
    '                .Path = Drops(L)
    '                .Type = "Bitmap"
    '                .Name = "画像"
    '            End With

    '            cnt = cnt + 1
    '            ReDim Preserve Element(cnt)
    '        Next
    '    End If
    '    Element_Add_Data(gLab, gProjID, Element)
    '    'Proj_Action_Files("NOT", DropS)
    '    ' Proj_Drop_At_ListView(sender, e)


    'End Sub


    Private Sub ListView2_DragOver(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドラッグしている時
        '--------------------------------
        ListView2.Focus()
        e.Effect = DragDropEffects.Move

    End Sub


    Private Sub ListView2_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドロップされたとき
        '------------------------------------
        Dim cp As Point = sender.PointToClient(New Point(e.X, e.Y))
        Elements_Drop(e)


    End Sub
    Private Function Elements_Drop(ByRef e As DragEventArgs) As Element_Type()
        'ドロップされたとき
        '------------------------------------
        Dim Dat As VariantType
        Dim Drops() As String
        Dim Element() As Element_Type
        Dim cnt As Integer
        cnt = 0
        ReDim Preserve Element(cnt)

        ' FILE
        Drops = e.Data.GetData(DataFormats.FileDrop, False)
        If Not (Drops Is Nothing) Then

            For L = 0 To Drops.Length - 1
                Element(cnt).Path = Drops(L)
                Element(cnt).Name = System.IO.Path.GetFileNameWithoutExtension(Drops(L))
                Element(cnt).Type = "File"

                '  If flag = 1 Then Box_MakeFromElement(gLab, "FILE", cp, Element(cnt)) 'DESK上
                cnt = cnt + 1
                ReDim Preserve Element(cnt)
            Next


        End If

        '' Html
        'Drops(0) = e.Data.GetData(DataFormats.Html, False)
        'If Not (Drops Is Nothing) Then
        '    For L = 0 To Drops.Length - 1
        '        With Element(cnt)
        '            .Path = Drops(L)
        '            .Type = "Html"
        '            .Name = Mid(.Path, 1, 20) 'タイトルから取るべきだが
        '        End With

        '        cnt = cnt + 1
        '        ReDim Preserve Element(cnt)
        '    Next
        'End If


        ' TEXT
        ReDim Drops(0)
        Drops(0) = e.Data.GetData(DataFormats.Text, False)
        If Not (Drops(0) = "") Then
            For L = 0 To Drops.Length - 1
                With Element(cnt)
                    .Path = Drops(L)
                    If InStr(.Path, "http") > 0 Or InStr(.Path, ".htm") > 0 Then
                        .Type = "Url"
                        .Name = Mid(.Path, 1, 20) 'タイトルから取るべきだが
                    Else
                        .Type = "Text"
                        .Name = Mid(.Path, 1, 20)
                    End If
                End With

                cnt = cnt + 1
                ReDim Preserve Element(cnt)
            Next

        End If

        ' Bitmap
        Dim img As System.Drawing.Image
        img = e.Data.GetData(DataFormats.Bitmap, False)
        If Not (img Is Nothing) Then
            For L = 0 To Drops.Length - 1
                With Element(cnt)
                    .Path = Drops(L)
                    .Type = "Bitmap"
                    .Name = "画像"
                End With

                cnt = cnt + 1
                ReDim Preserve Element(cnt)
            Next
        End If
        Element_Add_Data(gLab, gProjID, Element)
        Return Element

    End Function

    Private Function Elements_Paste() As Element_Type()
        'pasteされたとき
        '------------------------------------

        Dim Drops() As String
        Dim Element() As Element_Type
        Dim cnt As Integer
        cnt = 0
        ReDim Preserve Element(cnt)

        'クリップボードのデータチェック
        Dim data As IDataObject = Clipboard.GetDataObject()
        If data Is Nothing Then Return Nothing　'空


        With My.Computer.Clipboard
            '◆ここから

            'テキスト
            If .ContainsText() Then
                Dim Dat() As String
                ReDim Dat(0)
                Dat(0) = .GetText()

                If Not (Dat(0) = "") Then
                    For L = 0 To Dat.Length - 1
                        With Element(cnt)
                            .Path = Dat(L)
                            If InStr(.Path, "http") > 0 Or InStr(.Path, ".htm") > 0 Then
                                .Type = "Url"
                                .Name = Mid(.Path, 1, 20) 'タイトルから取るべきだが
                            Else
                                If InStr(.Path, "【すたっく】") > 0 Then
                                    .Type = "Stack"
                                Else
                                    .Type = "Text"
                                    .Name = Mid(.Path, 1, 20)
                                End If
                            End If
                        End With

                        cnt = cnt + 1
                        ReDim Preserve Element(cnt)
                    Next

                End If

            End If

            'ファイル
            If .ContainsFileDropList() Then
                Dim Dat As System.Collections.Specialized.StringCollection
                Dat = .GetFileDropList()
                For Each fileName In Dat
                    Element(cnt).Path = fileName
                    Element(cnt).Name = System.IO.Path.GetFileNameWithoutExtension(fileName)
                    Element(cnt).Type = "File"
                    cnt = cnt + 1
                    ReDim Preserve Element(cnt)
                Next fileName
            End If

            '画像
            If .ContainsImage() Then
                Dim Dat As System.Drawing.Image
                Dat = .GetImage()
                PictureBox1.Image = Dat
            End If

            '音声
            If .ContainsAudio() Then

            End If
        End With


        Element_Add_Data(gLab, gProjID, Element)
        Return Element

    End Function




    Private Sub ListView3_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
        'ノードのドラッグを開始
        '--------------------------------
        Dim Lv As ListView = CType(sender, ListView)
        'Lv.SelectedNode = CType(e.Item, ListView)
        'Lv.Focus()

        'ノードのドラッグを開始する
        '  Dim dde As DragDropEffects = Lv.DoDragDrop(e.Item, DragDropEffects.All)

        ' 複数項目選択？
        If Lv.SelectedItems.Count >= 2 Then
            DoDragDrop(Lv.SelectedItems, DragDropEffects.All)
        Else
            DoDragDrop(e.Item, DragDropEffects.All)
        End If



        ''移動した時は、ドラッグしたノードを削除する
        'If (dde And DragDropEffects.Move) = DragDropEffects.Move Then
        '    Lv.Nodes.Remove(CType(e.Item, TreeNode))
        'End If
    End Sub





    'Private Sub Proj_Drop_At_TreeView(ByVal sender As Object, ByVal e As DragEventArgs)
    '    'TreeViewにドロップしたとき
    '    '------------------------------------
    '    'Proj_Actで始まる関数は、基本的にHISTORYに記憶して、進む、戻るの対象になるもの


    '    '--------------------
    '    'タイプ別アクション
    '    '--------------------
    '    Dim Act_Type As String
    '    Dim Act_Effect As String
    '    Dim Type_T(2), Type_N(2), Type_S(2) As String

    '    Act_Type = ""
    '    'Action設定
    '    Act_Type = ""
    '    Select Case e.Effect
    '        Case 1
    '            Act_Type = "COPY"
    '        Case 2
    '            Act_Type = "MOVE"
    '    End Select

    '    Act_Effect = ""


    '    '--------------------
    '    'ドロップ先のTreeView
    '    '--------------------

    '    Dim Target_Tree As TreeView = CType(sender, TreeView)
    '    'ドロップ先のTreeNodeを取得する
    '    Dim Target_Node As TreeNode = Target_Tree.GetNodeAt(Target_Tree.PointToClient(New Point(e.X, e.Y)))



    '    '-----------------
    '    'ドロップチェック
    '    '-----------------

    '    '------------------
    '    'ドロップ元がTreeNode
    '    '------------------
    '    If e.Data.GetDataPresent(GetType(TreeNode)) Then
    '        e.Effect = DragDropEffects.None

    '        'ドロップされたデータ(TreeNode)を取得
    '        Dim Source_Node As TreeNode = CType(e.Data.GetData(GetType(TreeNode)), TreeNode)

    '        'それぞれのアクションに振り分け
    '        'Act_Effect = Proj_Drop_Select_TreeViewToTreeView(gLab, Source_Node, Target_Node, Act_Type)

    '        Select Case Act_Effect
    '            Case "Non"
    '                e.Effect = DragDropEffects.None
    '            Case "Move" '元ノードを削除
    '                e.Effect = DragDropEffects.Move
    '        End Select

    '        Return
    '    End If


    '    '------------------
    '    'ドロップ元がListViewItem
    '    '------------------
    '    Dim Dmax As Integer
    '    Dim Source() As ListViewItem
    '    'ドロップされたデータがListViewItemか調べる
    '    Dmax = 0
    '    If e.Data.GetDataPresent(GetType(ListViewItem)) Then
    '        ReDim Source(Dmax)
    '        Source(Dmax) = CType(e.Data.GetData(GetType(ListViewItem)), ListViewItem)

    '        'それぞれのアクションに振り分け
    '        Proj_Drop_Select_ListViewToTreeView(gLab, Source, Target_Node, Act_Type)


    '    End If

    '    'ドロップされたデータがListViewItem(複数)か調べる
    '    If e.Data.GetDataPresent(GetType(ListView.SelectedListViewItemCollection)) Then
    '        Dim SourceS As ListView.SelectedListViewItemCollection = CType(e.Data.GetData(GetType(ListView.SelectedListViewItemCollection)), ListView.SelectedListViewItemCollection)
    '        If Not (SourceS Is Nothing) Then

    '            For L = 0 To SourceS.Count - 1
    '                ReDim Preserve Source(Dmax)
    '                Source(Dmax) = SourceS.Item(L)
    '                Dmax += 1
    '            Next

    '            'それぞれのアクションに振り分け
    '            Proj_Drop_Select_ListViewToTreeView(gLab, Source, Target_Node, Act_Type)

    '        End If
    '    End If


    'End Sub


    Private Sub Proj_Drop_At_ListView(ByVal sender As Object, ByVal e As DragEventArgs)
        'ListViewにドロップしたとき
        '------------------------------------      
        '現在はなにもしない


    End Sub

    Private Sub Desk_Paste(cp As Point)
        ' Deskにペーストしたとき
        '----------------------------------



        Dim DeskID As Integer
        DeskID = gLab.DeskID


        Dim Elm() As Element_Type
        Elm = Elements_Paste() 'DBに登録
        If Elm IsNot Nothing Then Desk_Drops_Show(gLab, Elm, DeskID, cp)



        Desk_Drops_Show(gLab, Elm, gLab.DeskID, cp)

    End Sub
    Private Sub Desk_Drops(ByVal sender As Object, ByVal e As DragEventArgs)
        ' Deskにドロップしたとき
        '------------------------------------

        '-----------------
        'ドロップチェック
        '-----------------
        Dim DropType As String
        'ドロップされたデータが何か調べる
        If e.Data.GetDataPresent(GetType(ListViewItem)) Then DropType = "Box"
        If e.Data.GetDataPresent(GetType(ListViewItem)) Then

        End If

        Dim cp As Point = sender.PointToClient(New Point(e.X, e.Y))

        gLab.DeskID = Val(sender.name)
        '------------------
        'ドロップ元がListViewItem スタック
        '------------------
        Dim Dmax As Integer
        Dim Source() As ListViewItem
        Dmax = 0

        If e.Data.GetDataPresent(GetType(ListViewItem)) Then
            ReDim Source(Dmax)
            Source(Dmax) = CType(e.Data.GetData(GetType(ListViewItem)), ListViewItem) '複数であっても１つだけ

            With gLab
                .DeskID = gLab.DeskID
                .DeskName = gLab.DeskID.ToString
                .BluePrint = Box_Get_BluePrint(gLab, Source(0).Tag) '設計
                .BluePrint = Replace(.BluePrint, "[BX]", cp.X.ToString)
                .BluePrint = Replace(.BluePrint, "[BY]", cp.Y.ToString)
                Box_Makes_OnDESK(gLab, .BluePrint)
            End With
            Return

        End If
        '------------------
        'ドロップ元がListViewItem ELEMENT
        '------------------
        '------------------
        'ドロップ元が外部　 ELEMENT
        '------------------
        Dim Elm() As Element_Type
        Elm = Elements_Drop(e) 'DBに登録
        If Elm IsNot Nothing Then Desk_Drops_Show(gLab, Elm, gLab.DeskID, cp)


    End Sub
    Private Sub Desk_Drops_Show(ByRef lab As Lab_Type, Elm() As Element_Type, DeskID As Integer, cp As Point)
        For L = 0 To Elm.Count - 1
            With Elm(L)
                If .Name <> Nothing Then
                    Select Case .Type
                        Case "File"

                            With gLab
                                .DeskID = DeskID
                                .DeskName = DeskID.ToString
                                .BluePrint = Box_Get_BluePrint(gLab, "Viewer") '設計★TYPE|★PATH|★NAME"
                                .BluePrint = Replace(.BluePrint, "[BX]", cp.X.ToString)
                                .BluePrint = Replace(.BluePrint, "[BY]", cp.Y.ToString)
                                .BluePrint = Replace(.BluePrint, "★TYPE", Elm(L).Type)
                                .BluePrint = Replace(.BluePrint, "★PATH", Elm(L).Path)
                                .BluePrint = Replace(.BluePrint, "★NAME", Elm(L).Name)
                                Box_Makes_OnDESK(gLab, .BluePrint)
                            End With
                        Case "Url"
                            With gLab
                                .DeskID = DeskID
                                .DeskName = DeskID.ToString
                                .BluePrint = Box_Get_BluePrint(gLab, "Viewer") '設計★TYPE|★PATH|★NAME"
                                .BluePrint = Replace(.BluePrint, "[BX]", cp.X.ToString)
                                .BluePrint = Replace(.BluePrint, "[BY]", cp.Y.ToString)
                                .BluePrint = Replace(.BluePrint, "★TYPE", Elm(L).Type)
                                .BluePrint = Replace(.BluePrint, "★PATH", Elm(L).Path)
                                .BluePrint = Replace(.BluePrint, "★NAME", Elm(L).Name)
                                Box_Makes_OnDESK(gLab, .BluePrint)
                            End With

                        Case "Text"
                            With gLab
                                .DeskID = DeskID
                                .DeskName = DeskID.ToString
                                .BluePrint = Box_Get_BluePrint(gLab, "Post_it") '設計★TYPE|★PATH|★NAME"
                                .BluePrint = Replace(.BluePrint, "[BX]", cp.X.ToString)
                                .BluePrint = Replace(.BluePrint, "[BY]", cp.Y.ToString)
                                .BoxID = Box_SaveDB(gLab, .BluePrint) 'DBに格納
                                Dim STACKS(), STK_PRT() As String

                                STACKS = Split(.BluePrint, "@")
                                STK_PRT = Split(STACKS(3), "|")
                                Dim TxID As Integer
                                TxID = DBAcs_Insert(gLab, "TBox", "BoxID,Num,Text", .BoxID.ToString & "," & STK_PRT(1).ToString & ",' " & Elm(L).Path & "'")
                                .BluePrint = Box_Re_Text(gLab, .BoxID, Val(STK_PRT(1)), "∃" & TxID.ToString)

                                Box_Factory(gLab, .BoxID, .BluePrint)
                            End With

                        Case "Bitmap"
                    End Select

                End If

            End With


        Next
    End Sub

    Private Function BxCogT_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As AxINKEDLib.AxInkEdit
        '@CogT|0|NEXT|Dock|TEXT|"
        Dim CogT As AxINKEDLib.AxInkEdit
        CogT = New AxINKEDLib.AxInkEdit

        'AddHandler TBox.DragEnter, AddressOf BxText_DragEnter
        'AddHandler TBox.DragDrop, AddressOf BxText_Drop

        OBJ.Controls.Add(CogT)
        With CogT
            .Dock = STK_PRT(3)

        End With
        Return CogT

        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function
    Private Function BxDraw_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As AxMSINKAUTLib.AxInkPicture
        '@Draw|0|NEXT|Dock|DaT| 
        Dim imgconv As New ImageConverter()
        Dim TEGAKI As AxMSINKAUTLib.AxInkPicture
        TEGAKI = New AxMSINKAUTLib.AxInkPicture

        'TEGAKI.Ink.
        'myink = New MSINKAUTLib.InkDisp
        'Dim DrID As Long

        'OBJ.Controls.Add(TEGAKI)
        'TEGAKI.Dock = STK_PRT(3)

        ''AddHandler TEGAKI.DragEnter, AddressOf BxTEGAKI_DragEnter
        ''AddHandler Draw.DragDrop, AddressOf BxDraw_Drop
        'AddHandler TEGAKI.MouseLeaveEvent, AddressOf BxDraw_MouseLeave
        ''  TEGAKI.SizeMode = PictureBoxSizeMode.Zoom
        'TEGAKI.AllowDrop = True

        'Select Case Mid(STK_PRT(4), 1, 1)
        '    Case "∇" 'DB登録
        '        '記録場所を確保
        '        DrID = DBAcs_Insert(Lab, "Bdat", "BoxID,Num,Type", Lab.BoxID.ToString & "," & STK_PRT(1).ToString & ",'Draw'")
        '        'Blueprint修正 (DB側のみ)
        '        Box_Re_Draw(Lab, Lab.BoxID, Val(STK_PRT(1)), "∃" & DrID.ToString)
        '        TEGAKI.Tag = DrID.ToString
        '        Dim TEGAKIData() As Byte
        '        TEGAKIData = TEGAKI.Ink.Save()
        '        My.Computer.FileSystem.WriteAllBytes(gWORK_Path & "\DDX" & DrID.ToString, TEGAKIData, True)
        '    Case "∃"
        '        DrID = Val(Mid(STK_PRT(4), 2))
        '        TEGAKI.Tag = DrID.ToString
        '        Dim TEGAKIData() As Byte

        '        TEGAKIData = My.Computer.FileSystem.ReadAllBytes(gWORK_Path & "\DDX" & DrID.ToString)

        '        'TEGAKI.InkEnabled = False
        '        ' TEGAKI.Ink.Load(CType(imgconv.ConvertFrom(TEGAKIData), Image))
        '        TEGAKI.Ink.Load(TEGAKIData)
        '        'TEGAKI.InkEnabled = True
        '        'TEGAKI.InkEnabled = True
        '    Case Else

        'End Select
        'TEGAKI.Name = TEGAKI.ToString

        Return TEGAKI
    End Function



    Private Function BxTabP_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As TabPage
        '@TABp|0|TAB|Text|BackColor|"
        Dim Tabp As TabPage
        Tabp = New TabPage

        'AddHandler TBox.DragEnter, AddressOf BxText_DragEnter
        'AddHandler TBox.DragDrop, AddressOf BxText_Drop

        OBJ.Controls.Add(Tabp)
        With Tabp
            .Text = STK_PRT(3)
            .BackColor = Color.FromName(STK_PRT(4))
        End With
        Return Tabp

        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function

    Private Function BxTaB_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As TabControl
        '@TAB|0|BASE|Dock|Alignment|ItemSizeX|ItemSizeY|"
        Dim Tab As TabControl
        Tab = New TabControl

        'AddHandler TBox.DragEnter, AddressOf BxText_DragEnter
        'AddHandler TBox.DragDrop, AddressOf BxText_Drop

        OBJ.Controls.Add(Tab)
        With Tab
            .Dock = STK_PRT(3)
            .Alignment = Val(STK_PRT(4))
            .ItemSize = New Size(Val(STK_PRT(5)), Val(STK_PRT(6)))
        End With
        Return Tab

        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function



    Private Function BxSPLIT_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As SplitContainer
        ' '@Split|0|NEXT|Dock|Orientation|SplitterDistance
        Dim SPLIT As SplitContainer
        SPLIT = New SplitContainer
        OBJ.Controls.Add(SPLIT)
        With SPLIT
            .Dock = STK_PRT(3)
            .Orientation = Val(STK_PRT(4))
            .SplitterDistance = Val(STK_PRT(5))
            If STK_PRT.Count > 7 Then
                If STK_PRT(6) <> "" Then .Width = Val(STK_PRT(6))
                If STK_PRT(7) <> "" Then .Height = Val(STK_PRT(7))

            End If
        End With
        Return SPLIT

    End Function
    Private Function BxWeb_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As WebView2
        ' "@WebView|0|Wb.Hed|Wb.Dock|Wb.S|"
        Dim WebBox As WebView2
        WebBox = New WebView2

        'AddHandler TBox.DragEnter, AddressOf BxText_DragEnter
        'AddHandler TBox.DragDrop, AddressOf BxText_Drop

        OBJ.Controls.Add(WebBox)
        With WebBox
            .Dock = STK_PRT(3)
            If STK_PRT(4) <> "" Then
                .Source = New Uri(STK_PRT(4))
            End If
        End With
        Return WebBox

        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function

    Private Function BxText_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As TextBox
        ' "@TEXT|0|T.Hed|T.Dock|T.Mul|T.Bc|T.Fc|T.FName|T.FSize|T.Text|"
        '"@TEXT|3|BASE|TBLR|0|16|200|80|-1|lemonchiffon|black|Arial|8|∇"
        Dim TxID As Long
        Dim TBox As TextBox
        TBox = New TextBox
        Dim BluePrint As String
        AddHandler TBox.MouseLeave, AddressOf BxText_MouseLeave
        'AddHandler TBox.DragDrop, AddressOf BxText_Drop

        OBJ.Controls.Add(TBox)
        With TBox
            Select Case STK_PRT(3)
                Case "0", "1", "2", "3", "4", "5"
                    .Dock = Val(STK_PRT(3))
                Case Else
                    For L = 1 To Len(STK_PRT(3))
                        Select Case Mid(STK_PRT(3), L, 1)
                            Case "T"
                                .Anchor = .Anchor Or AnchorStyles.Top
                            Case "B"
                                .Anchor = .Anchor Or AnchorStyles.Bottom
                            Case "L"
                                .Anchor = .Anchor Or AnchorStyles.Left
                            Case "R"
                                .Anchor = .Anchor Or AnchorStyles.Right
                        End Select
                    Next

            End Select

            .Left = Val(STK_PRT(4))
            .Top = Val(STK_PRT(5))

            .Width = OBJ.width - .Left * 2
            .Height = OBJ.height - .Top * 2

            .Multiline = Val(STK_PRT(8))
            .BackColor = Color.FromName(STK_PRT(9))
            .ForeColor = Color.FromName(STK_PRT(10))
            Dim f As New System.Drawing.Font(STK_PRT(11), STK_PRT(12))

            .Font = f
            .ScrollBars = ScrollBars.Vertical
            .BorderStyle = BorderStyle.None
            '.WordWrap = True
            Select Case Mid(STK_PRT(13), 1, 1)
                Case "∇" 'DB登録
                    '記録場所を確保
                    TxID = DBAcs_Insert(Lab, "TBox", "BoxID,Num,Text", Lab.BoxID.ToString & "," & STK_PRT(1).ToString & ",' " & Mid(STK_PRT(13), 2) & "'")
                    'Blueprint修正 (DB側のみ)
                    Box_Re_Text(Lab, Lab.BoxID, Val(STK_PRT(1)), "∃" & TxID.ToString)

                    .Text = Mid(STK_PRT(13), 2)
                Case "∃"
                    Dim Ds As DataSet
                    TxID = Val(Mid(STK_PRT(13), 2))
                    Ds = DBAcs_Get_DataS(Lab, "Tbox", "Text", "ID = " & Mid(STK_PRT(13), 2))
                    ' If Ds.Tables(0).Rows.Count = 0 Then Return  '異常

                    .Text = Ds.Tables(0).Rows(0).Item(0)
                Case Else
                    .Text = STK_PRT(13)
            End Select
            .Name = TxID.ToString

        End With


        Return TBox

        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function

    Private Function BxViewer_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As Object
        ' "@Viewer|2|BASE|★TYPE|★PATH|★NAME"



        Dim Type, Path, Name, Atr As String
        Type = STK_PRT(3)
        Path = STK_PRT(4)
        Name = STK_PRT(5)
        Atr = System.IO.Path.GetExtension(Path)

        Dim WebBox As WebView2
        WebBox = New WebView2
        OBJ.Controls.Add(WebBox)


        Dim flag0 As Integer
        flag0 = 0

        With WebBox
            '後でちゃんと直そうね
            .Anchor = .Anchor Or AnchorStyles.Top
            .Anchor = .Anchor Or AnchorStyles.Bottom
            .Anchor = .Anchor Or AnchorStyles.Left
            .Anchor = .Anchor Or AnchorStyles.Right
            .Left = 0
            .Top = 8

            .Width = OBJ.width - 4
            .Height = OBJ.height - .Top * 2


            Select Case Type
                Case "Url"
                    .Source = New Uri(Path)
                Case "File"
                    If InStr(".txt|.pdf|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & Path)
                        flag0 = 1
                    End If
                    If InStr(".ico|.bmp|.jpg|.gif|.png|.exig|.tiff|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & Path)
                        flag0 = 1
                    End If
                    If InStr(".mp3|.wav|.mid|.mp4|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & Path)
                        flag0 = 1
                    End If

                    If InStr(".mp3|.wav|.mid|.mp4|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & Path)
                        flag0 = 1
                    End If

                    If InStr(".ppt|.pptx|.pptm|.xls|.xlsx|.xlsm|.doc|.docx|.dotm|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & gIcon_Path & "\settings.gif")
                        Dim OutF As String = gWORK_Path & "\" & gLab.BoxID & ".pdf"
                        Dim Cmd = Path & "," & OutF

                        '  Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(gWORK_Path & "\PDFX.exe ", Path & "," & OutF)

                        flag0 = 1
                    End If

                    If InStr(".exe|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & gIcon_Path & "\X.ico")
                        flag0 = 1
                    End If
                    'Dim answer As DialogResult
                    If flag0 = 0 Then
                        .Source = New Uri("file:///" & gIcon_Path & "\hatena.ico")
                        flag0 = 1
                    End If




            End Select

        End With


        Return WebBox

    End Function
    'Private Function BxViewer_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As SplitContainer
    '    ' "@Viewer|2|BASE|★TYPE|★PATH|★NAME"



    '    Dim Type, Path, Name, Atr As String
    '    Type = STK_PRT(3)
    '    Path = STK_PRT(4)
    '    Name = STK_PRT(5)
    '    Atr = System.IO.Path.GetExtension(Path)
    '    Dim Body As SplitContainer
    '    Body = New SplitContainer
    '    OBJ.Controls.Add(Body)
    '    With Body
    '        .Height = 80
    '        .Dock = DockStyle.Fill
    '        .Orientation = Orientation.Horizontal
    '        .SplitterDistance = 25
    '        .Padding = New Padding(0)
    '        .Margin = New Padding(0)
    '        .FixedPanel = System.Windows.Forms.FixedPanel.Panel1
    '        .Panel1.BackColor = Color.RoyalBlue
    '    End With

    '    Dim Head As SplitContainer
    '    Head = New SplitContainer
    '    Body.Panel1.Controls.Add(Head)
    '    With Head
    '        .Height = 25
    '        .Dock = DockStyle.Fill
    '        .Orientation = 1
    '        .SplitterDistance = 24
    '        .Padding = New Padding(0)
    '        .Margin = New Padding(0)
    '        .FixedPanel = System.Windows.Forms.FixedPanel.Panel1
    '    End With
    '    Dim tex As TextBox
    '    tex = New TextBox
    '    Head.Panel2.Controls.Add(tex)
    '    With tex
    '        .Dock = DockStyle.Fill
    '        .Height = 24
    '        .Multiline = False
    '        .Padding = New Padding(0)
    '        .Margin = New Padding(0)
    '        .Text = Name
    '        .ReadOnly = True
    '        .BackColor = Color.RoyalBlue
    '        .BorderStyle = BorderStyle.None
    '    End With
    '    Dim appIcon As Icon


    '    Dim pic As PictureBox
    '    pic = New PictureBox
    '    Head.Panel1.Controls.Add(pic)
    '    With pic
    '        .Dock = DockStyle.Fill
    '        .BackColor = Color.RoyalBlue
    '        .Height = 24
    '        .Width = 24
    '        .Padding = New Padding(0)
    '        .Margin = New Padding(0)
    '        .SizeMode = PictureBoxSizeMode.Zoom
    '        .WaitOnLoad = False

    '        .Tag = Path

    '        Select Case Type
    '            Case "File"
    '                Try
    '                    appIcon = System.Drawing.Icon.ExtractAssociatedIcon(Path)
    '                    .Image = appIcon.ToBitmap()
    '                Catch ex As Exception
    '                    .BackColor = Color.Red
    '                End Try

    '            Case "Url"
    '                .BackColor = Color.Red
    '        End Select
    '        AddHandler pic.DoubleClick, AddressOf BxPic_DoubleClick

    '    End With



    '    Dim WebBox As WebView2
    '    WebBox = New WebView2
    '    Body.Panel2.Controls.Add(WebBox)

    '    Dim flag0 As Integer
    '    flag0 = 0

    '    With WebBox
    '        .Dock = DockStyle.Fill
    '        Select Case Type
    '            Case "Url"
    '                .Source = New Uri(Path)
    '            Case "File"
    '                If InStr(".txt|.pdf|", Atr & "|") > 0 Then
    '                    .Source = New Uri("file:///" & Path)
    '                    flag0 = 1
    '                End If
    '                If InStr(".ico|.bmp|.jpg|.gif|.png|.exig|.tiff|", Atr & "|") > 0 Then
    '                    .Source = New Uri("file:///" & Path)
    '                    flag0 = 1
    '                End If
    '                If InStr(".mp3|.wav|.mid|.mp4|", Atr & "|") > 0 Then
    '                    .Source = New Uri("file:///" & Path)
    '                    flag0 = 1
    '                End If
    '                Dim answer As DialogResult
    '                If flag0 = 0 Then
    '                    answer = MsgBox("(今だけ仕様)　このファイルを表示しますか？", vbQuestion + vbYesNo, "")
    '                    If answer = DialogResult.Yes Then
    '                        .Source = New Uri("file:///" & Path)
    '                    End If

    '                End If


    '        End Select

    '    End With


    '    Return Body

    'End Function
    Private Function BxHeader_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As Object
        '@Header|2|BASE|★TYPE|★PATH|★NAME"
        Dim Type, Path, Name, Atr As String
        Type = STK_PRT(3)
        Path = STK_PRT(4)
        Name = STK_PRT(5)
        Atr = System.IO.Path.GetExtension(Path)


        Dim appIcon As Icon

        Dim pic As PictureBox
        pic = New PictureBox

        With pic
            .Dock = DockStyle.None
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left
            .Location = New Point(0, 0)
            .BackColor = Color.RoyalBlue
            .Height = 8
            .Width = 16
            .Padding = New Padding(0)
            .Margin = New Padding(0)
            .SizeMode = PictureBoxSizeMode.Zoom
            .WaitOnLoad = False

            .Tag = Path

            Select Case Type
                Case "File"
                    Try
                        appIcon = System.Drawing.Icon.ExtractAssociatedIcon(Path)
                        .Image = appIcon.ToBitmap()
                    Catch ex As Exception
                        .BackColor = Color.Red
                    End Try

                Case "Url"
                    .BackColor = Color.Red
                Case "Text"
                    .BackColor = Color.Blue
                Case "Pen"
                    .BackColor = Color.Green

            End Select
            AddHandler pic.DoubleClick, AddressOf BxPic_DoubleClick
            OBJ.Controls.Add(pic)



        End With

        Return pic

    End Function

    Private Function BxPict_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As PictureBox
        ' "@PICT|0|P.Hed|P.Dock|P.C|P.DAT|"
        Dim Pict As PictureBox
        Pict = New PictureBox
        Dim PxID As Long
        AddHandler Pict.DragEnter, AddressOf BxPict_DragEnter
        AddHandler Pict.DragDrop, AddressOf BxPict_Drop
        Pict.SizeMode = PictureBoxSizeMode.Zoom
        Pict.AllowDrop = True

        OBJ.Controls.Add(Pict)
        Pict.Dock = Val(STK_PRT(3))
        Pict.BackColor = Color.FromName(STK_PRT(4))

        Select Case Mid(STK_PRT(5), 1, 1)
            Case "∇" 'DB登録
                '記録場所を確保
                'パス記憶用  PxID = DBAcs_Insert(Lab, "PBox", "BoxID,Num", Lab.BoxID.ToString & "," & STK_PRT(1).ToString)
                PxID = DBAcs_Insert(Lab, "BDat", "BoxID,Num", Lab.BoxID.ToString & "," & STK_PRT(1).ToString)
                'Blueprint修正 (DB側のみ)
                Box_Re_Pict(Lab, Lab.BoxID, Val(STK_PRT(1)), "∃" & PxID.ToString)

            Case "∃"
                Dim Ds As DataSet
                PxID = Val(Mid(STK_PRT(5), 2))
                'Bx_Get_Pict(Lab, Pict, PxID) 'ファイルパスから
                Bx_Get_Pict2(Lab, Pict, PxID) 'DBから

            Case Else

        End Select
        Pict.Tag = PxID.ToString

        Return Pict

    End Function

    Private Function BxBase_Set(ByRef Lab As Lab_Type, ByRef Obj As Object, ByRef STK_PRT() As String) As Panel
        'BASE生成
        '-------------------
        ' BluePrint &= "@BASE|0|B.Hed|B.Type|B.Dock|B.X|B.Y|B.Z|B.W|B.H|B.C|"

        Dim Base As Panel
        Base = New Panel

        AddHandler Base.MouseEnter, AddressOf Base_MouseEnter
        AddHandler Base.MouseLeave, AddressOf Base_MouseLeave
        AddHandler Base.MouseUp, AddressOf Base_MouseUp
        AddHandler Base.MouseDown, AddressOf Base_MouseDown
        AddHandler Base.MouseMove, AddressOf Base_MouseMove
        AddHandler Base.DoubleClick, AddressOf Base_DoubleClick

        Obj.Controls.Add(Base)
        With Base
            .Dock = Val(STK_PRT(3))
            .Left = Val(STK_PRT(4))
            .Top = Val(STK_PRT(5))
            .Width = Val(STK_PRT(7))
            .Height = Val(STK_PRT(8))
            ' .BackColor = Color.FromName("blue")
            .Padding = New Padding(5)
            .ContextMenuStrip = ContextMenuStrip3
        End With
        Base.Visible = True
        Return Base
    End Function

    Private Function BxLine_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As Panel
        '@LINE|0|BASE|X|Y|Z|W|H|Col|"
        Dim Line As Panel
        Line = New Panel
        If OBJ Is Nothing Then Return Nothing
        AddHandler Line.MouseEnter, AddressOf Line_MouseEnter
        AddHandler Line.MouseLeave, AddressOf Line_MouseLeave
        AddHandler Line.MouseUp, AddressOf Line_MouseUp
        AddHandler Line.MouseDown, AddressOf Line_MouseDown
        AddHandler Line.MouseMove, AddressOf Line_MouseMove
        AddHandler Line.DragOver, AddressOf Line_DragOver
        AddHandler Line.DragDrop, AddressOf Line_DragDrop
        '   AddHandler Line.DoubleClick, AddressOf Base_DoubleClick





        OBJ.Controls.Add(Line)
        With Line
            .Left = 1
            .Top = Val(STK_PRT(4))

            .Width = OBJ.width
            .Height = 6
            .BackColor = Color.FromName(STK_PRT(8))
            .Padding = New Padding(2)
            .AllowDrop = True
            .Tag = "LINE"
            .ContextMenuStrip = ContextMenuStrip3
        End With
        Return Line
        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function
    Private Sub Box_Factory(ByRef Lab As Lab_Type, BoxID As Long, BluePrint As String)
        '設計図に従ってBOXの中身を作る
        '-----------------------------
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim ObjBase As Object
        Dim ObjLine As Object
        Dim ObjTab As Object
        Dim ObjParts As Object
        Dim ObjSpl As SplitContainer
        Dim ObjSpl1 As Object
        Dim ObjSpl2 As Object
        Dim Obj As Object

        STACKS = Split(BluePrint, "@")
        For L = 1 To STACKS.Count - 1
            STK_PRT = Split(STACKS(L), "|")

            Select Case STK_PRT(2)
                Case "DESK"
                    Obj = Lab.DeskAT.View.TabPages(Lab.DeskName)
                Case "LINE"
                    Obj = ObjLine
                    Lab.Box = Obj
                Case "BASE"
                    Obj = ObjBase
                    Lab.Box = Obj
                Case "TAB"
                    Obj = ObjTab
                Case "SPLIT1"
                    Obj = ObjSpl1
                Case "SPLIT2"
                    Obj = ObjSpl2
                Case "NEXT"
                    Obj = ObjParts
                Case Else
                    Obj = ObjParts
            End Select

            Select Case STK_PRT(0)
                Case "BASE"
                    ObjBase = BxBase_Set(Lab, Obj, STK_PRT)
                    ObjBase.name = BoxID.ToString
                    ObjBase.tag = "OnDESK"
                Case "LINE"
                    ObjLine = BxLine_Set(Lab, Obj, STK_PRT)
                    If ObjLine Is Nothing Then Return
                    ObjLine.name = BoxID.ToString
                    ObjLine.tag = "LINE"
                Case "Viewer"
                    ObjParts = BxViewer_Set(Lab, Obj, STK_PRT)

                Case "Header"
                    ObjParts = BxHeader_Set(Lab, Obj, STK_PRT)

                Case "TAB"
                    ObjTab = BxTaB_Set(Lab, Obj, STK_PRT)
                    ObjTab.name = BoxID.ToString & "." & STK_PRT(1).ToString

                Case "SPLIT"
                    ObjSpl = BxSPLIT_Set(Lab, Obj, STK_PRT)
                    ObjSpl1 = ObjSpl.Panel1
                    ObjSpl2 = ObjSpl.Panel2

                Case "PICT"
                    ObjParts = BxPict_Set(Lab, Obj, STK_PRT)

                Case "TEXT"
                    ObjParts = BxText_Set(Lab, Obj, STK_PRT)

                Case "WebView"
                    ObjParts = BxWeb_Set(Lab, Obj, STK_PRT)

                Case "TABp"
                    ObjParts = BxTabP_Set(Lab, Obj, STK_PRT)
                    ObjParts.name = BoxID.ToString & "." & STK_PRT(1).ToString
                Case "Draw"
                    ObjParts = BxDraw_Set(Lab, Obj, STK_PRT)

                Case "CogT"
                    ObjParts = BxCogT_Set(Lab, Obj, STK_PRT)


            End Select

        Next

    End Sub

    Private Sub Box_Makes_OnDESK(ByRef Lab As Lab_Type, BluePrint As String)

        Dim BoxID As Long '製造番号

        BoxID = Box_SaveDB(Lab, BluePrint) 'DBに格納
        Lab.BoxID = BoxID
        Box_Factory(Lab, BoxID, BluePrint) '製造と配置と
        Lab.Box.BringToFront()
    End Sub
    Private Function Box_SaveDB(ByRef Lab As Lab_Type, BluePrint As String) As Long
        Dim BoxID As Long '製造番号
        Dim Data As String
        With Lab
            Data = ""
            Data &= .DeskID.ToString & ","
            Data &= .ProjID.ToString & ","
            Data &= .OwnerID.ToString & ","
            Data &= "'" & BluePrint & "',"
            Data &= "1"
        End With

        BoxID = DBAcs_Insert(gLab, "Box", "DeskID,ProjID,OwnerID,BluePrint,State", Data)


        Return BoxID
    End Function

    Private Function Box_Get_BluePrint(ByRef Lab As Lab_Type, Type As String) As String
        '   設計図生成
        '-----------------------------
        Dim BluePrint As String

        'パラメータ一覧
        'Dock 1上　２下　1左　４→
        '@BASE|0|DESK|Dock|X|Y|Z|W|H|"
        '@PICT|0|BASE|Dock|C|DAT|"
        '@TEXT|0|BASE|Dock|Mul|Bc|Fc|FName|FSize|Text|"
        '@WebView|0|BASE|Dock|URL|"
        '@TAB|0|BASE|Dock|Alignment|ItemSizeX|ItemSizeY|"
        '@TABp|0|TAB|Text|BackColor|"
        '@CogT|0|NEXT|Dock|TEXT|"
        '@Draw|0|NEXT|Dock|DAT|
        '@Split|0|NEXT|Dock|Orientation|SplitterDistance|

        With Lab
            Select Case Type
                Case "Fig"
                    '設計図
                    BluePrint = "Fig"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|100|100"
                    BluePrint &= "@PICT|2|BASE|5|cadetblue|∇"
                    BluePrint &= "@TEXT|3|BASE|2|-1|black|white|Arial|8|∇Fig."

                'Case "File"
                '    '設計図
                '    BluePrint = "FILE"
                '    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|100|20"

                '    BluePrint &= "@PICT|2|BASE|1|darkblue|★"
                '    BluePrint &= "@TEXT|3|BASE|5|0|black|white|Arial|16|∇"

                Case "Viewer"
                    '設計図
                    BluePrint = "Viewer"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|50"
                    BluePrint &= "@Header|2|BASE|★TYPE|★PATH|★NAME"
                    BluePrint &= "@Viewer|3|BASE|★TYPE|★PATH|★NAME"

                Case "Post_it"
                    '設計図
                    BluePrint = "Post_it"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|100"
                    BluePrint &= "@Header|2|BASE|Text|★PATH|Post_it"
                    BluePrint &= "@TEXT|3|BASE|TBLR|0|8|200|80|-1|lemonchiffon|black|Arial|8|∇"

                Case "Pen_Memo"
                    '設計図
                    BluePrint = "Pen_Memo"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|200"
                    BluePrint &= "@Header|2|BASE|Text|★PATH|Pen_Memo"
                    BluePrint &= "@Draw|3|NEXT|5|∇"


                Case "Web_View"
                    '設計図
                    BluePrint = "Web_View"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|400|250"
                    BluePrint &= "@WebView|2|BASE|5|https://www.google.com/"



                Case "G-Docs"
                    '設計図
                    BluePrint = "Web_View"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|400|250"
                    BluePrint &= "@WebView|2|BASE|5|https://docs.google.com/document/d/1MFxmfOIaZKf7KFYoyE7vDC1-77IASuqEb82M9AnP11U/edit#|"



                Case "Web_J-STAGE"
                    '設計図
                    BluePrint = "Web_View"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|400|250"
                    BluePrint &= "@WebView|2|BASE|5|https://www.jstage.jst.go.jp/browse/-char/ja"



                'Case "Post_it"
                '    '設計図
                '    BluePrint = "Post_it"
                '    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|100"
                '    BluePrint &= "@TAB|2|BASE|5|0|6|6"
                '    BluePrint &= "@TABp|3|TAB|■|lemonchiffon"
                '    BluePrint &= "@TEXT|4|NEXT|5|-1|lemonchiffon|black|Arial|8|∇"
                '    BluePrint &= "@TABp|5|TAB|■|lightcyan"
                '    BluePrint &= "@Draw|6|NEXT|5|∇"
                '    BluePrint &= "@TABp|7|TAB|■|lightcyan"
                '    BluePrint &= "@CogT|8|NEXT|5"

                Case "Clothesline"
                    '設計図"@LINE|0|DESK|[BX]|[BY]|[BZ]|W|H|Col"
                    BluePrint = "Clothesline"
                    BluePrint &= "@LINE|0|DESK|1|[BY]|Z|640|5|lavender"



                Case "ID_Card"
                    BluePrint = "ID_Card"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|200|"
                    BluePrint &= "@SPLIT|2|BASE|5|0|80|"
                    BluePrint &= "@TEXT|3|SPLIT2|5|-1|white|black|Arial|8|∇introduce"
                    BluePrint &= "@SPLIT|4|SPLIT1|5|1|60|"
                    BluePrint &= "@PICT|5|SPLIT1|5|white|∇"
                    BluePrint &= "@TEXT|6|SPLIT2|1|0|white|black|Arial|8|∇Tel"
                    BluePrint &= "@TEXT|7|SPLIT2|1|0|white|black|Arial|8|∇Mail"
                    BluePrint &= "@TEXT|8|SPLIT2|1|0|white|black|Arial|8|∇Name"
                    BluePrint &= "@TEXT|9|SPLIT2|1|0|red|white|Arial|8|∇Title"

                Case "ID_Card2"
                    BluePrint = "ID_Card"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|200|"
                    BluePrint &= "@SPLIT|2|BASE|5|0|80|"
                    BluePrint &= "@TEXT|3|SPLIT2|5|-1|white|black|Arial|8|∇introduce"
                    BluePrint &= "@SPLIT|4|SPLIT1|5|1|60|"
                    BluePrint &= "@PICT|5|SPLIT1|5|white|∇"
                    BluePrint &= "@TEXT|6|SPLIT2|1|0|white|black|Arial|8|∇Tel"
                    BluePrint &= "@TEXT|7|SPLIT2|1|0|white|black|Arial|8|∇Mail"
                    BluePrint &= "@TEXT|8|SPLIT2|1|0|white|black|Arial|8|∇Name"
                    BluePrint &= "@TEXT|9|SPLIT2|1|0|red|white|Arial|8|Director of Laboratory"



                Case "Summary"
                    BluePrint = "Summary"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|650|100|"
                    BluePrint &= "@SPLIT|2|BASE|5|1|50|"

                    BluePrint &= "@TEXT|3|SPLIT1|1|0|white|black|Arial|8|Worries"
                    BluePrint &= "@TEXT|4|SPLIT1|1|0|black|white|Arial|8|Hypothesis"
                    BluePrint &= "@TEXT|5|SPLIT1|1|0|black|white|Arial|8|Summary"
                    BluePrint &= "@TEXT|6|SPLIT1|1|0|black|white|Arial|8|TITLE"

                    BluePrint &= "@TEXT|7|SPLIT2|1|0|white|black|Arial|8|∇"
                    BluePrint &= "@TEXT|8|SPLIT2|1|0|white|black|Arial|8|∇"
                    BluePrint &= "@TEXT|9|SPLIT2|1|0|white|black|Arial|8|∇"
                    BluePrint &= "@TEXT|10|SPLIT2|1|0|white|black|Arial|8|∇"


            End Select

        End With
        Return BluePrint
    End Function


    'Private Sub Proj_Drop_Select_ListViewToTreeView(ByRef Proj As Lab_Type, ByRef Source_Item() As ListViewItem, ByRef Target_Node As TreeNode, ByRef Act_type As String)
    '    'ListからTreeにドロップしたときの処理
    '    '----------------
    '    If Source_Item Is Nothing Then Return


    '    With Proj

    '        Select Case Source_Item(0).ListView.Name
    '            Case "ListView1", "ListView3"      'データから
    '                'Select Case Proj_Get_RootType(Target_Node)
    '                '    Case "[Proj]"
    '                '        Proj_Drop_DicToRepo(Proj, Source_Item, Target_Node, Act_type)
    '                '    Case "[Leaf]", "[Branch]", "[Tree]"
    '                '        Proj_Drop_DicToDicbox(Proj, Source_Item, Target_Node, Act_type)
    '                'End Select

    '            Case "ListView2"        'ATRフィールドから
    '                'Select Case Proj_Get_RootType(Target_Node)
    '                '    Case "[Proj]" ' to [Proj] 
    '                '        Proj_Drop_AtrToRepo(Proj, Source_Item, Target_Node, Act_type)
    '                '    Case "[Leaf]", "[Branch]", "[Tree]"
    '                '        Proj_Drop_AtrToDicbox(Proj, Source_Item, Target_Node, Act_type)
    '                'End Select

    '        End Select





    '    End With


    'End Sub


    'Private Function Node_Get_Child(ByRef P_Node As TreeNode) As TreeNode()
    '    'ノード配下のノードをListup 、 For Each Node 内部で処理しないのは、内部でノード状態を変化させると副作用があるから
    '    '----------------------
    '    Dim C_Node() As TreeNode
    '    C_Node = Nothing
    '    Dim C_max As Integer
    '    C_max = 0

    '    For Each Node In P_Node.Nodes
    '        ReDim Preserve C_Node(C_max)
    '        C_Node(C_max) = Node
    '        C_max += 1
    '    Next

    '    Return C_Node

    'End Function





    Private Sub Dir_Check(ByVal Path As String)
        If System.IO.Directory.Exists(Path) Then Return

        System.IO.Directory.CreateDirectory(Path)

    End Sub
    Private Sub Lab_Set_Env(ByRef Lab As Lab_Type)
        '環境変数セット
        '------------------------------------------
        Log_Show(1, "環境変数セット")


        'グローバル変数セット
        '--------------------
        gApp_Path = IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)    'EXE位置
        gWORK_Path = gApp_Path & "\WORK" : Dir_Check(gWORK_Path)   '各種WORKデータへのパス  
        gENV_Path = gApp_Path & "\ENV" : Dir_Check(gENV_Path)    '各種設定データへのパス  
        gLOG_Path = gApp_Path & "\LOG" : Dir_Check(gLOG_Path)  'ログ出力先のパス  
        gIcon_Path = gApp_Path & "\ENV\ICONS" : Dir_Check(gIcon_Path)  'ICON格納フォルダ



        ''ネットワークセット
        ''--------------------
        'Dim IPHstEnt As IPHostEntry
        'IPHstEnt = Dns.Resolve("192.168.0.32")
        'Log_Show(1, "このPCは" & System.Net.Dns.GetHostName() & Chr(13) & "通信相手は" & IPHstEnt.HostName)

        'Ping(IPHstEnt.HostName)

        ''DBセット
        ''--------------------
        ''SQLサーバー
        ''------------
        With gLab
            Select Case .Db_Type
                Case "SQLite"
                    gDB_Path = System.Windows.Forms.Application.StartupPath & "\ENV\GIANT.sqlite3"
                    DB_Check_Exist_DB(.Db_Type, .DbPath) 'DB 存在チェック&ないなら作成


                Case "MSSQLW"  ''WINDOWS認証 
                    Dim IPHstEnt As IPHostEntry
                    IPHstEnt = Dns.Resolve("192.168.0.32")
                    gDB_Path = "Data Source=" & System.Net.Dns.GetHostName() & "\" & "SQLEXPRESS"
                    gDB_Path = "Data Source= " & IPHstEnt.HostName & "\SQLEXPRESS"
                    gDB_Path &= ";Integrated Security=SSPI"            '192.168.0.32

                Case "MSSQL" ''SQL認証
                    gDB_Path = "Data Source= 192.168.0.32,1433\SQLEXPRESS"
                    gDB_Path &= ";User ID = 'sa';Password = 'MONAKA123'"

            End Select
        End With

        'DB設定
        '------------------
        With Lab
            .DbPath = gDB_Path
            .Db_Type = gLab.Db_Type
            .DBACS_Type = New Dictionary(Of String, String)
            .DBACS_Field = New Dictionary(Of String, String)
        End With

        DBAcs_Init(Lab, "Person") 'Person用DB
        DBAcs_Init(Lab, "Proj") 'Proj用DB
        DBAcs_Init(Lab, "Role") '
        DBAcs_Init(Lab, "Desk") 'Desk用DB
        DBAcs_Init(Lab, "Box") 'Box用DB
        DBAcs_Init(Lab, "Element") 'Element用DB

        DBAcs_Init(Lab, "Tbox") 'テキスト用DB
        DBAcs_Init(Lab, "Pbox") '画像用DB
        DBAcs_Init(Lab, "BDat") 'バイトデータ用DB
        'DBAcs_Init(Lab, "Draw") 'バイトデータ用DB

        'イメージ設定 (アイコンフォルダの画像イメージを利用にかかわらずすべて登録
        '--------------------
        ' Log_Show_Splash("イメージ設定")
        gLab.Image_List = Image_Init(gIcon_Path)
        PictureBox1.AllowDrop = True
        PictureBox1.AutoSize = True

        'デコレーション
        '-----------------------------------
        'gDeco.BackColor = "DarkRed"
        ' gDeco.ForeColor = "Black"

        gDeco.BackColor = "Black"
        gDeco.ForeColor = "White"


        'コントロール設定
        '--------------------

        Controls_Init()

        'AT登録

        With gLab.ProjAT.View
            .ListView = ListView1
            Columns_Set_Env(.Columns, "[Proj]")
        End With

        With gLab.ElementAT.View
            .ListView = ListView2
            Columns_Set_Env(.Columns, "[Element]")
        End With

        With gLab.StackAT.View
            .ListView = ListView3
            Columns_Set_Env(.Columns, "[Stack]")
        End With


        With gLab.DeskAT
            .View = TabControl3


            'ReDim .Desk(2) ' 0-データパーツ用 ,1--Proj用
            'ReDim .Person(0)
        End With
        ' RepoList_Init(gLab.View) '一覧設定


        'コンテキストメニューの設定
        '-----------------------------------
        Menu_Set(ContextMenuStrip1, "プロジェクトの新規作成/プロジェクト名変更/プロジェクトの削除", "NEW/RENAME/DEL")
        Menu_Set(ContextMenuStrip2, "作業デスクの追加/作業デスクの名前変更/作業デスクの削除", "NEW/RENAME/DEL")
        Menu_Set(ContextMenuStrip3, "スタックのコピー/スタックの貼付け/スタックの削除", "COPY/PASTE/DEL")
        Menu_Set(ContextMenuStrip4, "資料の削除", "DEL")



        gLab.ProjAT.View.ListView.ContextMenuStrip = ContextMenuStrip1
        TabControl3.ContextMenuStrip = ContextMenuStrip2
        ListView2.ContextMenuStrip = ContextMenuStrip4


    End Sub
    'Private Sub Box_Set_Env()
    '    'らっぱーの設定
    '    '----------
    '    Dim gBox(10) As Box_Type

    '    'With gBox(0)
    '    '    .Name = "googleドキュメント"
    '    '    .Panel.Controls.Add(.WebB)
    '    '    ' .Panel.Margin=
    '    '    .WebB.Dock = DockStyle.Fill
    '    '    .Panel.Width = 150
    '    '    .Panel.Height = 50

    '    'End With

    'End Sub
    Private Sub Box_Set_OnDESK(ByRef lab As Lab_Type, DeskID As Long)
        'BOXをDESKに再配置（データを再度SAVEしないように注意）
        '-----------------------------
        ' .DBACS_Field.Add(Key, "ID, PageID, PageName, Type, Style, X, Y, Z, W, H")
        Dim DS As DataSet
        Dim cmd As String
        cmd = "DeskID = " & DeskID.ToString
        cmd &= " AND State > 0"

        DS = DBAcs_Get_DataS(lab, "Box", "ID,BluePrint,State", cmd)
        If DS.Tables(0).Rows.Count = 0 Then Return
        With DS.Tables(0)
            For L = 0 To DS.Tables(0).Rows.Count - 1
                lab.BoxID = Val(.Rows(L).Item(0))
                Box_Factory(lab, lab.BoxID, Trim(.Rows(L).Item(1)))
                If .Rows(L).Item(2) > 1 Then lab.Box.Tag = "OnLINE." & .Rows(L).Item(2)
            Next
        End With

    End Sub

    Private Sub Box_Show_List()
        'スタックをLISTVIEWに表示
        '-----------------------------

        With gLab.StackAT.View
            Columns_Set_Listview(.ListView, .Columns)
        End With

        Box_Set_Item(gLab, "Post_it", "Post_it", "Post_it")
        Box_Set_Item(gLab, "Pen_Memo", "Pen_Memo", "Pen_Memo")
        Box_Set_Item(gLab, "物干し竿", "Clothesline", "Clothesline")

        'Box_Set_Item(gLab, "Fig", "Fig", "Fig")
        'Box_Set_Item(gLab, "Webブラウザ", "Web_View", "web")
        'Box_Set_Item(gLab, "スタンプ（公開）", "Stamp(OPEN)", "Stamp_OPEN")
        'Box_Set_Item(gLab, "スタンプ（非公開）", "Stamp(CLOSE)", "Stamp_CLOSE")
        ''  Box_Set_Item(gLab, "スタンプ（welcome）", "Stamp(welcome)", "Stamp_Welcome")
        'Box_Set_Item(gLab, "ID_Card", "ID_Card", "ID_Card")
        'Box_Set_Item(gLab, "サマリー", "Summary", "Summary")
        'Box_Set_Item(gLab, "Webブラウザ(Googleドキュメント)", "G-Docs", "G-Docs")
        'Box_Set_Item(gLab, "Webブラウザ(J-Stage)", "Web_J-STAGE", "J-STAGE_logo")
        'Box_Set_Item(gLab, "クリッパー(製作中)", "ハサミ", "ハサミ")
        'Box_Set_Item(gLab, "引用マーカ(製作中)", "マーカー", "マーカー")
        'Box_Set_Item(gLab, "スパイダー(製作中)", "SPIDER_J-STAGE", "スパイダー")
        'Box_Set_Item(gLab, "フォローJ(製作中)", "Web_J-STAGE", "ターゲット")


    End Sub

    Private Sub Box_Set_Item(ByRef Lab As Lab_Type, Name As String, Key As String, IconKey As String)
        'ElementデータをLISTVIEWに表示
        '-----------------------------
        Dim Itm As ListViewItem

        With Lab.StackAT.View.ListView

            Itm = .Items.Add(Name)
            Itm.ImageIndex = Image_Get_ID(Lab.Image_List, IconKey)
            ' Itm.ImageIndex = Image_Get_ID(Lab.Image_List, IconKey)
            Itm.Tag = Key

        End With

    End Sub


    Private Sub Controls_Init()

        'バックから―
        With gDeco
            '
            'SplitContainer1.Panel1.BackColor = Color.FromName(.BackColor)
            'SplitContainer1.Panel2.BackColor = Color.FromName(.BackColor)

            SplitContainer5.Panel1.BackColor = Color.FromName(.BackColor)

            ListView1.BackColor = Color.FromName(.BackColor)
            ListView2.BackColor = Color.FromName(.BackColor)
            ListView3.BackColor = Color.FromName(.BackColor)

            'SplitContainer1.Panel1.ForeColor = Color.FromName(.ForeColor)
            'SplitContainer1.Panel2.ForeColor = Color.FromName(.ForeColor)



            SplitContainer5.Panel1.ForeColor = Color.FromName(.ForeColor)


            ListView1.ForeColor = Color.FromName(.ForeColor)
            ListView2.ForeColor = Color.FromName(.ForeColor)
            ListView3.ForeColor = Color.FromName(.ForeColor)


            TextBox1.ForeColor = Color.FromName(.ForeColor)
            'TextBox2.ForeColor = Color.FromName(.ForeColor)
            'TextBox3.ForeColor = Color.FromName(.ForeColor)
            'TextBox4.ForeColor = Color.FromName(.ForeColor)
            'TextBox5.ForeColor = Color.FromName(.ForeColor)
            TextBox6.ForeColor = Color.FromName(.ForeColor)

            TextBox1.BackColor = Color.FromName(.BackColor)
            'TextBox2.BackColor = Color.FromName(.BackColor)
            'TextBox3.BackColor = Color.FromName(.BackColor)
            'TextBox4.BackColor = Color.FromName(.BackColor)
            'TextBox5.BackColor = Color.FromName(.BackColor)
            TextBox6.BackColor = Color.FromName(.BackColor)


        End With




        'ツールチップ
        With ToolTip1
            .ToolTipTitle = ""
            '.ToolTipIcon = "" ' ToolTipIcon.Info
            '.SetToolTip(PictureBox2, "研究画面を開きます。")
            ''.SetToolTip(PictureBox3, "研究を探します。")
            '.SetToolTip(PictureBox3, "新しい研究を始めます。")
            '.SetToolTip(PictureBox4, "私について")

            '.SetToolTip(PictureBox3, "研究ノートを開きます。")

        End With

        'LISTVIEW

        With ListView1
            .View = View.Details
            '.LargeImageList = gImg_List.Image
            '.SmallImageList = gImg_List.Image
            '.StateImageList = gImg_List.Image

        End With
        With ListView2 'スタック用
            '.View = View.LargeIcon
            .View = View.Details
            '.LabelWrap = True
            ' .LargeImageList = gImg_List.Image
            .SmallImageList = gImg_List.Image
            ' .StateImageList = gImg_List.Image
        End With
        With ListView3 '資料用
            '   .View = View.LargeIcon
            .View = View.Details
            '  .LabelWrap = True
            '.LargeImageList = gImg_List.Image
            .SmallImageList = gImg_List.Image
            '.StateImageList = gImg_List.Image

        End With





    End Sub

    Function Owner_Get_ID() As Integer
        'オーナー番号取得
        Dim tmpID As Integer
        Dim sr As StreamReader
        '     Dim tmpDir As String = System.IO.Path.GetTempPath()
        Dim tmpDir As String = System.Windows.Forms.Application.StartupPath() & "\ENV"
        If System.IO.File.Exists(tmpDir & "\ginger.id") = False Then
            Return 0
        Else
            sr = New StreamReader(tmpDir & "\ginger.id")
            tmpID = Val(sr.ReadToEnd)
            sr.Close()
        End If

        Return tmpID
    End Function
    Sub Owner_Save_ID(ByVal ID As Integer)
        'OWNER ID をファイルに保存
        ' Dim tmpDir As String = System.IO.Path.GetTempPath()
        Dim tmpDir As String = System.Windows.Forms.Application.StartupPath() & "\ENV"
        Dim sw As New System.IO.StreamWriter(tmpDir & "\ginger.id", False, System.Text.Encoding.GetEncoding("shift_jis"))

        sw.WriteLine(ID.ToString)
        sw.Close()


    End Sub

    Sub Owner_Show_Data(Lab As Lab_Type)
        'OWNER Data を表示
        '-------------------

        'With Lab.Person(0)
        '    TextBox4.Text = .Name
        '    TextBox2.Text = .Tel
        '    TextBox3.Text = .Mail
        '    TextBox5.Text = .Prof
        '    If System.IO.File.Exists(.Image) = True Then PictureBox1.ImageLocation = .Image
        'End With

    End Sub

    Sub Ping(tmpS As String)
        'Pingオブジェクトの作成
        Dim p As New System.Net.NetworkInformation.Ping()
        '"www.yahoo.com"にPingを送信する
        Dim reply As System.Net.NetworkInformation.PingReply = p.Send(tmpS)

        '結果を取得
        If reply.Status = System.Net.NetworkInformation.IPStatus.Success Then
            'Console.WriteLine("Reply from {0}:bytes={1} time={2}ms TTL={3}",
            '    reply.Address, reply.Buffer.Length,
            '    reply.RoundtripTime, reply.Options.Ttl)
            Log_Show(2, "Ping送信OK" & tmpS & reply.Status)
        Else
            Log_Show(2, "Ping送信に失敗。" & tmpS & reply.Status)
        End If

        p.Dispose()
    End Sub


    Private Sub ToolStripButton_Set_Env(ByRef btn As ToolStripButton, ByRef key As String)

        With btn
            '.ImageKey = Image_Get_ID(key, gImg_List)
            .ImageKey = key
            .ToolTipText = key
        End With

    End Sub
    Private Sub Button_Set_Env(ByRef btn As Button, ByRef key As String)
        'ボタンに絵を出す

        With btn
            .ImageList = gImg_List.Image
            .ImageKey = key
            '.ToolTipText = key
        End With

    End Sub

    Private Sub Proj_Make_New(ByRef Lab As Lab_Type)
        'Projの追加 (DB修正)
        '-------------------------------------------

        With Lab
            'DBに仮登録
            .ProjID = DBAcs_Insert(Lab, "Proj", "OwnerID,Owner,Name,State", gX_ownerID & ",'" & gX_owner & "'," & "'新プロジェクト',1")

            Log_Show(1, "ProjID=" & .ProjID.ToString)
            'ページ作成

            Dim KEY(3) As String

            KEY(0) = Desk_Make_New(Lab, "Summary", "Summary", True)

            'スタック追加

            'With Lab
            '    .BluePrint = Box_Get_BluePrint(gLab, "Summary") '設計
            '    .BluePrint = Replace(.BluePrint, "[BX]", "1")
            '    .BluePrint = Replace(.BluePrint, "[BY]", "1")
            '    Box_Makes_OnDESK(Lab, .BluePrint)
            'End With

            'With Lab
            '    .BluePrint = Box_Get_BluePrint(gLab, "ID_Card2") '設計
            '    .BluePrint = Replace(.BluePrint, "[BX]", "1")
            '    .BluePrint = Replace(.BluePrint, "[BY]", "120")
            '    Box_Makes_OnDESK(Lab, .BluePrint)
            'End With


            KEY(1) = Desk_Make_New(Lab, ”Desk”, "Desk", False)


            'LISTVIEW追加
            Dim DS As DataSet
            DS = DBAcs_Get_DataS(Lab, "Proj", .ProjAT.View.Columns.DBKeys, "OwnerID = " & gX_ownerID)
            Proj_Show_List(.ProjAT.View.ListView, .ProjAT.View.Columns, DS, True)
        End With
    End Sub



    Private Function ProjDB_Insert(ByRef Lab As Lab_Type, Name As String, ByRef Key As String, Val As String) As Long
        'DB Node　を追加
        '--------------------------
        Dim tmpR As String

        Dim DBCon As SqlConnection
        Dim DBCom As SqlCommand
        Dim TmpCmd As String


        With Lab
            Select Case .Db_Type
                Case "SQLite"
                Case "MSSQL"
                    DBCon = New SqlConnection(Lab.DbPath & ";Initial Catalog =" & .Db_Type & Name)
                    TmpCmd = "INSERT INTO " & "TBL_" & Name
                    TmpCmd &= Key & "VALUES" & Val
                    TmpCmd &= "; SELECT SCOPE_IDENTITY();"
            End Select

        End With

        DBCom = New SqlCommand(TmpCmd, DBCon)
        DBCom.CommandTimeout = 0

        DBCom.Connection.Open()

        tmpR = -1
        For I As Integer = 1 To 50
            Try
                tmpR = DBCom.ExecuteScalar()

                Exit For
            Catch ex As Exception
                Log_Show(1, "DB登録エラー" & I & ex.Message)
                For J As Integer = 1 To 100
                    System.Windows.Forms.Application.DoEvents()
                Next
            End Try
        Next

        DBCom.Connection.Close()
        DBCon.Close()
        DBCon.Dispose()
        Return tmpR

    End Function
    'Private Function ProjDB_Insert_OLD(ByRef Lab As Lab_Type, Name As String, ByRef Proj As Proj_Type) As Integer
    '    'DB Node　を追加
    '    '--------------------------
    '    Dim tmpR As String

    '    Dim DBCon As SqlConnection
    '    Dim DBCom As SqlCommand
    '    Dim TmpCmd As String

    '    DBCon = New SqlConnection(Lab.DbPath & ";Initial Catalog =" & Lab.DBname & Name)
    '    With Proj

    '        TmpCmd = "INSERT INTO " & "TBL_" & Name
    '        TmpCmd &= " ( Name, Prof, Note, OwnerID, Owner, PID) VALUES ("

    '        TmpCmd &= "'" & .Name & "',"

    '        TmpCmd &= "'" & .OwnerID & "',"
    '        TmpCmd &= "'" & .Owner & "',"

    '        TmpCmd &= ")"
    '        TmpCmd &= "; SELECT SCOPE_IDENTITY();"
    '    End With

    '    DBCom = New SqlCommand(TmpCmd, DBCon)
    '    DBCom.CommandTimeout = 0

    '    DBCom.Connection.Open()

    '    tmpR = -1
    '    For I As Integer = 1 To 50
    '        Try
    '            tmpR = DBCom.ExecuteScalar()

    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(1, "DB登録エラー" & I & ex.Message)
    '            For J As Integer = 1 To 100
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next

    '    DBCom.Connection.Close()
    '    DBCon.Close()
    '    DBCon.Dispose()
    '    Return tmpR

    'End Function



    Private Sub Menu_Act_Dicbox(ByRef Proj As Lab_Type, ByRef Cmd() As String)
        'メニュー選択時のアクションまとめ TREE
        '-----------------------------------------
        Dim DeskID As Integer
        DeskID = 0
        Dim Act_Effect As String
        Dim N_Name, O_Name As String
        Dim O_NID As Integer
        Dim TmpV As String

        Dim D_Node() As TreeNode
        D_Node = Nothing
        Dim N_Type, N_State As String
        N_Type = "" : N_State = ""


        N_Name = ""

        'Dim History As History_Type
        'History = Proj.Desk(DeskID).History



    End Sub


    Private Sub Menu_Act_DicList(ByRef Proj As Lab_Type, ByRef LV As ListView, ByRef Cmd() As String)
        'メニュー選択時のアクションまとめ DIC
        '-----------------------------------------
        'Dim DeskID As Integer
        'DeskID = 0

        'Dim Flag_RID, Flag_DID As Integer
        'Dim Item() As ListViewItem
        'Dim Cb_data As String

        'With LV
        '    '対象ノード非選択
        '    If .SelectedItems.Count = 0 Then Return
        '    '対象ノード
        '    ReDim Item(.SelectedItems.Count - 1)
        '    For L = 0 To .SelectedItems.Count - 1
        '        Item(L) = .SelectedItems(L)
        '        Progres_Show(L + 1, .SelectedItems.Count, False)
        '    Next

        '    Select Case Cmd(1)

        '        Case "コピー"
        '            gBff_Act_Type = "COPY_ITEM"
        '            gBff_Node = Nothing
        '            gBff_Item = Item.Clone


        '            'クリップボードにコピー
        '            Clipboard.Clear()   'クリップボードを初期化
        '            Clipboard.SetText(ListView_Get_TextData(LV))



        '        Case "切り取り"
        '            gBff_Act_Type = "MOVE_ITEM"
        '            gBff_Node = Nothing
        '            gBff_Item = Item.Clone

        '        Case "貼り付け"
        '            '■何もしない
        '            Log_Show(2, "指定のオブジェクトは貼り付けできません")

        '        Case "削除"
        '            '''Histroy_Start(gLab.Desk(DeskID).History)
        '            'For L = Item.Length - 1 To 0 Step -1
        '            '    Proj_Act_Dat_Del_Item(Proj, Item(L), True)
        '            '    Item(L).Remove()
        '            'Next
        '            '''Histroy_End(gLab.Desk(DeskID))

        '        Case "新規"
        '            '----------------------------------------------------
        '            '■何もしない
        '            Log_Show(2, "新規/変更はできません")

        '        Case "変更"
        '            '----------------------------------------------------
        '            '■何もしない
        '            Log_Show(2, "新規/変更はできません")
        '    End Select

        'End With

        ''現在開いているものと一致するなら再描画
        ''If Flag_RID <> 0 Then Proj_Show_List(Proj, gDic_TblID)
        ''If Flag_DID <> 0 Then Proj_Show_Atr(Proj, gAtr_NID, gAtr_LID)



    End Sub

    'Private Function ListView_Get_TextData(ByRef LV As ListView) As String
    '    '選択されているListViewのアイテムをテキスト化する
    '    '-----------------------------------------
    '    Dim Lv_data As String
    '    Dim Item() As ListViewItem
    '    Lv_data = ""

    '    With LV
    '        '対象ノード非選択
    '        If .SelectedItems.Count = 0 Then Return Lv_data

    '        '対象ノード
    '        ReDim Item(.SelectedItems.Count - 1)
    '        For L = 0 To .SelectedItems.Count - 1
    '            Item(L) = .SelectedItems(L)

    '            For K = 0 To .SelectedItems(L).SubItems.Count - 1
    '                Lv_data &= .SelectedItems(L).SubItems(K).Text & Chr(9)
    '            Next
    '            Progres_Show(L + 1, .SelectedItems.Count, False)
    '            Lv_data &= Chr(13) & Chr(10)
    '        Next

    '    End With
    '    Return Lv_data
    'End Function

    Private Sub Menu_Act_DicAtr(ByRef Proj As Lab_Type, ByRef Cmd() As String)
        'メニュー選択時のアクションまとめ ATR
        '-----------------------------------------
        '画面上の操作のみで、HISTORYの対象外

        'Dim DeskID As Integer
        'DeskID = 1


        'Dim Item() As ListViewItem


        'With ListView2
        '    '対象ノード非選択
        '    If .SelectedItems.Count = 0 Then Return

        '    '対象ノード
        '    ReDim Item(.SelectedItems.Count - 1)
        '    For L = 0 To .SelectedItems.Count - 1
        '        Item(L) = .SelectedItems(L)
        '    Next


        '    Select Case Cmd(1)

        '        Case "コピー"
        '            gBff_Act_Type = "COPY_ATR"
        '            gBff_Node = Nothing
        '            gBff_Item = Item.Clone

        '        Case "切り取り"
        '            gBff_Act_Type = "MOVE_ATR"
        '            gBff_Node = Nothing
        '            gBff_Item = Item.Clone

        '            ListView_Del(ListView2)

        '        Case "貼り付け"

        '            Select Case gBff_Act_Type
        '                Case "COPY_NODE", "MOVE_NODE" '同じあつかい　元ノードは消さない
        '                    '----------------------------------------------------

        '                    .Items.Insert(Item(0).Index, Item(0).Clone)
        '                    Dim Itm As String
        '                    '    Itm = Proj_ATR_Make_Item_byNodeName(gBff_Node.Text)
        '                   ' Proj_Atr_Rep_Item(.SelectedItems(0), Itm)

        '                Case "COPY_ITEM", "MOVE_ITEM"
        '                    '----------------------------------------------------
        '                    '■何もしない
        '                    Log_Show(2, "指定のオブジェクトは貼り付けできません")

        '                Case "MOVE_ATR", "COPY_ATR" '同じあつかい　元ノードは消されてる
        '                    For K = 0 To gBff_Item.Length - 1
        '                        .Items.Insert(Item(0).Index, gBff_Item(K).Clone)
        '                    Next

        '            End Select

        '        Case "削除"
        '            ListView_Del(ListView2)

        '        Case "新規"
        '            ''If Cmd.Length <3 Then Return
        '            ''Select Case Cmd(2)
        '            ''    Case "S", "C", "G"
        '            ''        Proj_Atr_New_Item(ListView2, Cmd(2))
        '            ''End Select
        '            'Proj_Atr_New_Item(ListView2, "S")

        '        Case "変更"
        '            'If Cmd.Length < 3 Then Return
        '            'Dim Itm As String
        '            'Itm = Proj_ATR_Make_Item_byMenuCmd(Cmd)
        '            'Proj_Atr_Rep_Item(.SelectedItems(0), Itm)
        '    End Select


        'End With

    End Sub




    Private Sub Menu_Set(ByRef Menu As ContextMenuStrip, ByRef Tx As String, ByRef Kx As String)

        'メニューに追加
        '-----------------------------------------
        Dim dx() As String
        Dim tag() As String

        If Tx = "" Then Return
        dx = Split(Tx, "/")
        tag = Split(Kx, "/")
        Dim Itm(dx.Length) As ToolStripMenuItem
        For K = 0 To dx.Length - 1
            If dx(K) <> "" Then
                Itm(K) = New ToolStripMenuItem
                Itm(K).Text = dx(K)
                Itm(K).Tag = tag(K)
                Menu.Items.Add(Itm(K))
            End If
        Next

        Return

    End Sub








    Private Sub Owner_Set_Env(ByRef Lab As Lab_Type, ID As Integer)

        Dim Ds As DataSet
        Dim Cnd As String
        Cnd = "ID = " & ID.ToString


        'With Lab.Person(0)
        '    DBAcs_Init2(.Db, "OwnerDB", "OWNER", "OWNER")
        '    Ds = DBAcs_Get_DataS(.Db, "Name,Mail,Tel,Prof,Image", Cnd)
        '    If Ds.Tables(0).Rows.Count > 0 Then
        '        .ID = ID
        '        .Name = Trim(Ds.Tables(0).Rows(0).Item(0))
        '        .Mail = Trim(Ds.Tables(0).Rows(0).Item(1))
        '        .Tel = Trim(Ds.Tables(0).Rows(0).Item(2))
        '        .Prof = Trim(Ds.Tables(0).Rows(0).Item(3))
        '        .Image = Trim(Ds.Tables(0).Rows(0).Item(4))
        '        Lab.Owner = .Name
        '        Lab.OwnerID = ID
        '    End If
        'End With




    End Sub



    Private Function Person_Set_NewData(ByRef Lab As Lab_Type) As Long
        'OWNER プロファイル登録
        '-------------------
        'Dim key, tmpD As String
        'Dim Person As Person_Type
        'With Person
        '    .Name = "Unknown"
        '    .Tel = "Tel"
        '    .Mail = "Mail"
        '    .Prof = "Prof"
        '    .Address = "Address"
        'End With
        'key = "(Name)"()
        ''DB登録
        'Proj.ProjID = ProjDB_Insert(Lab, "Proj", "(OwnerID,Name)", "(" & Lab.OwnerID.ToString & ",'新しいプロジェクト')")
        ''With Lab.Person(0)
        ''    DBAcs_Init2(.Db, "OwnerDB", "OWNER", "OWNER")
        ''    tmpD = ""
        ''    .ID = Owner_Get_ID()
        ''    If .ID = 0 Then

        ''        key = Mid(.Db.TBL_Field, Len("ID,") + 1)
        ''        tmpD &= "'" & .Name & "',"
        ''        tmpD &= "'" & .Mail & "',"
        ''        tmpD &= "'" & .Tel & "',"
        ''        tmpD &= "'" & .Prof & "',"
        ''        tmpD &= "'" & .Image & "'"
        ''        .ID = DBAcs_Insert(.Db, .Db.TBL_Name, key, tmpD) 'DB登録
        ''        Lab.OwnerID = .ID
        ''        Owner_Save_ID(.ID) 'ファイル保存
        ''    Else
        ''        tmpD &= "Name = '" & .Name & "',"
        ''        tmpD &= "Mail = '" & .Mail & "',"
        ''        tmpD &= "Tel = '" & .Tel & "',"
        ''        tmpD &= "Prof = '" & .Prof & "',"
        ''        tmpD &= "Image = '" & .Image & "'"
        ''        DBAcs_Update_Data_ByKey(.Db, .Db.TBL_Name, .ID, tmpD) 'DB修正
        ''        Lab.OwnerID = .ID

        ''    End If



        ''End With

        ''Owner_Show_Data(Lab)


    End Function

    Private Sub Desk_Set_NewData(ByRef Lab As Lab_Type)
        'OWNER プロファイル登録
        '-------------------
        Dim key, tmpD As String

        'With Lab.Desk
        '    DBAcs_Init2(.DBACS, "OwnerDB", "OWNER", "OWNER")
        '    tmpD = ""
        '    .ID = Owner_Get_ID()
        '    If .ID = 0 Then

        '        key = Mid(.Db.TBL_Field, Len("ID,") + 1)
        '        tmpD &= "'" & .Name & "',"
        '        tmpD &= "'" & .Mail & "',"
        '        tmpD &= "'" & .Tel & "',"
        '        tmpD &= "'" & .Prof & "',"
        '        tmpD &= "'" & .Image & "'"
        '        .ID = DBAcs_Insert(.Db, .Db.TBL_Name, key, tmpD) 'DB登録
        '        Lab.OwnerID = .ID
        '        Owner_Save_ID(.ID) 'ファイル保存
        '    Else
        '        tmpD &= "Name = '" & .Name & "',"
        '        tmpD &= "Mail = '" & .Mail & "',"
        '        tmpD &= "Tel = '" & .Tel & "',"
        '        tmpD &= "Prof = '" & .Prof & "',"
        '        tmpD &= "Image = '" & .Image & "'"
        '        DBAcs_Update_Data_ByKey(.Db, .Db.TBL_Name, .ID, tmpD) 'DB修正
        '        Lab.OwnerID = .ID

        '    End If



        'End With

        Owner_Show_Data(Lab)


    End Sub

    Private Sub DBAcs_Init(ByRef Lab As Lab_Type, ByRef Key As String)
        'DBアクセスの為の型とエントリ名を設定

        With Lab
            Select Case Key

                Case "Proj"
                    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT,State INTEGER ,OwnerID int,Owner  TEXT ,Name  TEXT ,Title  TEXT,Purpose text,Summary text,Hypothesis Text,Problem text)")
                    .DBACS_Field.Add(Key, "ID,State,OwnerID,Owner,Name,Title,Purpose,Summary,Hypothesis,Problem")


                Case "Person"
                    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT ,  Name  TEXT ,Mail  TEXT,Tel   TEXT , Prof text, Image  TEXT ,State int)")
                    .DBACS_Field.Add(Key, "ID,Name,Mail,Tel,Prof,Image ,State")


                Case "Desk"
                    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,OwnerID int,ProjID int,Page int,Type  TEXT ,Title  TEXT ,State int,ShiftX int,ShiftY int,ZM int)")
                    .DBACS_Field.Add(Key, "ID,OwnerID,ProjID,Page,Type,Title ,State,ShiftX,ShiftY,ZM")

                Case "Box"
                    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,DeskID int,ProjID int,OwnerID int,BluePrint text ,State int)")
                    .DBACS_Field.Add(Key, "ID,DeskID,ProjID,OwnerID,BluePrint ,State")


                Case "Tbox"
                    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,BoxID int,Num int,Text Text)")
                    .DBACS_Field.Add(Key, "ID,BoxID,Num,Text ,State")

                'Case "Draw"
                '    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,BoxID int,Num int,Text Text ,State int,Path string)")
                '    .DBACS_Field.Add(Key, "ID,BoxID,Num,Text , State,Path")

                Case "Pbox"
                    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,BoxID int,Num int,Pict TEXT ,State int)")
                    .DBACS_Field.Add(Key, "ID,BoxID,Num,Pict ,State")

                Case "BDat"
                    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,BoxID int,Num int,Dat BLOB ,State int,Type String,Path string)")
                    .DBACS_Field.Add(Key, "ID,BoxID,Num,Dat ,State,Type,Path")

                'Key= "Webbox"
                '    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,BaseID int,Type  TEXT ,Style int,Text Text)"
                '    .DBACS_Field.Add(Key,"ID,BaseID,Type,Path,Style,Text"
                'DB_Check(Lab, key)

                Case "Element"
                    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,ProjID int,Type  TEXT ,Name  TEXT,Path TEXT ,State int)")
                    .DBACS_Field.Add(Key, "ID,ProjID,Type,Name,Path ,State")

                Case "Role"
                    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,LID INTEGER ,PID int,Role  TEXT ,State INTEGER )")
                    .DBACS_Field.Add(Key, "ID,LID,PID,Role ,State")

            End Select

        End With
        'SQLサーバ
        'With Lab
        '    Select Case Key

        '        Case "Proj"
        '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1),State int ,OwnerID int,Owner TEXT(100) ,Name nvarchar(200) ,Title nvarchar(200),Purpose text,Summary text,Hypothesis Text,Problem text)")
        '            .DBACS_Field.Add(Key, "ID,State,OwnerID,Owner,Name,Title,Purpose,Summary,Hypothesis,Problem")


        '        Case "Person"
        '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1),  Name nchar(200) ,Mail nchar(100),Tel  nvarchar(100) , Prof text, Image nvarchar(100) ,State int)")
        '            .DBACS_Field.Add(Key, "ID,Name,Mail,Tel,Prof,Image ,State")


        '        Case "Desk"
        '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,OwnerID int,ProjID int,Page int,Type nvarchar(100) ,Title nvarchar(300) ,State int)")
        '            .DBACS_Field.Add(Key, "ID,OwnerID,ProjID,Page,Type,Title ,State")

        '        Case "Box"
        '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,DeskID int,ProjID int,OwnerID int,BluePrint text ,State int)")
        '            .DBACS_Field.Add(Key, "ID,DeskID,ProjID,OwnerID,BluePrint ,State")


        '        Case "Tbox"
        '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,BoxID int,Num int,Text Text)")
        '            .DBACS_Field.Add(Key, "ID,BoxID,Num,Text ,State")

        '        Case "Draw"
        '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,BoxID int,Num int,Text Text ,State int)")
        '            .DBACS_Field.Add(Key, "ID,BoxID,Num,Text , State")

        '        Case "Pbox"
        '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,BoxID int,Num int,Pict TEXT ,State int)")
        '            .DBACS_Field.Add(Key, "ID,BoxID,Num,Pict ,State")

        '        Case "BDat"
        '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,BoxID int,Num int,Dat varbinary(max) ,State int)")
        '            .DBACS_Field.Add(Key, "ID,BoxID,Num,Dat ,State")

        '        'Key= "Webbox"
        '        '    .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,BaseID int,Type nvarchar(100) ,Style int,Text Text)"
        '        '    .DBACS_Field.Add(Key,"ID,BaseID,Type,Path,Style,Text"
        '        'DB_Check(Lab, key)

        '        Case "Element"
        '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,ProjID int,Type nvarchar(100) ,Name nvarchar(100),Path nvarchar(500) ,State int)")
        '            .DBACS_Field.Add(Key, "ID,ProjID,Type,Name,Path ,State")

        '        Case "Role"
        '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,LID int ,PID int,Role nvarchar(100) ,State int )")
        '            .DBACS_Field.Add(Key, "ID,LID,PID,Role ,State")

        '    End Select

        'End With
        DB_Check(Lab, Key)


    End Sub


    Private Function DBAcs_Del_Dat(ByRef Db As DBACS_Type, ByRef Cnd As String) As Integer
        'DBレコードの削除
        '----------------
        'DB 指定Key の　値　を格納
        '--------------------------
        Dim tmpR As String
        Dim DBCon As SqlConnection
        Dim DBCom As SqlCommand
        Dim TmpCmd As String

        With Db
            DBCon = New SqlConnection(.DB_Path & "initial catalog = " & .DB_Name)
            TmpCmd = "DELETE   FROM " & .TBL_Name & "  WHERE " & Cnd
        End With

        DBCom = New SqlCommand(TmpCmd, DBCon)
        DBCom.CommandTimeout = 0

        DBCom.Connection.Open()


        For I As Integer = 1 To 50
            Try
                tmpR = DBCom.ExecuteScalar()

                Exit For
            Catch ex As Exception
                Log_Show(2, "DB登録エラー" & I & ex.Message)
                For J As Integer = 1 To 100
                    System.Windows.Forms.Application.DoEvents()
                Next
            End Try
        Next

        DBCom.Connection.Close()
        DBCon.Close()
        DBCon.Dispose()




        Return tmpR
    End Function



    'Private Function DBAcs_Tbl_CheckExist(ByRef Db As DBACS_Type) As Integer

    '    Dim TmpR As Integer
    '    With Db
    '        TmpR = DB_Tbl_CheckExist(.DB_Path, .DB_Name, .TBL_Name, .TBL_Type)
    '    End With


    '    Return TmpR
    'End Function
    Private Sub DB_Check(ByRef Lab As Lab_Type, Key As String)
        With Lab

            Select Case .Db_Type
                Case "SQLite"
                    DB_Tbl_CheckExistSQLite(Lab, Key, .DBACS_Type(Key))'TBL存在チェック&ないなら作成
                Case "MSSQL", "MSSQLW"
                    DB_Check_Exist_DB_MSSQL(.Db_Type, .DbPath, .Db_Type & Key) 'DB 存在チェック&ないなら作成
                    DB_Tbl_CheckExistMSSQL(.Db_Type, .DbPath, .Db_Type & Key, "TBL_" & Key, .DBACS_Type(Key)) 'TBL存在チェック&ないなら作成
            End Select

        End With

    End Sub

    Private Sub DB_Check_MSSQL(ByRef Lab As Lab_Type, Key As String)
        With Lab
            DB_Check_Exist_DB_MSSQL(.Db_Type, .DbPath, .Db_Type & Key) 'DB 存在チェック&ないなら作成
            DB_Tbl_CheckExistMSSQL(.Db_Type, .DbPath, .Db_Type & Key, "TBL_" & Key, .DBACS_Type(Key)) 'TBL存在チェック&ないなら作成
        End With

    End Sub


    Private Function DB_Check_Exist_DB(ByRef DB_type As String, ByRef Db_Path As String) As Integer
        'DBがないときは空DBをコピー
        Try
            If System.IO.File.Exists(gDB_Path) Then
                Return True
            Else
                System.IO.File.Copy(System.Windows.Forms.Application.StartupPath & "\ENV\GI.sqlite3", gDB_Path)
            End If
            Return True

        Catch ex As Exception
            Log_Show(1, ex.Message)
        End Try

        Return False

    End Function

    Private Function DB_Check_Exist_DB_MSSQL(ByRef DB_type As String, ByRef Db_Path As String, ByRef DB_Name As String) As Integer

        Dim DBCon, cn As SqlConnection
        Dim DBCom, Cmd As SqlCommand
        Dim TmpR As Integer

        TmpR = True
        'DBの存在チェック
        'cn = New SqlConnection("Server=.\SQLEXPRESS;" & "Database=master;" & "integrated security=true")
        Log_Show(1, "DB存在チェック")
        Try
            cn = New SqlConnection(Db_Path)
            Cmd = New SqlCommand("SELECT COUNT(*) FROM sysdatabases WHERE name='" & DB_Name & "'", cn)

            Cmd.Connection.Open()

            TmpR = True
        Catch ex As Exception
            Log_Show(1, ex.Message)
            TmpR = False
        End Try

        'データベースが存在しない
        If (Cmd.ExecuteScalar() = 0) Then

            'データベースを作成
            DBCon = New SqlConnection(Db_Path)
            DBCom = New SqlCommand("create database " & DB_Name, DBCon)

            DBCom.CommandTimeout = 0

            DBCom.Connection.Open()

            Try
                DBCom.ExecuteNonQuery()
                TmpR = True
            Catch ex As Exception
                Log_Show(1, ex.Message)
                TmpR = False
            End Try
            DBCom.Connection.Close()
            DBCon.Close()
        End If

        cn.Close()

        Return TmpR


    End Function
    Private Function DB_Tbl_CheckExistSQLite(ByRef Lab As Lab_Type, ByRef Tbl_Name As String, ByRef Tbl_Type As String) As Integer

        'テーブルの存在チェック
        '---------------------------------
        Try
            Dim TmpR As Integer
            TmpR = False
            If Tbl_Name = "" Then Return TmpR


            ' コネクション作成
            Dim con As SQLiteConnection
            con = New SQLiteConnection(GetConnectionString(Lab.DbPath))
            con.Open()

            'テーブルチェック
            Using cmd = con.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE TYPE = 'table' AND name= '" & Tbl_Name & "'"
                Dim tmpC As Integer = cmd.ExecuteScalar()
                If (tmpC <> 0) Then

                    Return True
                End If
            End Using

            ' テーブル作成SQL
            Dim TmpCmd As String
            TmpCmd = "CREATE TABLE " & Tbl_Name & Tbl_Type
            Using cmd = con.CreateCommand()
                cmd.CommandText = TmpCmd
                cmd.ExecuteNonQuery()
            End Using
            Return True

        Catch ex As Exception
            Log_Show(1, ex.Message)
        End Try
        Return False

    End Function


    Private Function DB_Tbl_CheckExistMSSQL(ByRef DB_Type As String, ByRef Db_Path As String, ByRef DB_Name As String, ByVal Tbl_Name As String, ByRef Tbl_Type As String) As Integer

        'テーブルの存在チェック
        '---------------------------------
        '戻り値　存在　2　新規　1　エラー0　
        Dim TmpR As Integer
        TmpR = 0
        If Tbl_Name = "" Then Return TmpR

        Dim DBCon, cn As SqlConnection
        Dim DBCom, Cmd As SqlCommand



        'テーブルの存在チェック
        cn = New SqlConnection(Db_Path & ";Initial Catalog =" & DB_Name)

        Cmd = New SqlCommand("Select COUNT(*) FROM sys.objects WHERE object_id = OBJECT_ID( '" & Tbl_Name & "')", cn)

        For J = 0 To 50
            Try
                Cmd.Connection.Open()
                Exit For

            Catch ex As Exception

                Log_Show(1, ex.Message)
            End Try

        Next

        'SELECT COUNT(*) FROM sys.objects WHERE object_id = OBJECT_ID('dbo.Table')
        'テーブルが存在しない
        If (Cmd.ExecuteScalar() = 0) Then

            'テーブルを作成
            Dim TmpCmd As String
            TmpCmd = " create table " & Tbl_Name & Tbl_Type

            DBCon = New SqlConnection(Db_Path & ";Initial Catalog =" & DB_Name)
            DBCom = New SqlCommand(TmpCmd, DBCon)
            Try


                DBCom.Connection.Open()
                DBCom.ExecuteNonQuery()
                DBCom.Connection.Close()
                DBCon.Close()
                '  DBCon.Dispose()

                TmpR = 1

            Catch ex As Exception

                Log_Show(1, ex.Message)
            End Try
        Else
            TmpR = 2
        End If
        cn.Close()

        Return TmpR

    End Function

    Private Function DBACS_Get_ID_ByCnd(ByRef DB As DBACS_Type, ByRef Cnd As String) As Integer()
        'DB 指定条件 の　ID　を返却
        '--------------------------
        Dim XR() As Integer
        ReDim XR(0)
        Dim DBCon As SqlConnection
        Dim DBCom As SqlCommand
        Dim DBReader As SqlClient.SqlDataReader

        Dim TmpCmd, TmpDB As String

        With DB
            TmpDB = .DB_Path & ";Initial Catalog = " & .DB_Name
            TmpCmd = "Select * FROM [" & .DB_Name & "].[dbo].[" & .TBL_Name & "] WHERE  " & Cnd
        End With

        DBCon = New SqlConnection(TmpDB)
        DBCom = New SqlCommand(TmpCmd, DBCon)

        DBCom.CommandTimeout = 0
        DBCom.Connection.Open()

        For I As Integer = 1 To 50
            Try
                DBReader = DBCom.ExecuteReader()

                While DBReader.Read
                    XR(XR.Length - 1) = Val(DBReader(0) & ".0")
                    ReDim Preserve XR(XR.Length)
                End While
                DBReader.Read()


                Exit For
            Catch ex As Exception
                Log_Show(2, "DBエラー" & I & ex.Message)
                For J As Integer = 1 To 50
                    System.Windows.Forms.Application.DoEvents()
                Next
            End Try
        Next
        DBCom.Dispose()
        DBReader.Close()
        DBCon.Close()
        DBCon.Dispose()

        Return XR



    End Function



    Private Sub DBAcs_Update_Data_ByKey(ByRef DB As DBACS_Type, ByRef TBL_Name As String, ByRef ID As Integer, ByRef Cmd As String)
        'DB 指定Key の　値　を格納
        '--------------------------
        Dim tmpR As String
        Dim DBCon As SqlConnection
        Dim DBCom As SqlCommand
        Dim TmpCmd As String


        TmpCmd = "UPDATE " & TBL_Name & " Set  " & Cmd & "  WHERE ID = " & ID

        With DB
            DBCon = New SqlConnection(.DB_Path & ";Initial Catalog =" & .DB_Name)
        End With

        DBCom = New SqlCommand(TmpCmd, DBCon)
        DBCom.Connection.Open()


        For I As Integer = 1 To 50
            Try
                tmpR = DBCom.ExecuteScalar()

                Exit For
            Catch ex As Exception
                Log_Show(2, "DB登録エラー" & I & ex.Message)
                For J As Integer = 1 To 100
                    System.Windows.Forms.Application.DoEvents()
                Next
            End Try
        Next

        DBCom.Connection.Close()
        DBCon.Close()
        DBCon.Dispose()




    End Sub
    Private Function DBAcs_Insert(ByRef Lab As Lab_Type, ByRef Name As String, ByRef Key As String, ByRef Val As String) As Long
        'DB 指定Key の　値　を格納
        '--------------------------
        Dim tmpR As Integer
        Try
            Dim TmpCmd As String

            TmpCmd = "INSERT INTO " & Name
            TmpCmd &= " ( " & Key & ")   VALUES ( " & Val & ")"

            Using con = New SQLiteConnection(GetConnectionString(Lab.DbPath))
                con.Open()
                Using cmd = con.CreateCommand()
                    cmd.CommandText = TmpCmd
                    cmd.ExecuteNonQuery()
                End Using
                Using cmd = con.CreateCommand()
                    'cmd.CommandText = "SELECT last_insert_rowid();"
                    cmd.CommandText = "Select max(ID) FROM " & Name
                    tmpR = cmd.ExecuteScalar()
                End Using
                Return tmpR
            End Using

        Catch ex As Exception
            Log_Show(1, ex.Message)
            Return False
        End Try
        Return tmpR
    End Function

    Private Function DBAcs_Insert_MSSQL(ByRef Lab As Lab_Type, ByRef Name As String, ByRef Key As String, ByRef Val As String) As Long
        'DB 指定Key の　値　を格納
        '--------------------------
        Dim tmpR As Integer

        Dim DBCon As SqlConnection
        Dim DBCom As SqlCommand
        Dim TmpCmd As String
        tmpR = -1

        With Lab

            DBCon = New SqlConnection(.DbPath & ";Initial Catalog =" & .Db_Type & Name)

            TmpCmd = "INSERT INTO " & "TBL_" & Name
            TmpCmd &= " ( " & Key & ")  OUTPUT INSERTED.ID  VALUES ( " & Val & ")"
            TmpCmd &= "; Select SCOPE_IDENTITY();"
        End With

        DBCom = New SqlCommand(TmpCmd, DBCon)
        DBCom.Connection.Open()



        For I As Integer = 1 To 50
            Try
                tmpR = CInt(DBCom.ExecuteScalar())

                Exit For
            Catch ex As Exception
                Log_Show(1, "DB登録エラー" & I & ex.Message)
                For J As Integer = 1 To 100
                    System.Windows.Forms.Application.DoEvents()
                Next
            End Try
        Next

        DBCom.Connection.Close()
        DBCon.Close()
        DBCon.Dispose()

        Return tmpR
    End Function

    Private Function DBAcs_Update(ByRef Lab As Lab_Type, ByRef Name As String, ByRef Key As String, ByRef Val As String) As Integer
        'DB 指定Key の　値　を格納
        '--------------------------
        Try
            Dim TmpCmd As String

            TmpCmd = "UPDATE " & Name
            TmpCmd &= "  SET " & Key & " WHERE  " & Val

            Using con = New SQLiteConnection(GetConnectionString(Lab.DbPath))
                con.Open()
                Using cmd = con.CreateCommand()
                    cmd.CommandText = TmpCmd
                    cmd.ExecuteNonQuery()
                End Using
            End Using


        Catch ex As Exception
            Log_Show(1, ex.Message)
            Return False
        End Try

        Return True
    End Function
    Private Function DBAcs_UpdateMSSQL(ByRef Lab As Lab_Type, ByRef Name As String, ByRef Key As String, ByRef Val As String) As Integer
        'DB 指定Key の　値　を格納
        '--------------------------
        Dim tmpR As Integer

        Dim DBCon As SqlConnection
        Dim DBCom As SqlCommand
        Dim TmpCmd As String
        tmpR = -1

        With Lab

            DBCon = New SqlConnection(.DbPath & ";Initial Catalog =" & .Db_Type & Name)
            TmpCmd = "UPDATE " & "TBL_" & Name
            TmpCmd &= "  SET " & Key & " WHERE  " & Val
            '  TmpCmd &= "; Select SCOPE_IDENTITY();"

        End With

        DBCom = New SqlCommand(TmpCmd, DBCon)
        DBCom.Connection.Open()

        For I As Integer = 1 To 50
            Try
                tmpR = CInt(DBCom.ExecuteScalar())

                Exit For
            Catch ex As Exception
                Log_Show(1, "DB登録エラー" & I & ex.Message)
                For J As Integer = 1 To 100
                    System.Windows.Forms.Application.DoEvents()
                Next
            End Try
        Next

        DBCom.Connection.Close()
        DBCon.Close()
        DBCon.Dispose()

        Return tmpR
    End Function
    'Private Function DBAcs_Get_MaxID(ByRef Lab As Lab_Type, ByRef Name As String, ByRef Key As String) As Long

    '    Dim tmpR As Integer

    '    Dim DBCon As SqlConnection
    '    Dim DBCom As SqlCommand
    '    Dim TmpCmd As String
    '    tmpR = -1

    '    With Lab

    '        DBCon = New SqlConnection(.DbPath & ";Initial Catalog =" & .DBname & Name)
    '        TmpCmd = "SELECT MAX(" & Key & ") FROM " & "TBL_" & Name

    '    End With

    '    DBCom = New SqlCommand(TmpCmd, DBCon)
    '    DBCom.Connection.Open()



    '    For I As Integer = 1 To 50
    '        Try
    '            tmpR = CLng(DBCom.ExecuteScalar())

    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(1, "DB登録エラー" & I & ex.Message)
    '            For J As Integer = 1 To 100
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next

    '    DBCom.Connection.Close()
    '    DBCon.Close()
    '    DBCon.Dispose()

    '    Return tmpR

    'End Function

    Private Function DBAcs_Get_DataS(ByRef Lab As Lab_Type, Name As String, Action As String, ByRef Condition As String) As DataSet

        'DB 指定Key の　値　を格納
        '--------------------------
        Try
            Dim adapter = New SQLiteDataAdapter()
            '  Dim dtTbl As New DataTable()
            Dim ds As DataSet
            ds = New DataSet()

            Dim TmpCmd As String
            TmpCmd = "Select " & Action & " FROM " & Name & " WHERE " & Condition

            Using con = New SQLiteConnection(GetConnectionString(Lab.DbPath))
                con.Open()
                Using cmd = con.CreateCommand()
                    cmd.CommandText = TmpCmd
                    adapter = New SQLiteDataAdapter(TmpCmd, con)
                    '  adapter.Fill(dtTbl)
                    adapter.Fill(ds)
                End Using
            End Using

            Return ds


        Catch ex As Exception
            Log_Show(1, ex.Message)

        End Try

    End Function

    Private Function DBAcs_Get_DataS_MSSQL(ByRef Lab As Lab_Type, Name As String, Action As String, ByRef Condition As String) As DataSet

        'DB 指定Key の　値　を格納
        '--------------------------
        Dim tmpR As Integer

        Dim DBCon As SqlConnection
        Dim TmpCmd As String
        tmpR = -1
        With Lab
            TmpCmd = "Select " & Action & " FROM [" & .Db_Type & Name & "].[dbo].[" & "TBL_" & Name & "] WHERE " & Condition
            DBCon = New SqlConnection(.DbPath & ";Initial Catalog =" & .Db_Type & Name)
            DBCon.Open()

        End With
        'DBCom = New SqlCommand(TmpCmd, DBCon)

        Dim adapter As SqlDataAdapter
        Dim ds As DataSet
        ds = New DataSet()
        adapter = New SqlDataAdapter()
        'データの取得

        For I As Integer = 1 To 50
            Try

                adapter.SelectCommand = New SqlCommand(TmpCmd, DBCon)

                adapter.SelectCommand.CommandType = CommandType.Text

                adapter.Fill(ds)
                Exit For
            Catch ex As Exception
                Log_Show(1, "エラー" & I & ex.Message)
                For J As Integer = 1 To 100
                    System.Windows.Forms.Application.DoEvents()
                Next
            End Try
        Next
        DBCon.Close()
        DBCon.Dispose()
        adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

        DBCon.Close()

        Return ds



    End Function



    Private Function File_Set_byDir(ByRef FolderName As String, ByRef Key As String, ByRef TmpL() As String) As Integer
        '拡張子指定でファイル名リストを作成する
        '---------------------------------------
        Dim FileName As String
        Dim Lmax As Integer
        '指定のフォルダがない
        If System.IO.Directory.Exists(FolderName) = False Then
            Return 0
        End If

        '指定のフォルダ内の指定の拡張子のファイルを全て列挙する
        If TmpL Is Nothing Then
            ReDim TmpL(0)
            Lmax = 0
        Else
            Lmax = TmpL.Length
        End If

        If FolderName.Length > 1 Then
            '指定の拡張子のファイルだけ取得する場合
            For Each FileName In System.IO.Directory.GetFiles(FolderName, "*" & Key, System.IO.SearchOption.TopDirectoryOnly)

                ReDim Preserve TmpL(Lmax)
                TmpL(Lmax) = FileName
                Lmax = Lmax + 1
            Next
        End If

        Return Lmax

    End Function





    Private Sub Log_Show(ByRef ID As Integer, ByRef TmpM As String)

        Select Case ID

            Case 1
                If TmpM = "DEL" Then
                    ToolStripStatusLabel1.Text = ""
                Else

                    ToolStripStatusLabel1.Text = TmpM & Chr(13) & Chr(10)
                End If

            Case 2

                ToolStripStatusLabel2.Text = TmpM


            Case 5
                ToolStripStatusLabel3.Text = TmpM


        End Select
        System.Windows.Forms.Application.DoEvents()


    End Sub



    Private Sub Proj_Rename(ByRef lab As Lab_Type, NewName As String)
        DBAcs_Update(lab, "Proj", "Name = '" & NewName & "'", "ID = " & lab.ProjID.ToString)
        gLab.ProjName = NewName
        ' Me.Text = "Ginger-breadboard       " & gLab.ProjName
    End Sub



    Private Function Image_Init(ByRef TmpD As String) As Image_List_Type
        'ICON類の収集
        '----------------------------------------
        Log_Show(1, "ICON類の収集")


        Dim Image0 As System.Drawing.Image
        Dim Key As String
        If System.IO.Directory.Exists(TmpD) = False Then
            Log_Show(5, "イメージリソースが見つかりません ")
            Return Nothing
        End If



        Dim Flist() As String
        ReDim Flist(0)
        Dim Cnt As Integer
        File_Set_byDir(TmpD, ".ico", Flist)
        File_Set_byDir(TmpD, ".gif", Flist)
        File_Set_byDir(TmpD, ".png", Flist)
        File_Set_byDir(TmpD, ".jpg", Flist)
        File_Set_byDir(TmpD, ".jpeg", Flist)
        With gImg_List
            .Int(Flist.Length)
            .Image.ImageSize = New Size(32, 32)
            For L = 0 To Flist.Length - 1
                If Flist(L) <> "" Then
                    Key = IO.Path.GetFileNameWithoutExtension(Flist(L))

                    Image0 = Image.FromFile(Flist(L))
                    .Image.Images.Add(Key, Image0)

                    ReDim Preserve .Key(.Image.Images.Count - 1)
                    .Key(.Image.Images.Count - 1) = Key
                End If
            Next
        End With
        Return gImg_List

        ' Button20.Image = gImg_List.Image.Images(19)  
    End Function


    Private Function Image_Get_ID(ByRef Img As Image_List_Type, ByRef Key As String) As Integer
        'ICONのID取得
        '----------------------------------------
        Dim tmpR As Integer
        tmpR = 0
        With Img
            For L = 0 To .Image.Images.Count - 1
                If Key = .Key(L) Then
                    Return L
                End If
            Next
        End With

        Return 0
    End Function





    Private Sub Element_Act_Item(ByRef Lab As Lab_Type, DID As Long)
        Dim Ds As DataSet
        Dim Cnd As String

        Cnd = "ID = " & DID.ToString
        Ds = DBAcs_Get_DataS(Lab, "Element", "Type,path", Cnd)

        With Ds.Tables(0).Rows(0)
            Select Case .Item(0)
                Case "File"
                    If System.IO.File.Exists(.Item(1)) Then
                        System.Diagnostics.Process.Start(.Item(1))
                        Log_Show(1, "OK")
                    Else

                        Log_Show(1, "ファイルが消えてる")
                    End If

                Case "Text"
                Case "Url"
                    If InStr(.Item(1), "http") > 0 Then
                        System.Diagnostics.Process.Start(.Item(1))
                    Else
                        System.Diagnostics.Process.Start("https://" & .Item(1))
                    End If
            End Select

        End With





    End Sub
    Private Sub Box_MakeFromElement(ByRef Lab As Lab_Type, Type As String, Cp As Point, ELM As Element_Type)
        'ポトンされたデペンドをスタック化
        '----------------------------
        Select Case Type
            Case "FILE"
                Select Case ELM.Type.ToLower
                    Case "ico", "bmp", "jpg", "gif", "png", "exig", "tiff"
                    Case "text", "bmp", "jpg", "gif", "png", "exig", "tiff"
                    Case "ico", "bmp", "jpg", "gif", "png", "exig", "tiff"
                End Select


                'Box_Makes_OnDESK(gLab, "Viewer", Cp.X, Cp.Y, 1)


            Case Else
        End Select


        'スタック作成
        '   Box_Makes_OnDESK(gLab, Source(0).Tag, cp.X, cp.Y, 1)

        'Element(cnt).Path = Drops(L)
        'Element(cnt).Name = System.IO.Path.GetFileNameWithoutExtension(Drops(L))
        'Element(cnt).Type = "File"


    End Sub

    Private Sub Element_Add_Data(ByRef Lab As Lab_Type, RID As Long, ByRef ELM() As Element_Type)
        'ポトンされたデペンドを登録 (同じものでも別登録になる/サボリ後で考える)
        '----------------------------
        If RID = 0 Then Return
        Dim tmpD As String
        Dim DID As Long
        Dim Cnd As String
        Dim Key As String
        'DB登録
        Cnd = "ID In ( "
        For L = 0 To ELM.Length - 1
            tmpD = ""
            If ELM(L).Name <> "" Then

                With ELM(L)
                    tmpD = gProjID.ToString & ","
                    tmpD &= "'" & .Type & "',"
                    tmpD &= "'" & .Name & "',"
                    tmpD &= "'" & .Path & "',"
                    tmpD &= "1"
                End With
                With Lab.ElementAT
                    Key = Mid(Lab.DBACS_Field("Element"), Len("ID,") + 1)
                    DID = DBAcs_Insert(Lab, "Element", Key, tmpD) 'ElementDB登録
                End With
                Cnd = Cnd & DID.ToString & ","
                ELM(L).ID = DID
            End If

        Next


        'ElementDBから読み出して、Listviewに登録
        Cnd = "ProjID= " & Str(Lab.ProjID)

        Dim Ds As DataSet
        With Lab.ElementAT
            Ds = DBAcs_Get_DataS(Lab, "Element", .View.Columns.DBKeys, Cnd)
            Element_Show_List(Lab, Ds, True) '表示
        End With



    End Sub


    Private Sub Button25_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs)

        'ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If


    End Sub
    Private Sub Button25_DragDrop(sender As Object, e As System.Windows.Forms.DragEventArgs)
        Dim DropS As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

        'Proj_Action_Files("NOT", DropS)


    End Sub




    Private Sub Button21_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs)

        'ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If


    End Sub

    Private Sub Button21_DragDrop(sender As Object, e As System.Windows.Forms.DragEventArgs)
        Dim DropS As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

        'Proj_Action_Files("FIXED", DropS)

    End Sub



    Private Sub Button25_Click(sender As System.Object, e As System.EventArgs)
        'Histroy_Go(gLab, 1)
    End Sub


    Private Sub Proj_Make_List(ByRef Lab As Lab_Type, ByRef CNode As TreeNode)
        'データの表示
        '--------------------------

        Dim DS As DataSet
        Log_Show(1, "データ一覧作成中")


    End Sub






    Private Sub Columns_Set_Listview(ByRef Listview As ListView, ByRef columns As Columns_Type)
        'LISTVIEWのカラム指定
        '-----------------------
        With Listview
            If .Items.Count > 0 Then
                .Clear()
            End If
        End With

        With columns
            For L = 0 To .Name.Length - 1
                Listview.Columns.Add(.Name(L), Integer.Parse(.Width(L)), HorizontalAlignment.Left)
            Next

        End With

    End Sub
    Private Sub Columns_Set_Env(ByRef columns As Columns_Type, ClmKey As String)
        'カラム指定
        '-----------------------
        Dim Name As String
        Dim Wid As String
        Dim DBkeys As String

        Select Case ClmKey

            Case "[Proj]"
                DBkeys = "Name,Owner,ID"
                Name = "Name,Owner,ID"
                Wid = "300,80,5"
            Case "[Element]"
                DBkeys = "Name,Type,Path,ID"
                Name = "Name,Type,Path,ID"
                Wid = "300,30,80,10"
            Case "[Stack]"
                DBkeys = "Name,Type,ID"
                Name = "Name,Type,ID"
                Wid = "300,30,10"
        End Select
        With columns
            .Type = ClmKey
            .DBKeys = DBkeys
            .Name = Split(Name, ",")
            .Width = Split(Wid, ",")

        End With





    End Sub



    Private Sub Proj_Show_List(ByRef View As ListView, ByRef Columns As Columns_Type, ByRef Ds As DataSet, clearFlg As Boolean)
        'LISTVIEWにProjデータを表示
        '-----------------------------
        Try

            If Ds.Tables(0).Rows.Count = 0 Then
                View.Clear()
                Return
            End If


            Dim ITM As ListViewItem
            If clearFlg = True Then
                Columns_Set_Listview(View, Columns)
            End If

            With Ds.Tables(0)

                For L = 0 To .Rows.Count - 1
                    ITM = View.Items.Add(.Rows(L).Item(0))
                    ITM.Tag = Val(.Rows(L).Item(.Columns.Count - 1)) '最後のアイテムをレポジトリのIDにしてるので
                    For K = 1 To .Columns.Count - 1

                        If .Rows(L).Item(K) Is DBNull.Value Then
                        Else
                            View.Items(L).SubItems.Add(.Rows(L).Item(K))
                        End If

                    Next
                Next

                If .Rows.Count > 0 Then
                    View.Items(0).Selected = True
                End If
            End With
        Catch ex As Exception
            Log_Show(1, ex.Message)
        End Try

    End Sub
    Private Sub Element_Show_List(ByRef Lab As Lab_Type, ByRef Ds As DataSet, clearFlg As Boolean)
        'ElementデータをLISTVIEWに表示
        '-----------------------------
        Dim Itm As ListViewItem
        Dim Img As Image

        With Lab.ElementAT.View
            If Ds.Tables(0).Rows.Count = 0 Then
                .ListView.Clear()
                Return
            End If

            If clearFlg = True Then
                Columns_Set_Listview(.ListView, .Columns)
            End If

            For L = 0 To Ds.Tables(0).Rows.Count - 1
                Itm = .ListView.Items.Add(Ds.Tables(0).Rows(L).Item(0))
                Itm.ImageIndex = Element_Get_ImageID(Lab, Ds.Tables(0).Rows(L).Item(2))

                Itm.Tag = Val(Ds.Tables(0).Rows(L).Item(Ds.Tables(0).Columns.Count - 1)) '最後のアイテムをDIDにしてるので
                For K = 1 To Ds.Tables(0).Columns.Count - 1
                    .ListView.Items(L).SubItems.Add(Ds.Tables(0).Rows(L).Item(K))
                Next
            Next

            'If Ds.Tables(0).Rows.Count > 0 Then
            '    .View.Items(0).Selected = True
            'End If
            .ListView.Refresh()
        End With
    End Sub

    Private Function Element_Get_ImageID(ByRef Lab As Lab_Type, Key As String) As Integer

        Dim tmpR As Integer
        Dim Ext As String
        tmpR = Image_Get_ID(Lab.Image_List, Key)
        If tmpR > 0 Then Return tmpR

        If InStr(Key, "http") > 0 Then
            tmpR = Image_Get_ID(Lab.Image_List, "http")
            If tmpR > 0 Then Return tmpR
        End If

        If System.IO.Directory.Exists(Key) Then
            tmpR = Image_Get_ID(Lab.Image_List, "Folder")
            If tmpR > 0 Then Return tmpR
        End If
        Try
            Ext = Path.GetExtension(Key)
            If Ext <> "" Then
                tmpR = Image_Get_ID(Lab.Image_List, Ext)
                If tmpR > 0 Then Return tmpR

                tmpR = Image_Add_DatabyPath(Lab.Image_List, Key, Ext)
                If tmpR > 0 Then Return tmpR

            End If
        Catch

        End Try




        Return 1

    End Function

    Private Function Image_Add_DatabyPath(ByRef Image_List As Image_List_Type, path As String, Key As String) As Integer
        'ファイルに関連したアイコンをイメージリストに設定
        Dim tmpIcon As Icon
        Dim max As Integer
        Try
            tmpIcon = Icon.ExtractAssociatedIcon(path)
        Catch
            Return 0
        End Try
        If tmpIcon Is Nothing Then Return 0

        With Image_List
            .Image.Images.Add(Key, tmpIcon.ToBitmap)
            max = .Image.Images.Count - 1
            ReDim Preserve .Key(max)
            .Key(max) = Key
            Return max
        End With

    End Function


    Private Function Icon_GetbyPath(Fpath As String) As Icon
        Dim tmpIcon As Icon
        tmpIcon = Icon.ExtractAssociatedIcon(Fpath)
        Return tmpIcon

    End Function


    Private Sub Proj_ListView_Item_Remove(ByRef ObjCtrl As System.Windows.Forms.ListView)
        'Listから指定TAG 0を削除する
        '----------------------------------
        Dim DX() As Integer

        With ObjCtrl
            For L = .Items.Count - 1 To 0 Step -1
                If .Items(L).Tag = 0 Then
                    .Items(L).Remove()
                End If
            Next
        End With

    End Sub



    'Private Sub Proj_Show_Atr(ByRef Lab As Lab_Type, ByRef RID As Long)
    '    'レポジトリのATR表示
    '    '--------------------------
    '    Dim Ds As DataSet
    '    Ds = DBAcs_Get_DataS(Lab, "Proj", "Note,Prof", "ID=" & Str(RID))
    '    If Ds.Tables(0).Rows.Count = 0 Then Return
    '    With Ds.Tables(0)

    '        'Lab.ElementAT.Note.Text = .Rows(0).Item(0)
    '        'Lab.ElementAT.Prof.Text = .Rows(0).Item(1)
    '    End With
    '    Proj_Show_Element(Lab, RID)
    '    'RichTextBox3.Text = Lab.ElementAT.Note.Text


    'End Sub
    Private Sub Proj_Show_Element(ByRef Lab As Lab_Type, ByRef PID As Long)
        'レポジトリ対応のElement表示
        '--------------------------
        Dim Ds As DataSet
        Dim Cnd As String
        '■ちょっと待ってね
        'Ds = DBAcs_Get_DataS(Lab, "Element", "ID", "ProjID=" & Str(PID))
        'If Ds.Tables(0).Rows.Count = 0 Then
        '    Lab.ElementAT.View.ListView.Clear()
        '    '   Columns_Set_Listview(Lab.Element.View, Lab.Element.Columns)
        '    Return
        'End If

        'Cnd = "ID IN ( "
        'With Ds.Tables(0)
        '    For L = 0 To .Rows.Count - 2
        '        Cnd = Cnd & .Rows(L).Item(0) & ","
        '    Next
        '    Cnd = Cnd & .Rows(.Rows.Count - 1).Item(0) & ")"
        'End With

        Cnd = "ProjID= " & Str(PID)
        With Lab.ElementAT
            Ds = DBAcs_Get_DataS(Lab, "Element", .View.Columns.DBKeys, Cnd)
            Element_Show_List(Lab, Ds, True)

        End With

    End Sub




    Private Function Proj_Atr_Add_Item(ByRef LV As System.Windows.Forms.ListView, ByRef DID As Integer, ByRef BType As String, ByRef SType As String, ByRef FX As String, ByRef VX As String, ByRef CX As String, ByRef State As Integer, ByRef Ratio As Double) As ListViewItem
        'ATR ITEMをADD
        '---------------------

        Dim lvi As ListViewItem
        lvi = LV.Items.Add(BType)
        lvi.Tag = DID


        lvi.ImageIndex = Image_Get_ID(gImg_List, BType & "_")
        lvi.SubItems.Add(SType)
        lvi.SubItems.Add(FX)
        lvi.SubItems.Add(VX)
        lvi.SubItems.Add(CX)
        lvi.SubItems.Add(State)
        lvi.SubItems.Add(Ratio)

        Return lvi
    End Function

    Private Sub Proj_Atr_New_Item(ByRef ListV As ListView, ByRef BType As String)
        ''ATR ITEMを新規追加
        '-----------------------
        With ListV
            Proj_Atr_Add_Item(ListView2, 0, BType, "", "", "", "", 0, 0)
            .SelectedItems.Clear()
            .Items(.Items.Count - 1).EnsureVisible()
            .Items(.Items.Count - 1).Selected = True


        End With


    End Sub










    Private Sub Button14_Click(sender As System.Object, e As System.EventArgs)
        'ルールの適用

        'Rule_Apply(gStudy(ComboBox8.SelectedIndex)) 'ルールの適用

    End Sub


    Private Sub Button17_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs)
        'ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub




    Private Sub ContextMenuStrip1_ItemClicked(sender As Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContextMenuStrip1.ItemClicked

        'Menu_Act(gBff_ObjName, "0\" & e.ClickedItem.ToString)

        Select Case e.ClickedItem.Tag
            Case "NEW"
                Proj_Make_New(gLab)
            Case "RENAME"
                If ListView1.SelectedItems.Count = 0 Then Return

                ListView1.SelectedItems(0).BeginEdit()
            Case "DEL"
                Proj_Del(gLab)

            Case Else
        End Select
    End Sub






    Private Sub ListView1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles ListView1.KeyDown


        Dim DeskID As Integer
        DeskID = 1
        With ListView1

            '削除
            If e.KeyData = Keys.Delete Then
                If IsNothing(.SelectedItems) Then Return
                If .SelectedItems.Count = 0 Then Return

                'Histroy_Start(gLab.Desk(DeskID).History)

                For L = .SelectedItems.Count - 1 To 0 Step -1

                    '   Proj_Act_Dat_Del_Item(gLab, .SelectedItems(L), True)

                    .SelectedItems(L).Remove()
                Next

                'Histroy_End(gLab.Desk(DeskID))  '▲History終了
                If Not (.FocusedItem Is Nothing) Then
                    .FocusedItem.Selected = True
                End If

            End If

            '作成はしない
            If e.KeyData = Keys.Insert Then
                'If Proj_Node_Make_Check(.SelectedNode, "NEW") = True Then
                '     'Histroy_start(gLab.Desk(1).history)
                '    Proj_Act_Node_Make(gLab,1.History, .SelectedNode, "NEW", True)
                'End If
            End If

        End With




    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        'プロジェクトの読み込み

        With ListView1
            If .SelectedItems Is Nothing Then Return
            If .SelectedItems.Count <= 0 Then Return
            If .SelectedItems.Count > 1 Then Return

            If .SelectedItems(0).Tag Is Nothing Then Return
            gProjID = Val(.SelectedItems(0).Tag)

            gLab.ProjName = .SelectedItems(0).Name
            '  Me.Text = "Ginger-breadboard       " & gLab.ProjName
        End With
        Proj_Set(gLab, gProjID)


    End Sub
    Private Sub Proj_Set(ByRef Lab As Lab_Type, ProjID As Long)
        ' '--------------------------
        Dim Ds As DataSet


        gLab.ProjID = gProjID

        Ds = DBAcs_Get_DataS(gLab, "Proj", "Owner,OwnerID", "ID=" & ProjID.ToString)
        If Ds.Tables(0).Rows.Count = 0 Then Return

        With Ds.Tables(0).Rows(0)
            If .Item(0) Is DBNull.Value Then
            Else
                Lab.Owner = Trim(.Item(0))
            End If
            Lab.OwnerID = Val(.Item(1))
        End With


        Proj_Show_Element(Lab, ProjID)
        Desk_ShowByProjID(Lab, ProjID)
    End Sub

    Private Sub Desk_ShowByProjID(ByRef Lab As Lab_Type, ProjID As Long)
        'デスクの表示
        '------------------------------

        ' Lab.view .SelectedItems(0).SubItems(4).Text
        Dim DS As DataSet
        Dim cnd As String
        cnd = "ProjID=" & ProjID.ToString
        cnd &= " AND State  > 0 "
        DS = DBAcs_Get_DataS(Lab, "Desk", "*", cnd)
        '.DBACS_Field.Add(Key, "ID,OwnerID,ProjID,Page,Type,Title")

        If DS.Tables(0).Rows.Count = 0 Then Return
        ReDim Lab.Desk(DS.Tables(0).Rows.Count)
        Lab.DeskAT.View.TabPages.Clear()
        For L = 0 To DS.Tables(0).Rows.Count - 1
            With Lab.Desk(L)
                .ID = Val(DS.Tables(0).Rows(L).Item(0))
                .OwnerID = Val(DS.Tables(0).Rows(L).Item(1))
                .ProjID = Val(DS.Tables(0).Rows(L).Item(2))
                .DeskNo = Val(DS.Tables(0).Rows(L).Item(3))
                .Type = Trim(DS.Tables(0).Rows(L).Item(4))
                .Text = Trim(DS.Tables(0).Rows(L).Item(5))
                .State = Trim(DS.Tables(0).Rows(L).Item(6))
                .ShiftX = Val(DS.Tables(0).Rows(L).Item(7))
                .ShiftY = Val(DS.Tables(0).Rows(L).Item(8))
                .ZM = Val(DS.Tables(0).Rows(L).Item(9))
                'デスクの表示
                .Tab = New TabPage()
                .Tab.Tag = .Type & "|" & .ID.ToString

                .Tab.Text = .Text
                .Tab.Name = .ID.ToString
                .Tab.AllowDrop = True

                Lab.DeskName = .ID.ToString
                Lab.TabID = .Tab.TabIndex

                AddHandler .Tab.DragOver, AddressOf Desk_DragOver
                AddHandler .Tab.DragDrop, AddressOf Desk_DragDrop
                AddHandler .Tab.MouseDown, AddressOf Desk_MouseDown
                AddHandler .Tab.MouseMove, AddressOf Desk_MouseMove
                AddHandler .Tab.MouseUp, AddressOf Desk_MouseUp

                AddHandler .Tab.MouseDoubleClick, AddressOf Desk_DoubleClick
                AddHandler .Tab.SizeChanged, AddressOf Desk_ReSize
                Lab.DeskAT.View.TabPages.Add(.Tab)

                'BOXの表示

                Box_Set_OnDESK(Lab, .ID)

            End With
        Next

        '本来なら以前の作業中のページ
        With Lab

            .TabID = 0
            .DeskSx = .Desk(.TabID).ShiftX
            .DeskSy = .Desk(.TabID).ShiftY
            .DeskZm = .Desk(.TabID).ZM
        End With
        Lab.DeskAT.View.SelectedIndex = Lab.TabID
        ' Lab.DeskID = Val(Lab.DeskAT.View.TabPages(1).name)



    End Sub
    Private Sub Bx_Re_Draw(ByRef Lab As Lab_Type, ByRef Drbox As AxMSINKAUTLib.AxInkPicture)
        ''イメージデータを保存

        ''Dim Bs() As Byte
        ''Bs = Drbox.Ink.Save(1)
        ''DBACS_Update_Byte(Lab, "Draw", "Text", "ID =" & Drbox.name, Bs)

        ''BASE64
        'Dim Bs As String
        'Bs = Drbox.Ink.Save(1)
        'DBAcs_Update(Lab, "Draw", "Text = '" & Drbox.Name & "'", "ID =" & Bs)


    End Sub
    Private Sub Bx_Re_Text(ByRef Lab As Lab_Type, TboxID As Long, ByRef Tdata As String)
        'テキストを保存
        DBAcs_Update(Lab, "Tbox", "Text = '" & Tdata & "'", "ID = " & TboxID.ToString)
    End Sub

    Private Sub Bx_Re_Pict(ByRef Lab As Lab_Type, TAG As String, path As String)
        'イメージデータのパスを保存
        DBAcs_Update(Lab, "Pbox", "PICT = '" & path & "'", "ID = " & TAG)


    End Sub

    Private Sub Bx_Re_Pict2(ByRef Lab As Lab_Type, ByRef PID As String, ByRef File As String)
        'イメージデータを保存

        ' ファイルを読み込む
        Dim Fs As New System.IO.FileStream(File, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        'ファイルを読み込むバイト型配列を作成する
        Dim Bs(Fs.Length - 1) As Byte
        Fs.Read(Bs, 0, Bs.Length)
        Fs.Close()

        'Dim BData() As Byte イメージデータのコンバータ
        'BData = ImageToByteArray(bs)

        DBACS_Update_Byte(Lab, "BDat", "Dat", "ID =" & PID, Bs)


    End Sub

    Private Sub DBACS_Update_Byte(ByRef Lab As Lab_Type, Key As String, Act As String, Cnd As String, ByRef BData() As Byte)
        'バイト型配列の保存
        '-----------------------
        'Dim strSQL As String
        'strSQL = ""
        'strSQL &= "UPDATE TBL_" & Key & " SET "
        'strSQL &= " Dat = @pic"
        'strSQL &= " WHERE " & Cnd


        ''    登録処理
        'Select Case Lab.Db_Type
        '    Case "SQLite"
        '    Case "MSSQL"
        '        Using con As New SqlClient.SqlConnection(Lab.DbPath & ";Initial Catalog =" & Lab.Db_Type & Key)
        '            Dim cmd As New SqlClient.SqlCommand(strSQL, con)
        '            cmd.Parameters.Add("@pic", SqlDbType.Binary, BData.Length).Value = BData
        '            con.Open()
        '            For I As Integer = 1 To 50
        '                Try
        '                    cmd.ExecuteNonQuery()
        '                    Exit For
        '                Catch ex As Exception
        '                    Log_Show(1, "DBエラー" & I & ex.Message)
        '                    For J As Integer = 1 To 100
        '                        System.Windows.Forms.Application.DoEvents()
        '                    Next
        '                End Try
        '            Next
        '            con.Close()
        '        End Using
        'End Select



    End Sub

    Private Sub DBACS_Update_ByteMSSQL(ByRef Lab As Lab_Type, Key As String, Act As String, Cnd As String, ByRef BData() As Byte)
        'バイト型配列の保存
        '-----------------------
        Dim strSQL As String
        strSQL = ""
        strSQL &= "UPDATE TBL_" & Key & " SET "
        strSQL &= " Dat = @pic"
        strSQL &= " WHERE " & Cnd


        '    登録処理
        Using con As New SqlClient.SqlConnection(Lab.DbPath & ";Initial Catalog =" & Lab.Db_Type & Key)
            Dim cmd As New SqlClient.SqlCommand(strSQL, con)
            cmd.Parameters.Add("@pic", SqlDbType.Binary, BData.Length).Value = BData
            con.Open()
            For I As Integer = 1 To 50
                Try
                    cmd.ExecuteNonQuery()
                    Exit For
                Catch ex As Exception
                    Log_Show(1, "DBエラー" & I & ex.Message)
                    For J As Integer = 1 To 100
                        System.Windows.Forms.Application.DoEvents()
                    Next
                End Try
            Next
            con.Close()
        End Using
    End Sub





    Private Sub Bx_Get_Draw(ByRef Lab As Lab_Type, ByRef Draw As AxMSINKAUTLib.AxInkPicture, DrawID As Long)
        Dim imgconv As New ImageConverter()
        Dim DS As DataSet
        DS = DBAcs_Get_DataS(Lab, "Draw", "Text", "ID = " & DrawID.ToString)
        If DS.Tables(0).Rows.Count = 0 Then Return

        ' Draw.Ink.Load(CType(imgconv.ConvertFrom(DS.Tables(0).Rows(0).Item(0)), Image))
        ' Draw.Ink.Load(DS.Tables(0).Rows(0).Item(0), isf))
        'ConvertFrom(Context As ITypeDescriptorContext, culture As CultureInfo, value As Object)
        'With Draw
        '    .Ink.Dispose()
        '    .Ink = New Microsoft.Ink.Ink()
        '    .Ink.Load(DS.Tables(0).Rows(0).Item(0))
        '    .InkEnabled = True
        '    .Refresh()
        'End With

        ''If DS.Tables(0).Rows(0).Item(0) = Nothing Then Return
        'Dim bs() As Byte = System.Convert.FromBase64String(DS.Tables(0).Rows(0).Item(0))
        'Draw.Strokes = New StrokeCollection(MS)
        '' Draw.Ink.Strokes.Cast(Of Base64FormattingOptions) = a
    End Sub


    Private Sub Bx_Get_Pict(ByRef Lab As Lab_Type, ByRef Pict As PictureBox, PictID As Long)
        'イメージデータをパスから読み出し
        Dim DS As DataSet
        DS = DBAcs_Get_DataS(Lab, "Pbox", "Dat", "ID = " & PictID.ToString)
        If DS.Tables(0).Rows.Count = 0 Then Return
        Pict.ImageLocation = DS.Tables(0).Rows(0).Item(0)

    End Sub


    Private Sub Bx_Get_Pict2(ByRef Lab As Lab_Type, ByRef Pict As PictureBox, PictID As Long)
        Dim imgconv As New ImageConverter()
        Dim DS As DataSet
        DS = DBAcs_Get_DataS(Lab, "BDat", "Dat", "ID = " & PictID.ToString)
        If DS.Tables(0).Rows.Count = 0 Then Return
        If DS.Tables(0).Rows(0).Item(0) Is DBNull.Value Then Return
        'If DS.Tables(0).Row.isnull("Dat") Then Return
        ' Dim img As System.Drawing.Image = System.Drawing.Image.FromFile(files(0))

        ' Pict.Image = ByteArrayToImage(DS.Tables(0).Rows(0).Item(0))

        Pict.Image = CType(imgconv.ConvertFrom(DS.Tables(0).Rows(0).Item(0)), Image)
    End Sub
    Private Function Bx_get_ByteDATA(ByRef Lab As Lab_Type, DbKey As String, Act As String, cnd As String) As Byte()
        'イメージデータをDBから読み出し
        'Dim DS As DataSet
        ''Dim bData As Byte()
        ''DS = DBAcs_Get_DataS(Lab, "Pbox", "Dat", "ID = " & PictID.ToString)

        ''PICT_GET(Lab.DbPath, "Pbox", "Dat", "ID = " & PictID.ToString)



        'Dim strSQL As String
        'strSQL = " SELECT "" Pict"
        'strSQL &= vbCrLf & " FROM  TBL_Pbox2"
        'strSQL &= vbCrLf & " WHERE  ID =" & PictID


        '''    読み出し処理
        'Using con As New SqlClient.SqlConnection(Lab.DbPath & ";Initial Catalog =" & "DB_Pbox2")
        '    Dim cmd As New SqlClient.SqlCommand(strSQL, con)
        '    Dim objRs As SqlDataReader

        '    con.Open()


        '    'objRs = cmd.ExecuteReader()
        '    'bData = CType(objRs.Item(0), Byte())
        '    ''' Dim imageData As Byte() = DirectCast(cmd.ExecuteScalar(), Byte())
        '    ''Dim imagedata() As Byte = CType(cmd.ExecuteScalar(), Byte())
        '    ''' Pict.Image = ByteArrayToImage(Reader)
        '    For I As Integer = 1 To 50
        '        Try
        '            Dim imageData As Image = cmd.ExecuteScalar()


        '            Exit For
        '        Catch ex As Exception
        '            Log_Show(1, "DBエラー" & I & ex.Message)
        '            For J As Integer = 1 To 100
        '                System.Windows.Forms.Application.DoEvents()
        '            Next
        '        End Try
        '    Next
        '    con.Close()
        'End Using



    End Function

    'Private Function PICT_GET(ByRef DB_Path As String, Name As String, Action As String, ByRef Condition As String) As DataSet

    '    'DB 指定Key の　値　を格納
    '    '--------------------------
    '    Dim tmpR As Integer

    '    Dim DBCon As SqlConnection
    '    Dim TmpCmd As String
    '    tmpR = -1

    '    TmpCmd = "Select " & Action & " FROM [" & .dbname & Name & "].[dbo].[" & "TBL_" & Name & "] WHERE " & Condition
    '    ' TmpCmd = "Select * FROM [" & .DB_Name & "].[dbo].[" & .TBL_Name & "] "
    '    DBCon = New SqlConnection(DB_Path & ";Initial Catalog =" & .dbname & Name)
    '    DBCon.Open()
    '    Dim cmd As New SqlClient.SqlCommand(TmpCmd, DBCon)

    '    Dim objRs As SqlDataReader
    '    objRs = cmd.ExecuteReader()
    '    Dim bData As Byte()
    '    bData = CType(objRs.Item(0), Byte())
    'End Function


    Public Shared Function ByteArrayToImage(ByVal b As Byte()) As Image
        ' バイト配列をImageオブジェクトに変換
        Dim imgconv As New ImageConverter()
        Dim img As Image = CType(imgconv.ConvertFrom(b), Image)
        Return img
    End Function


    Public Shared Function ImageToByteArray(ByVal img As Image) As Byte()
        ' Imageオブジェクトをバイト配列に変換
        Dim imgconv As New ImageConverter()
        Dim b As Byte() = CType(imgconv.ConvertTo(img, GetType(Byte())), Byte())
        Return b
    End Function

    Private Sub ListView2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs)

        With ListView2
            '            Dim Target_Item As ListViewItem
            '削除
            If e.KeyData = Keys.Delete Then
                ListView_Del(ListView2)　'
            End If

            '作成
            If e.KeyData = Keys.Insert Then
                Proj_Atr_New_Item(ListView2, "S")
            End If

        End With

    End Sub


    Private Sub ListView_Del(ByRef LV As ListView)
        'projectの削除
        '-----------------

        With LV
            If IsNothing(.SelectedItems) Then Return
            If .SelectedItems.Count = 0 Then Return

            '  'Histroy_start(gLab.Desk(1).history)
            For L = .SelectedItems.Count - 1 To 0 Step -1
                .SelectedItems(L).Remove()
            Next

            If Not (.FocusedItem Is Nothing) Then
                .FocusedItem.Selected = True
            End If


        End With

    End Sub



    Private Sub ListView2_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs)


        'Select Case e.Button
        '    Case Windows.Forms.MouseButtons.Right
        '        gBff_Act_Tree = Nothing
        '        gBff_Act_List = ListView2
        'End Select




    End Sub







    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

        'Dim menu As ContextMenuStrip = DirectCast(sender, ContextMenuStrip)
        ''ContextMenuStripを表示しているコントロールを取得する
        'Dim source As Control = menu.SourceControl
        'If source IsNot Nothing Then
        '    gBff_ObjName = source.Name
        'End If


    End Sub

    Private Sub ContextMenuStrip2_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs)

    End Sub
    Private Sub Line_MouseEnter(sender As Object, e As System.EventArgs)
        Dim Panel As Panel = CType(sender, Panel)

        Panel.BorderStyle = BorderStyle.FixedSingle
        Panel.BringToFront()
        gLab.LineID = Val(sender.name)

        gLab.Box = sender
        gLab.Flg_BoxMouse = "OnLINE"
        Panel.Height = 12
        Dim Tag As String = "OnLINE." & gLab.LineID.ToString ' gLab.Line.Name

        With TabControl3

            For Each Box As Panel In .TabPages(.SelectedIndex).Controls
                If Box.Tag = Tag Then
                    Box.BringToFront()
                End If
            Next
        End With

    End Sub
    Private Sub Line_MouseLeave(sender As Object, e As System.EventArgs)

        Dim Panel As Panel = CType(sender, Panel)
        gLab.LineID = 0
        Panel.BorderStyle = BorderStyle.None

        Panel.Height = 6
        gLab.Flg_BoxMouse = ""
    End Sub
    Private Sub Base_MouseLeave(sender As Object, e As System.EventArgs)
        sender.BorderStyle = BorderStyle.None

        Me.Cursor = Cursors.Default


    End Sub
    Private Sub Base_MouseEnter(sender As Object, e As System.EventArgs)
        sender.BorderStyle = BorderStyle.FixedSingle
        If gLab.BoxID <> Val(sender.name) Then sender.BringToFront()
        gLab.BoxID = Val(sender.name)
        gLab.Box = sender
        gLab.Flg_BoxMouse = "OnBOX"

    End Sub


    'Private Sub TreeView2_MouseEnter(sender As Object, e As System.EventArgs)

    '    ' Menu_SetType(1)

    'End Sub

    'Private Sub TreeView2_MouseLeave(sender As Object, e As System.EventArgs)
    '    ' Menu_SetType(0)
    'End Sub

    Private Sub TextBox8_TextChanged(sender As System.Object, e As System.EventArgs)

    End Sub


    'Private Sub Desk_Kill_Child_All(ByRef DB As DBACS_Type, ByRef PID() As Integer)
    '    '配下を削除DB中のみ
    '    Dim CID() As Integer
    '    If PID Is Nothing Then Return
    '    For L = 0 To PID.Length - 1

    '        If PID(L) > 0 Then

    '            CID = DBACS_Get_ID_ByCnd(DB, "PID = " & PID(L))
    '            Desk_Kill_Child_All(DB, CID)
    '            DBAcs_Del_Dat(DB, "ID = " & PID(L))
    '            Log_Show(2, PID(L))

    '        End If


    '    Next



    'End Sub


    Private Sub Proj_Set_Env(ByRef Lab As Lab_Type)
        'Projの基本パラメータを設定　　（データの読み込みはしない
        '-----------------------------
        'Log_Show(1, "Projの基本パラメータを設定")
        'With Lab

        '    ''DB作成
        '    '--------------------------
        '    DBAcs_Init2(.ProjAT.DBACS, "RepoDB", "Proj", "Lab_Key")


        'End With

    End Sub

    Private Sub Desk_Set_Env(ByRef Lab As Lab_Type)

        'Desk設定
        '--------------
        Dim DeskID As Integer
        DeskID = 1

        'With Lab.Desk(DeskID)
        '    'TreeView
        '    '--------------------------
        '    '.Tree = Me.TreeView1
        '    '.Tree.ImageList = gImg_List.Image

        '    'Listview
        '    '--------------------------
        '    '.View = ListView1
        '    'ReDim .Columns(3)
        '    'Columns_Set_Env(.Columns(0), "[Proj]")
        '    '' Columns_Set_Env(.Columns(1), "[Branch]")
        '    'Columns_Set_Listview(.View, .Columns(0))


        '    ''Desk_DB作成
        '    '--------------------------
        '    DBAcs_Init2(.DBACS, "Desk", "Node", "Desk_Node")
        '    'Dim Ds As DataSet
        '    'Dim Cnd As String
        '    'Cnd = "OwnerID = " & Lab.OwnerID.ToString

        '    'Ds = DBAcs_Get_DataS(.DBACS, "ID", Cnd)

        '    'If Ds.Tables(0).Rows.Count = 0 Then  'TBLがからっぽなら、ＤＢに、デフォルト登録

        '    '    .Node_Trash = Desk_Make_Node(Lab, "[Trash]")
        '    '    .Node_Society = Desk_Make_Node(Lab, "[Society]")
        '    '    .Node_Closed = Desk_Make_Node(Lab, "[Closed]")
        '    '    .Node_Opened = Desk_Make_Node(Lab, "[Opened]")
        '    '    .Node_Current = Desk_Make_Node(Lab, "[Current]")

        '    'End If


        '    'If DB_Get_Data_ByKey(.DBACS, .DBACS.TBL_Name, "ID", 1) = "" Then


        '    'Else
        '    '    'DBからの読みだしはもっとあとでやる


        '    'End If



        'End With

    End Sub
    'Private Sub Desk_Read_NodeData(ByRef Lab As Lab_Type)

    '    '    'Desk設定
    '    '    '--------------
    '    '    Dim Ds As DataSet
    '    '    Dim Cnd As String

    '    '    Dim DeskID As Integer
    '    '    DeskID = 1
    '    '    Dim NID As Integer

    '    '    With Lab.Desk(DeskID)
    '    '        Cnd = "OwnerID = " & Lab.OwnerID.ToString
    '    '        Ds = DBAcs_Get_DataS(.DBACS, "*", Cnd)

    '    '        If Ds.Tables(0).Rows.Count = 0 Then  'TBLがからっぽなら、ＤＢに、デフォルト登録

    '    '            Return
    '    '        Else
    '    '            'DBから（自分の）全ノードデータを読み込む "ID,NID,PID,RID,State,NodeType,Name,Owner,OwnerID"
    '    '            For L = 0 To Ds.Tables(0).Rows.Count - 1

    '    '                NID = Val(Ds.Tables(0).Rows(L).Item(1))
    '    '                If NID > .NodeMax Then
    '    '                    .NodeMax = NID
    '    '                    ReDim Preserve Lab.Desk(DeskID).Node(.NodeMax)
    '    '                End If

    '    '                With Lab.Desk(DeskID).Node(NID)
    '    '                    .ID = Val(Ds.Tables(0).Rows(L).Item(0))
    '    '                    .PID = Val(Ds.Tables(0).Rows(L).Item(2))
    '    '                    .RID = Val(Ds.Tables(0).Rows(L).Item(3))
    '    '                    .State = Trim(Ds.Tables(0).Rows(L).Item(4))
    '    '                    .NodeType = Trim(Ds.Tables(0).Rows(L).Item(5))
    '    '                    .Name = Trim(Ds.Tables(0).Rows(L).Item(6))
    '    '                    .Owner = Trim(Ds.Tables(0).Rows(L).Item(7))
    '    '                    .OwnerID = Val(Ds.Tables(0).Rows(L).Item(8))
    '    '                End With
    '    '            Next
    '    '        End If
    '    '    End With

    'End Sub

    Private Sub Element_Set_Env(ByRef Lab As Lab_Type)
        'Elmの基本パラメータを設定　　（データの読み込みはしない
        '-----------------------------
        Log_Show(1, "Elementの基本パラメータを設定")

        With Lab.ElementAT.View
            .ListView = ListView2
            Columns_Set_Env(.Columns, "[Element]")
            Columns_Set_Listview(.ListView, .Columns)
        End With

    End Sub










    'Private Sub CheckedListBox1_DragOver(sender As Object, e As System.Windows.Forms.DragEventArgs)
    '    If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
    '        e.Effect = DragDropEffects.Copy
    '    Else
    '        e.Effect = DragDropEffects.None
    '    End If
    'End Sub

    Private Sub GrouPbox1_Enter(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SplitContainer9_Panel2_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub SplitContainer8_Panel2_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub SplitContainer13_SplitterMoved(sender As System.Object, e As System.Windows.Forms.SplitterEventArgs)

    End Sub

    Private Sub ToolStripLabel1_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub

    Private Sub ListView4_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Invalidated(sender As Object, e As InvalidateEventArgs) Handles Me.Invalidated

    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub

    Private Sub ContextMenuStrip1_MouseClick(sender As Object, e As MouseEventArgs) Handles ContextMenuStrip1.MouseClick

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ContextMenuStrip1_MouseDown(sender As Object, e As MouseEventArgs) Handles ContextMenuStrip1.MouseDown

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TreeView1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs)

    End Sub

    Private Sub ListView2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView2.MouseDoubleClick

        With ListView2
            If .SelectedItems.Count = 0 Then Return
            Element_Act_Item(gLab, .SelectedItems(0).Tag)
        End With
    End Sub





    Private Sub Button2_Click_3(sender As Object, e As EventArgs) Handles Button2.Click

        'OWNER プロファイル登録
        '-------------------


        'With gLab.Person(0)
        '    .Image = PictureBox1.ImageLocation
        '    .Name = TextBox4.Text
        '    .Tel = TextBox2.Text
        '    .Mail = TextBox3.Text
        '    .Prof = TextBox5.Text
        'End With

        Person_Set_NewData(gLab)
        Lab_Start(gLab)
        MsgBox("登録されました")
    End Sub

    Private Sub PictureBox1_DragDrop(sender As Object, e As DragEventArgs) Handles PictureBox1.DragDrop
        Dim strFileName As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        '  PictureBox1.Image = System.Drawing.Image.FromFile(strFileName(0))

        PictureBox1.ImageLocation = strFileName(0)

    End Sub

    Private Sub PictureBox1_DragEnter(sender As Object, e As DragEventArgs) Handles PictureBox1.DragEnter
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub PictureBox2_MouseEnter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
        'PictureBox2.BackColor =Color.FromName(gdeco.BackColor)
    End Sub


    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub

    Private Sub PictureBox3_MouseLeave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub

    Private Sub PictureBox3_MouseEnter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
    End Sub

    Private Sub PictureBox4_MouseEnter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
    End Sub

    Private Sub PictureBox4_MouseLeave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub



    Private Sub Lab_Act(Lab As Lab_Type, Key As String)

        'Select Case Key
        '    Case "研究ノート"
        '        TabControl1.SelectedIndex = 0


        '    Case "新研究"

        '        'RichTextBox2.Text = "新研究"

        '        'TabControl1.SelectedIndex = 1

        '    Case "私"
        '        TabControl1.SelectedIndex = 1
        '        'Case "研究"
        '        '    TabControl1.SelectedIndex = 1


        '        'Case "検索"
        '        '    TabControl1.SelectedIndex = 2
        'End Select



    End Sub

    Private Sub PictureBox2_MouseClick(sender As Object, e As MouseEventArgs)


        Lab_Act(gLab, sender.name)


    End Sub

    Private Sub PictureBox3_MouseClick(sender As Object, e As MouseEventArgs)

        Lab_Act(gLab, sender.name)
    End Sub

    Private Sub PictureBox6_MouseEnter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
    End Sub

    Private Sub PictureBox6_MouseLeave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim DS As DataSet
        Dim Cmd, Key As String
        Key = TextBox6.Text
        Key = Replace(Key, Chr(13), "")
        Key = Replace(Key, Chr(10), "")


        If Key = "" Or Key = "*" Then
            Cmd = " State > 0 "
            Cmd &= "AND OwnerID = " & gX_ownerID
            Proj_Show_ListByCmd(gLab, Cmd) 'オーナーのレポ表示

        Else
            '  Cmd = " State ='[Opened]' AND OwnerID <> " & gLab.OwnerID
            Cmd = " State = 1 "
            Cmd &= "AND Name LIKE '%" & Key & "%'"
            Proj_Show_ListByCmd(gLab, Cmd)
        End If

        '  ListView1.Visible = True

    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs)
        'WebView21.Visible = True
        ''WebView21.Dock = "fill"

        'ListView1.Visible = False
        '' ListView4.Dock = "non"
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs)
        Lab_Act(gLab, sender.name)
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)
        Lab_Act(gLab, sender.name)
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs)
        Lab_Act(gLab, sender.name)
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs)
        Lab_Act(gLab, sender.name)
    End Sub
    Private Sub Lab_Selected_TabPage(ByRef Lab As Lab_Type, Key As String)

        'Select Case Key
        '    Case "私"
        '        TabControl1.SelectedIndex = 1
        '    Case "研究"
        '        WebView21.Visible = True
        '        ListView1.Visible = False

        '    Case "検索"
        '        WebView21.Visible = False
        '        ListView1.Visible = True


        'End Select



    End Sub



    Private Sub PictureBox6_Leave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub

    Private Sub PictureBox6_Enter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
    End Sub

    Private Sub PictureBox5_MouseEnter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
    End Sub

    Private Sub PictureBox5_MouseLeave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub

    Private Sub PictureBox5_Click_1(sender As Object, e As EventArgs)
        Lab_Act(gLab, sender.name)
    End Sub

    Private Sub ListView1_MouseDown(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDown

    End Sub

    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick

        With ListView1
            If .SelectedItems Is Nothing Then Return
            If .SelectedItems.Count <= 0 Then Return

            If .SelectedItems.Count > 1 Then Return

            If .SelectedItems(0).Tag Is Nothing Then Return

            gProjID = .SelectedItems(0).Tag
            gLab.ProjID = gProjID
            ' TabControl1.SelectedIndex = 1

        End With
        'Proj_Show_Data(gLab, gProjID)

    End Sub

    Private Async Sub InitializeAsync()
        Await WebView21.EnsureCoreWebView2Async(Nothing)
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click


        '  WebView21.Source = New Uri("https://freecalend.com/open/mem179529/index")
        '  WebView21.Source = New Uri("C:\Users\miura\Documents\ＪＳＡライフ２０２１.pdf")
        'WebView21.Source = New Uri("file:///C:/Users/miura/Documents/XXtest_.html")
        '   WebView21.Source = New Uri("https://query.wikidata.org/sparql?query=SELECT%20DISTINCT%20%20%3Fitem%20%3FitemLabel%20%3FgenderLabel%20%3FcountryLabel%20%3FbirthdateLabel%20%3FbirthplaceLabel%20%3FjobLabel%20%3FemployerLabel%20%3FinstLabel%20WHERE%20%7B%0A%20%20SERVICE%20wikibase%3Alabel%20%7B%20bd%3AserviceParam%20wikibase%3Alanguage%20%22ja%22.%20%7D%0A%20%20%7B%0A%20%20%20%20SELECT%20DISTINCT%20%20*%20WHERE%20%7B%0A%20%20%20%20%20%20%3Fitem%20rdfs%3Alabel%20%22%E4%BD%90%E5%80%89%E7%9C%9F%E8%A1%A3%22%40ja%20.%0A%20%20%20%20%20%20%23%3Fitem%20rdfs%3Alabel%20%22%E6%A1%9C%E4%BA%95%E6%B4%8B%E5%AD%90%22%40ja%20.%0A%20%20%20%20%20%20%3Fitem%20wdt%3AP21%20%3Fgender.%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP27%20%3Fcountry.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP569%20%3Fbirthdate.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP19%20%3Fbirthplace.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP106%20%3Fjob.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP108%20%3Femployer.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP1303%20%3Finst.%7D%0A%20%20%20%20%20%0A%20%20%20%20%7D%0A%20%20%20%20LIMIT%201000%0A%20%20%7D%0A%7D")
        Dim tmpURL, TmpKey As String
        tmpURL = "https://query.wikidata.org/embed.html#SELECT%20DISTINCT%20%20%3Fitem%20%3FitemLabel%20%3FgenderLabel%20%3FcountryLabel%20%3FbirthdateLabel%20%3FbirthplaceLabel%20%3FjobLabel%20%3FemployerLabel%20%3FinstLabel%20WHERE%20%7B%0A%20%20SERVICE%20wikibase%3Alabel%20%7B%20bd%3AserviceParam%20wikibase%3Alanguage%20%22ja%22.%20%7D%0A%20%20%7B%0A%20%20%20%20SELECT%20DISTINCT%20%20*%20WHERE%20%7B%0A%20%20%20%20%20%20%3Fitem%20rdfs%3Alabel%20%22%40%40JINMEI%40%40%22%40ja%20.%0A%20%20%20%20%20%20%23%3Fitem%20rdfs%3Alabel%20%22%E6%A1%9C%E4%BA%95%E6%B4%8B%E5%AD%90%22%40ja%20.%0A%20%20%20%20%20%20%3Fitem%20wdt%3AP21%20%3Fgender.%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP27%20%3Fcountry.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP569%20%3Fbirthdate.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP19%20%3Fbirthplace.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP106%20%3Fjob.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP108%20%3Femployer.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP1303%20%3Finst.%7D%0A%20%20%20%20%20%0A%20%20%20%20%7D%0A%20%20%20%20LIMIT%201000%0A%20%20%7D%0A%7D"
        TmpKey = System.Web.HttpUtility.UrlEncode("岩倉具視")
        tmpURL = Replace(tmpURL, "%40%40JINMEI%40%40", TmpKey)
        WebView21.CoreWebView2.Navigate(tmpURL)
        ' WebView21.CoreWebView2.Navigate("https://query.wikidata.org/embed.html#SELECT%20DISTINCT%20%20%3Fitem%20%3FitemLabel%20%3FgenderLabel%20%3FcountryLabel%20%3FbirthdateLabel%20%3FbirthplaceLabel%20%3FjobLabel%20%3FemployerLabel%20%3FinstLabel%20WHERE%20%7B%0A%20%20SERVICE%20wikibase%3Alabel%20%7B%20bd%3AserviceParam%20wikibase%3Alanguage%20%22ja%22.%20%7D%0A%20%20%7B%0A%20%20%20%20SELECT%20DISTINCT%20%20*%20WHERE%20%7B%0A%20%20%20%20%20%20%3Fitem%20rdfs%3Alabel%20%22%E4%BD%90%E5%80%89%E7%9C%9F%E8%A1%A3%22%40ja%20.%0A%20%20%20%20%20%20%23%3Fitem%20rdfs%3Alabel%20%22%E6%A1%9C%E4%BA%95%E6%B4%8B%E5%AD%90%22%40ja%20.%0A%20%20%20%20%20%20%3Fitem%20wdt%3AP21%20%3Fgender.%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP27%20%3Fcountry.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP569%20%3Fbirthdate.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP19%20%3Fbirthplace.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP106%20%3Fjob.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP108%20%3Femployer.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP1303%20%3Finst.%7D%0A%20%20%20%20%20%0A%20%20%20%20%7D%0A%20%20%20%20LIMIT%201000%0A%20%20%7D%0A%7D")

        ' WebView21.CoreWebView2.Navigate("https://www.microsoft.com/ja-jp")

        'Dim html = Await WebView21.ExecuteScriptAsync("document.documentElement.outerHTML")

        ' html = Regex.Unescape(html)

        'TextBox7.Text = html.ToString

        'Dim sUrl As String   '// ダウンロード対象ファイルのURL
        'Dim sDir As String   '// ダウンロードファイルを保存するローカルPCのフォルダパス

        ''// URL設定
        'sUrl = = "https://query.wikidata.org/embed.html#SELECT%20DISTINCT%20%20%3Fitem%20%3FitemLabel%20%3FgenderLabel%20%3FcountryLabel%20%3FbirthdateLabel%20%3FbirthplaceLabel%20%3FjobLabel%20%3FemployerLabel%20%3FinstLabel%20WHERE%20%7B%0A%20%20SERVICE%20wikibase%3Alabel%20%7B%20bd%3AserviceParam%20wikibase%3Alanguage%20%22ja%22.%20%7D%0A%20%20%7B%0A%20%20%20%20SELECT%20DISTINCT%20%20*%20WHERE%20%7B%0A%20%20%20%20%20%20%3Fitem%20rdfs%3Alabel%20%22%E4%BD%90%E5%80%89%E7%9C%9F%E8%A1%A3%22%40ja%20.%0A%20%20%20%20%20%20%23%3Fitem%20rdfs%3Alabel%20%22%E6%A1%9C%E4%BA%95%E6%B4%8B%E5%AD%90%22%40ja%20.%0A%20%20%20%20%20%20%3Fitem%20wdt%3AP21%20%3Fgender.%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP27%20%3Fcountry.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP569%20%3Fbirthdate.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP19%20%3Fbirthplace.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP106%20%3Fjob.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP108%20%3Femployer.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP1303%20%3Finst.%7D%0A%20%20%20%20%20%0A%20%20%20%20%7D%0A%20%20%20%20LIMIT%201000%0A%20%20%7D%0A%7D"
        'sDir = "C:\Users\miura\Downloads\Test"

        ''// ダウンロード
        'URLDownloadToFile(sUrl, sDir)


    End Sub




    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ListView2_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles ListView2.ItemSelectionChanged

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim gifData() As Byte
        gifData = AxInkPicture1.Ink.Save(2)

        Dim x As Integer

        x = 85
        With gImg_List
            TextBox7.Text = .Key(x)
            PictureBox9.Image = .Image.Images(x - 1)
        End With

    End Sub

    Private Sub TabPage7_Click(sender As Object, e As EventArgs) Handles TabPage7.Click

    End Sub

    Private Sub TabPage7_DragDrop(sender As Object, e As DragEventArgs) Handles TabPage7.DragDrop




    End Sub

    Private Sub Panel1_DragDrop(sender As Object, e As DragEventArgs)
        Desk_Drops(sender, e)



    End Sub
    Private Sub Desk_DoubleClick(sender As Object, e As MouseEventArgs)
        'BOXを全部元位置に
        Dim SX As Integer = 10
        With gLab
            .DeskID = Val(sender.name)

            For Each box As Panel In sender.Controls
                Select Case box.Tag
                    Case "LINE"
                    Case "OnDESK"

                        '拡大縮小を元に
                        For L = 1 To Math.Abs(.DeskZm)
                            box.Left = Box_Move_ZM(box.Left, .Desk(.TabID).ShiftX, .DeskZm, SX)
                            box.Top = Box_Move_ZM(box.Top, .Desk(.TabID).ShiftY, .DeskZm, SX)
                            box.Width = Box_Re_ZM(box.Width, .DeskZm, SX)
                            box.Height = Box_Re_ZM(box.Height, .DeskZm, SX)
                        Next

                        box.Left = box.Left - .Desk(.TabID).ShiftX
                        box.Top = box.Top - .Desk(.TabID).ShiftY
                        Box_Re_Move(gLab, Val(box.Name), box)
                        Box_Re_Size(gLab, Val(box.Name), box.Width, box.Height)
                End Select
            Next
            .DeskZm = 0
            Dim Dat As String
            With .Desk(.TabID)
                .ShiftX = 0
                .ShiftY = 0
                .ZM = 0

                Dat = " "
                Dat &= "ShiftX = " & .ShiftX & ","
                Dat &= "ShiftY = " & .ShiftY & ","
                Dat &= "ZM = " & .ZM
                Map_View(gLab, .ShiftX, .ShiftY, .ZM)
            End With
            DBAcs_Update(gLab, "Desk", Dat, "ID = " & .DeskID)
        End With
        '   Map_View(gLab, XX, YY)
        'リフレッシュ
        Me.Refresh()

    End Sub

    Private Sub Desk_MouseDown(sender As Object, e As MouseEventArgs)
        Log_Show(1, "X=" & e.X & "  Y=" & e.Y & "|" & gLab.DeskSx.ToString & " |" & gLab.DeskSy.ToString)
        If e.Button = MouseButtons.Left Then
            '移動前の位置を記録
            gBase_StartX = e.X
            gBase_StartY = e.Y
        End If
        gLab.Flg_BoxMouse = "OnDESK"
        Me.Cursor = Cursors.Hand

    End Sub
    Private Sub Line_DragDrop(sender As Object, e As DragEventArgs)
        Desk_Drops(sender, e) 'Stackぽっとん
    End Sub
    Private Sub Line_DragOver(sender As Object, e As DragEventArgs)
        Log_Show(1, e.X & "|" & e.Y)
        gLab.ElementAT.View.ListView.Focus()
        e.Effect = DragDropEffects.Move

    End Sub
    Private Sub Desk_DragDrop(sender As Object, e As DragEventArgs)
        Desk_Drops(sender, e) 'Stackぽっとん
    End Sub

    Private Sub Desk_DragOver(sender As Object, e As DragEventArgs)

        gLab.ElementAT.View.ListView.Focus()
        e.Effect = DragDropEffects.Move

    End Sub

    Private Sub WebView21_Click(sender As Object, e As EventArgs) Handles WebView21.Click

    End Sub
    Dim gBase_StartX As Integer
    Dim gBase_StartY As Integer
    Dim gFlag_Pull As String
    Dim gBase_W As Integer
    Dim gBase_H As Integer

    Private Sub Base_DoubleClick(sender As Object, e As MouseEventArgs)
        If sender.dock = DockStyle.None Then
            sender.dock = DockStyle.Fill
        Else
            sender.dock = DockStyle.None
        End If
    End Sub
    Private Sub Line_MouseDown(sender As Object, e As MouseEventArgs)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then
            '移動前の位置を記録
            gBase_StartX = e.X
            gBase_StartY = e.Y
            gBase_W = sender.Width
            gBase_H = sender.Height
            If (sender.Width - gBase_StartX) < 8 And (sender.Height - gBase_StartY) < 8 Then
                gFlag_Pull = "Resize"
            Else
                gFlag_Pull = "Move"
            End If

        End If
        Me.Cursor = Cursors.HSplit
    End Sub
    Private Sub Base_MouseDown(sender As Object, e As MouseEventArgs)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then
            '移動前の位置を記録
            gBase_StartX = e.X
            gBase_StartY = e.Y
            gBase_W = sender.Width
            gBase_H = sender.Height
            Dim tmpP As String
            tmpP = ""
            tmpP &= gBase_StartX & "|"
            tmpP &= gBase_StartY & "|"
            tmpP &= sender.left & "|"
            tmpP &= sender.top & "|"
            tmpP &= gBase_W & "|"
            tmpP &= gBase_H & "|"
            Log_Show(1, tmpP)

            gFlag_Pull = ""

            If (sender.Width - gBase_StartX) < 10 And (sender.Height - gBase_StartY) < 10 Then
                gFlag_Pull = "Resize"
                Me.Cursor = Cursors.SizeNWSE
            End If

            If (gBase_StartX) < 10 And (sender.Height - gBase_StartY) < 10 Then
                gFlag_Pull = "Resize"
                Me.Cursor = Cursors.SizeNESW
            End If

            If (sender.Width - gBase_StartX) < 10 And Math.Abs((sender.Height) / 2 - gBase_StartY) < 10 Then
                gFlag_Pull = "Resize"
                Me.Cursor = Cursors.SizeWE
            End If

            If Math.Abs((sender.Width) / 2 - gBase_StartX) < 10 And (sender.Height - gBase_StartY) < 10 Then
                gFlag_Pull = "Resize"
                Me.Cursor = Cursors.SizeNS
            End If


            If (gBase_StartX) < 10 And Math.Abs((sender.Height) / 2 - gBase_StartY) < 10 Then
                gFlag_Pull = "Resize"

                Me.Cursor = Cursors.SizeWE
            End If


            If gFlag_Pull = "" Then
                gFlag_Pull = "Move"
                Me.Cursor = Cursors.Hand
            End If
        End If
    End Sub
    Private Sub Desk_MouseUp(sender As Object, e As MouseEventArgs)
        With gLab
            .DeskID = Val(sender.name)
            '  If gLab.Flg_DeskMove = 1 Then
            For Each box As Panel In sender.Controls
                If box.Tag <> "LINE" Then Box_Re_Move(gLab, Val(box.Name), box)
            Next

            '  End If

            'DESKのSHIFT 
            Dim Dat As String
            With .Desk(.TabID)
                Dat = " "
                Dat &= "ShiftX = " & .ShiftX & ","
                Dat &= "ShiftY = " & .ShiftY & ","
                Dat &= "ZM = " & gLab.DeskZm
                .ZM = gLab.DeskZm
                Map_View(gLab, .ShiftX, .ShiftY, .ZM)
            End With
            DBAcs_Update(gLab, "Desk", Dat, "ID = " & .DeskID)


        End With

        gLab.Flg_DeskMove = 0
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Desk_MouseMove(sender As Object, e As MouseEventArgs)
        Dim XX, YY As Integer
        '左クリックの場合
        ' Log_Show(1, "X=" & e.X & "  Y=" & e.Y & "|" & gLab.DeskSx.ToString & " |" & gLab.DeskSy.ToString)
        If e.Button = MouseButtons.Left Then
            With gLab


                '微小な変化を削除
                XX = e.X - gBase_StartX
                YY = e.Y - gBase_StartY
                If Math.Abs(XX) < 1 And Math.Abs(YY) < 1 Then Return

                .Flg_DeskMove = 1
                .Desk(.TabID).ShiftX = .Desk(.TabID).ShiftX + XX
                .Desk(.TabID).ShiftY = .Desk(.TabID).ShiftY + YY

                gBase_StartX = e.X
                gBase_StartY = e.Y

                Map_View(gLab, .Desk(.TabID).ShiftX, .Desk(.TabID).ShiftY, .Desk(.TabID).ZM)
                ' Log_Show(1, "X=" & e.X & "  Y=" & e.Y & "|" & gLab.DeskSx.ToString & " |" & gLab.DeskSy.ToString)
            End With

            'コントロール取得
            With sender
                For Each box As Panel In sender.Controls
                    Select Case box.Tag
                        Case "LINE"

                        Case "OnDESK"
                            box.Left = box.Left + XX
                            box.Top = box.Top + YY
                    End Select
                Next


            End With



            'リフレッシュ
            Me.Refresh()


        End If
    End Sub
    Private Sub Map_View(ByRef lab As Lab_Type, SX As Integer, SY As Integer, ZM As Integer)
        Dim X, Y, W, H, RX, BX As Integer
        X = 54
        Y = 32
        W = 64
        H = 32
        RX = 60
        BX = 30
        With MAP
            .Left = X + (SX \ RX)
            .Top = Y + (SY \ RX)
            Select Case ZM
                Case 0
                    .Width = W
                    .Height = H
                Case > 0
                    .Width = W * (1 + ZM / BX)
                    .Height = H * (1 + ZM / BX)
                Case < 0
                    .Width = W * (1 + ZM / BX)
                    .Height = H * (1 + ZM / BX)
            End Select
            Log_Show(2, .Left.ToString & "|" & .Top.ToString & "|" & .Width.ToString & "|" & .Height.ToString)

        End With
        'With lab.MapF
        '    .X = .X + X
        '    If Math.Abs(.X \ SX) > 0 Then
        '        MAP.Left = MAP.Left + (.X \ SX)
        '        .X = 0
        '    End If

        '    .Y = .Y + Y
        '    If Math.Abs(.Y \ SX) > 0 Then
        '        MAP.Top = MAP.Top + (.Y \ SX)
        '        .Y = 0
        '    End If


        '    Log_Show(2, MAP.Left.ToString & "|" & MAP.Top.ToString)

        'End With
        '  Log_Show(1, MAP.Left.ToString & "|" & MAP.Top.ToString)

    End Sub
    Private Sub Base_MouseMove(sender As Object, e As MouseEventArgs)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then
            Dim Wd, He As Integer
            'コントロール取得
            Dim control As Control = CType(sender, Control)

            Select Case gFlag_Pull
                Case "Move" '移動
                    'コントロールの位置を設定
                    control.Left = control.Left + e.X - gBase_StartX
                    control.Top = control.Top + e.Y - gBase_StartY
                    Me.Cursor = Cursors.Hand
                  '  Box_Re_Move(gLab, control.name, control.Left, control.Top)
                Case "Resize" '1 'リサイズ
                    Wd = gBase_W + e.X - gBase_StartX
                    He = gBase_H + e.Y - gBase_StartY
                    If Wd < 10 Then Wd = 10
                    If He < 10 Then He = 10
                    control.Width = Wd
                    control.Height = He
                    ' Me.Cursor = Cursors.SizeNESW
                    ' Box_Re_Size(gLab, control.name, control.Left, control.Top)
            End Select

            'リフレッシュ
            Me.Refresh()
        Else
            'マウスカーソルを変化させるだけ
            Dim fflag As Integer
            fflag = 0
            If (sender.Width - e.X) < 10 And (sender.Height - e.Y) < 10 Then
                fflag = 1
                Me.Cursor = Cursors.SizeNWSE
            End If

            If (e.X) < 10 And (sender.Height - e.Y) < 10 Then
                fflag = 1
                Me.Cursor = Cursors.SizeNESW
            End If

            If (sender.Width - e.X) < 10 And Math.Abs((sender.Height) / 2 - e.Y) < 10 Then
                fflag = 1
                Me.Cursor = Cursors.SizeWE
            End If

            If Math.Abs((sender.Width) / 2 - e.X) < 10 And (sender.Height - e.Y) < 10 Then
                fflag = 1
                Me.Cursor = Cursors.SizeNS
            End If


            If (e.X) < 10 And Math.Abs((sender.Height) / 2 - e.Y) < 10 Then
                fflag = 1

                Me.Cursor = Cursors.SizeWE
            End If


            If fflag = 0 Then

                Me.Cursor = Cursors.Hand
            End If



        End If
    End Sub

    Private Sub Line_MouseMove(sender As Object, e As MouseEventArgs)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then
            Dim RX, RY As Integer
            'コントロール取得
            Dim control As Control = CType(sender, Control)

            Select Case gFlag_Pull
                Case "Move"  '移動


            End Select
            'コントロールの位置を設定
            '   RX = e.X - gBase_StartX
            RY = e.Y - gBase_StartY
            control.Top = control.Top + RY

            Dim Tag As String = "OnLINE." & sender.name
            With TabControl3
                For Each Box As Panel In .TabPages(.SelectedIndex).Controls
                    If Box.Tag = Tag Then
                        '               Box.Left = Box.Left + RX
                        Box.Top = Box.Top + RY
                    End If
                Next
            End With
            'gBase_StartX = e.X
            'gBase_StartY = e.Y
            'リフレッシュ
            Me.Refresh()
        End If
    End Sub
    Private Sub Line_Re_Move(ByRef lab As Lab_Type, BoxID As Long, X As Integer, Y As Integer)
        'DB記録
        'Clothesline@LINE|0|DESK|X|Y|1|640|5|blue
        Dim BluePrint As String
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim Ds As DataSet
        Ds = DBAcs_Get_DataS(lab, "Box", "BluePrint", "ID = " & BoxID.ToString)

        If Ds.Tables(0).Rows.Count = 0 Then Return '異常

        STACKS = Split(Ds.Tables(0).Rows(0).Item(0), "@")
        STK_PRT = Split(STACKS(1), "|")

        STK_PRT(3) = 1
        STK_PRT(4) = Y.ToString


        STACKS(1) = String.Join("|", STK_PRT)
        BluePrint = String.Join("@", STACKS)
        DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)


    End Sub

    Private Sub Line_ReSize(ByRef lab As Lab_Type, BoxID As Long, W As Integer)
        'DB記録
        'Clothesline@LINE|0|DESK|X|Y|1|640|5|blue
        Dim BluePrint As String
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim Ds As DataSet
        Ds = DBAcs_Get_DataS(lab, "Box", "BluePrint", "ID = " & BoxID.ToString)

        If Ds.Tables(0).Rows.Count = 0 Then Return '異常

        STACKS = Split(Ds.Tables(0).Rows(0).Item(0), "@")
        STK_PRT = Split(STACKS(1), "|")


        STK_PRT(5) = W.ToString


        STACKS(1) = String.Join("|", STK_PRT)
        BluePrint = String.Join("@", STACKS)
        DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)


    End Sub


    Private Sub Box_Re_Move(ByRef lab As Lab_Type, BoxID As Long, ByRef box As Panel)
        'DB記録
        Dim BluePrint As String
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim Ds As DataSet
        Ds = DBAcs_Get_DataS(lab, "Box", "BluePrint", "ID = " & BoxID.ToString)

        If Ds.Tables(0).Rows.Count = 0 Then Return '異常

        STACKS = Split(Ds.Tables(0).Rows(0).Item(0), "@")
        STK_PRT = Split(STACKS(1), "|")

        STK_PRT(4) = box.Left.ToString
        STK_PRT(5) = box.Top.ToString
        STK_PRT(6) = ""
        STK_PRT(7) = box.Width.ToString
        STK_PRT(8) = box.Height.ToString

        STACKS(1) = String.Join("|", STK_PRT)
        BluePrint = String.Join("@", STACKS)
        DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)

        'パラメータ一覧
        '@BASE|0|DESK|Dock|X|Y|Z|W|H|"
        '@PICT|0|BASE|Dock|C|DAT|"
        '@TEXT|0|BASE|Dock|Mul|Bc|Fc|FName|FSize|Text|"
        '@WebView|0|BASE|Dock|URL|"
        '@TAB|0|BASE|Dock|Alignment|ItemSizeX|ItemSizeY|"
        '@TABp|0|TAB|Text|BackColor|"
        '@CogT|0|NEXT|Dock|TEXT|"
        '@Draw|0|NEXT|Dock|DAT|
        '@Split|0|NEXT|Dock|Orientation|SplitterDistance|


    End Sub
    Private Sub Box_Re_Size(ByRef lab As Lab_Type, BoxID As Long, W As Integer, H As Integer)
        'パラメータ一覧
        '@BASE|0|DESK|Dock|X|Y|Z|W|H|"
        'DB記録
        Dim BluePrint As String
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim Ds As DataSet
        Ds = DBAcs_Get_DataS(lab, "Box", "BluePrint", "ID = " & BoxID.ToString)

        If Ds.Tables(0).Rows.Count = 0 Then Return '異常

        STACKS = Split(Ds.Tables(0).Rows(0).Item(0), "@")
        STK_PRT = Split(STACKS(1), "|")

        STK_PRT(7) = W.ToString
        STK_PRT(8) = H.ToString

        STACKS(1) = String.Join("|", STK_PRT)
        BluePrint = String.Join("@", STACKS)
        DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)



    End Sub
    Private Function Box_Re_Text(ByRef lab As Lab_Type, BoxID As Long, StackID As Integer, Text As String) As String
        '設計図上のテキストをDBへのアクセスに書き換え
        '@TEXT|0|BASE|Dock|Mul|Bc|Fc|FName|FSize|Text|"
        'DB記録
        Dim BluePrint As String
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim Ds As DataSet
        Ds = DBAcs_Get_DataS(lab, "Box", "BluePrint", "ID = " & BoxID.ToString)

        If Ds.Tables(0).Rows.Count = 0 Then Return "" '異常

        STACKS = Split(Ds.Tables(0).Rows(0).Item(0), "@")

        STK_PRT = Split(STACKS(StackID), "|")

        STK_PRT(13) = Text

        STACKS(StackID) = String.Join("|", STK_PRT)
        BluePrint = String.Join("@", STACKS)
        DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)
        Return BluePrint
    End Function
    Private Sub Box_Re_Draw(ByRef lab As Lab_Type, BoxID As Long, StackID As Integer, DBID As String)
        '設計図上のデータをDBへのアクセスに書き換え
        '@Draw|0|NEXT|Dock|DaT| 
        'DB記録
        Dim BluePrint As String
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim Ds As DataSet
        Ds = DBAcs_Get_DataS(lab, "Box", "BluePrint", "ID = " & BoxID.ToString)

        If Ds.Tables(0).Rows.Count = 0 Then Return '異常

        STACKS = Split(Ds.Tables(0).Rows(0).Item(0), "@")

        STK_PRT = Split(STACKS(StackID), "|")

        STK_PRT(4) = DBID

        STACKS(StackID) = String.Join("|", STK_PRT)
        BluePrint = String.Join("@", STACKS)
        DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)

    End Sub
    Private Sub Box_Re_Pict(ByRef lab As Lab_Type, BoxID As Long, StackID As Integer, DBID As String)
        '設計図上のデータをDBへのアクセスに書き換え
        '@PICT|0|BASE|Dock|C|DAT|"
        'DB記録
        Dim BluePrint As String
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim Ds As DataSet
        Ds = DBAcs_Get_DataS(lab, "Box", "BluePrint", "ID = " & BoxID.ToString)

        If Ds.Tables(0).Rows.Count = 0 Then Return '異常

        STACKS = Split(Ds.Tables(0).Rows(0).Item(0), "@")

        STK_PRT = Split(STACKS(StackID), "|")

        STK_PRT(5) = DBID

        STACKS(StackID) = String.Join("|", STK_PRT)
        BluePrint = String.Join("@", STACKS)
        DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)

    End Sub
    Private Function Line_Check_OnLine(ByRef Lab As Lab_Type, top As Integer) As Long
        'LINE 上かのチェック
        '----------------------

        With TabControl3
            For Each Line As Panel In .TabPages(.SelectedIndex).Controls
                If Line.Tag = "LINE" Then
                    If Math.Abs(Line.Top - top) < 10 Then
                        Lab.Line = Line

                        Return Val(Line.Name)
                    End If
                End If
            Next
        End With

        Return 0
    End Function
    Private Sub Base_MouseUp(sender As Object, e As MouseEventArgs)
        '左クリックの場合

        Dim control As Control = CType(sender, Control)

        If e.Button = MouseButtons.Left Then

            'コントロール取得

            Select Case gFlag_Pull
                Case "Move" '移動
                    'コントロールの位置を設定
                    control.Left = control.Left + e.X - gBase_StartX
                    control.Top = control.Top + e.Y - gBase_StartY
                    Me.Cursor = Cursors.Hand
                Case "Resize" '1 'リサイズ
                    Box_Re_Size(gLab, Val(control.Name), control.Width, control.Height)
            End Select

            gFlag_Pull = "Move"
            Log_Show(1, "BID[" & control.Name & "]" & control.Left & "|" & control.Top & "|" & control.Width & "|" & control.Height)

        End If

        Dim LID As Long
        LID = Line_Check_OnLine(gLab, control.Top)  'LINE上か
        If LID <> 0 Then
            control.Tag = "OnLINE." & LID.ToString
            control.Top = gLab.Line.Top + 10
            DBAcs_Update(gLab, "Box", "State = " & LID.ToString, "ID = " & control.Name) 'ステート変更
        Else
            If InStr(control.Tag, "OnLINE") > 0 Then
                control.Tag = "OnDESK"
                DBAcs_Update(gLab, "Box", "State = 1", "ID = " & control.Name) 'ステート変更
            End If

        End If
        Box_Re_Move(gLab, Val(control.Name), control)

        'リフレッシュ
        Me.Refresh()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Line_MouseUp(sender As Object, e As MouseEventArgs)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then

            'コントロール取得
            Dim control As Control = CType(sender, Control)

            Select Case gFlag_Pull
                Case "Move" '移動

                    control.Top = control.Top + e.Y - gBase_StartY
                    Line_Re_Move(gLab, Val(control.Name), control.Left, control.Top)
                    Me.Cursor = Cursors.Hand
            End Select

            gFlag_Pull = "Move"
            Log_Show(1, "LINE[" & control.Name & "]" & control.Left & "|" & control.Top & "|" & control.Width & "|" & control.Height)


            'リフレッシュ
            Me.Refresh()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub BxPict_Drop(sender As Object, e As DragEventArgs)
        Dim files() As String = DirectCast(e.Data.GetData(DataFormats.FileDrop, False), String())
        '取得したファイルのパスを元にピクチャーボックスに画像を表示

        ''DBへ画像のパスを保存
        'sender.ImageLocation = files(0)
        'Bx_Re_Pict(gLab, sender.name, files(0))

        'DBへ画像のデータを保存
        Dim img As System.Drawing.Image = System.Drawing.Image.FromFile(files(0))
        sender.Image = img


        Bx_Re_Pict2(gLab, sender.name, files(0))

    End Sub

    Private Sub BxPict_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            ' ドラッグ中のファイルやディレクトリの取得
            Dim drags() As String =
          CType(e.Data.GetData(DataFormats.FileDrop), String())

            For Each d As String In drags
                If Not System.IO.File.Exists(d) Then
                    ' ファイル以外であればイベント・ハンドラを抜ける
                    Return
                End If
            Next

            e.Effect = DragDropEffects.Copy
        End If

    End Sub

    Private Sub WebView21_NavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs) Handles WebView21.NavigationCompleted
        Dim mx As String
        mx = e.HttpStatusCode

        GGGX()

    End Sub



    Private Async Sub GGGX()
        Dim html = Await WebView21.ExecuteScriptAsync("document.documentElement.outerHTML")
        html = Regex.Unescape(html)
        html = html.Remove(0, 1)
        html = html.Remove(html.Length - 1, 1)
        Debug.Print(html)

    End Sub



    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'gOwnerID = 1
        'gProjID = 1
        'gLab.OwnerID = 1
        'gLab.ProjID = 1
        'Dim KEY(3) As String

        ''KEY(0) = Desk_Make_New(gLab, "Summary", "要約", True)
        ''KEY(1) = Desk_Make_New(gLab, "Member", "メンバー", False)
        ''KEY(2) = Desk_Make_New(gLab, "Desk", "デスク", False)
        ''gLab.DeskName = KEY(1)
        ''Box_Makes_OnDESK(gLab, "Fig", 1, 1, 1)
        '''Box_Makes_OnDESK(gLab, "ID_Card", 1, 1, 1)
        ''gLab.DeskName = KEY(0)
        ''Box_Makes_OnDESK(gLab, "Summary", 1, 1, 1)

        ''With DataGridView1
        ''    .ColumnCount = 1
        ''    .RowCount = 1
        ''    .DefaultCellStyle.BackColor = Color.Yellow
        ''    .ColumnHeadersVisible = False
        ''    .RowHeadersVisible = False
        ''    .ScrollBars = ScrollBars.None
        ''    .DefaultCellStyle.WrapMode = DataGridViewTriState.True
        ''    .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        ''    .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        ''    .Font = New Font("Arial", 8)

        ''    .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
        ''    .Columns(0).DefaultCellStyle.ForeColor = Color.Blue
        ''    .Columns(0).Width = DataGridView1.Width - 1
        ''    .Rows(0).Height = DataGridView1.Height - 1

        ''    '.Rows(1).Cells(1).Style.BackColor = Color.Red


        ''End With
        ''With DataGridView2
        ''    .ColumnCount = 2
        ''    .RowCount = 5
        ''    .DefaultCellStyle.BackColor = Color.PaleTurquoise
        ''    .ColumnHeadersVisible = False
        ''    .RowHeadersVisible = False
        ''    .ScrollBars = ScrollBars.None
        ''    .DefaultCellStyle.WrapMode = DataGridViewTriState.True
        ''    .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        ''    .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        ''    .Font = New Font("Arial", 8)

        ''    .Columns(0).Width = 80
        ''    .Columns(1).Width = DataGridView2.Width - 80

        ''    .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
        ''    .Columns(0).DefaultCellStyle.ForeColor = Color.White
        ''    .Columns(0).DefaultCellStyle.BackColor = Color.DarkOliveGreen

        ''    .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
        ''    .Columns(1).DefaultCellStyle.ForeColor = Color.Black
        ''    .Columns(1).DefaultCellStyle.BackColor = Color.PaleTurquoise
        ''    .Rows(0).Cells(0).Value = "Title"
        ''    .Rows(1).Cells(0).Value = "Name"
        ''    .Rows(2).Cells(0).Value = "Tel"
        ''    .Rows(3).Cells(0).Value = "Mail"
        ''    .Rows(4).Cells(0).Value = "Address"
        ''End With

    End Sub
    Private Function Desk_Make_New(ByRef Lab As Lab_Type, Type As String, Title As String, NewFlag As Boolean) As String
        'DESK 新設 &DB登録
        '-------------------
        Dim DeskID As Long

        Dim key, tmpD As String
        Dim NewPage As Integer
        Dim Desk As TabPage
        Desk = New TabPage()
        With Lab.DeskAT

            With .View
                If NewFlag = True Then
                    .TabPages.Clear()
                End If
                NewPage = .TabCount + 1
            End With
            ' DBAcs_Init2()

            'DBへの登録
            key = Mid(Lab.DBACS_Field("Desk"), Len("ID,") + 1)

            tmpD = ""
            tmpD &= gX_ownerID.ToString & ","
            tmpD &= Lab.ProjID.ToString & ","
            tmpD &= NewPage.ToString & ","
            tmpD &= "'" & Type & "',"
            tmpD &= "'" & Title & "',"
            tmpD &= "1,"
            tmpD &= "0,"
            tmpD &= "0,"
            tmpD &= "0"

            DeskID = DBAcs_Insert(Lab, "Desk", key, tmpD)


            With Desk
                .Tag = Type & "_" & DeskID.ToString
                .Text = Title
                .Name = DeskID.ToString
                .AllowDrop = True
                .BackColor = Color.WhiteSmoke
            End With
            AddHandler Desk.DragOver, AddressOf Desk_DragOver
            AddHandler Desk.DragDrop, AddressOf Desk_DragDrop
            AddHandler Desk.MouseDown, AddressOf Desk_MouseDown
            AddHandler Desk.MouseDoubleClick, AddressOf Desk_DoubleClick
            AddHandler Desk.KeyDown, AddressOf Desk_KeyDown
            .View.TabPages.Add(Desk)

        End With
        Lab.DeskID = DeskID
        Lab.DeskName = Desk.Name
        Lab.TabID = gLab.TabID = NewPage - 1
        Return Desk.Name
    End Function

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub DB_CreateTBl(ByRef lab As Lab_Type)
        ' コネクション作成
        Dim con As SQLiteConnection
        con = New SQLiteConnection(GetConnectionString(lab.DbPath))
        con.Open()
        Using cmd = con.CreateCommand()
            ' テーブル作成SQL
            cmd.CommandText = "CREATE TABLE users (" &
                                  "id INTEGER PRIMARY KEY," &
                                  "name TEXT NOT NULL," &
                                  "age INTEGER," &
                                  "email TEXT NOT NULL UNIQUE" &
                                  ")"
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub DB_INSERT(ByRef lab As Lab_Type)
        ' コネクション作成
        Using con = New SQLiteConnection(GetConnectionString(lab.DbPath))
            con.Open()

            ' トランザクション開始
            Using tx = con.BeginTransaction()
                Using cmd = con.CreateCommand()
                    ' 挿入するデータの作成 
                    Dim users = {New With {.id = 1, .name = "一郎", .age = 40, .email = "abc@example.com"},
                                 New With {.id = 2, .name = "二郎", .age = 32, .email = "123@example.com"},
                                 New With {.id = 3, .name = "三郎", .age = 28, .email = "test@example.com"}
                                }

                    ' レコード挿入SQL
                    cmd.CommandText = "INSERT INTO users (id, name, age, email)" &
                                      " VALUES (@id, @name, @age, @email)"

                    ' パラメータ作成
                    cmd.Parameters.Add("id", DbType.Int32)
                    cmd.Parameters.Add("name", DbType.String)
                    cmd.Parameters.Add("age", DbType.Int32)
                    cmd.Parameters.Add("email", DbType.String)

                    For Each user In users
                        cmd.Parameters("id").Value = user.id
                        cmd.Parameters("name").Value = user.name
                        cmd.Parameters("age").Value = user.age
                        cmd.Parameters("email").Value = user.email
                        cmd.ExecuteNonQuery()
                    Next

                    ' コミット
                    tx.Commit()
                End Using
            End Using
        End Using
    End Sub

    'Private Sub DB_Select()
    '    ' コネクション作成
    '    Using con = New SQLiteConnection(GetConnectionString(lab.DBpath))
    '        con.Open()
    '        Using cmd = con.CreateCommand()
    '            ' テーブル検索SQL - 全レコードを取得する。（id列を昇順で並べて）
    '            cmd.CommandText = "SELECT * FROM users ORDER BY id"
    '            Using dr = cmd.ExecuteReader()
    '                While (dr.Read())
    '                    Console.WriteLine($"id:{dr("id")}, name:{dr("name")}, age:{dr("age")}, email:{dr("email")}")
    '                End While
    '            End Using
    '        End Using
    '    End Using
    'End Sub


    'Private Sub BB_UPdata()
    '    ' コネクション作成
    '    Using con = New SQLiteConnection(GetConnectionString(lab.DBpath))
    '        con.Open()
    '        Using cmd = con.CreateCommand()
    '            ' レコード更新SQL - ageが30以上のレコードに対して、nameの前に文字列 ex- を付加する。
    '            cmd.CommandText = "UPDATE users SET name = 'ex-' || name" &
    '                              " WHERE age >= 30"
    '            cmd.ExecuteNonQuery()
    '        End Using
    '    End Using
    'End Sub

    '''' <summary>
    '''' Deleteボタンクリックイベントハンドラ
    '''' </summary>
    'Private Sub DB_Delete()
    '    ' コネクション作成
    '    Using con = New SQLiteConnection(GetConnectionString(lab.DBpath))
    '        con.Open()
    '        Using cmd = con.CreateCommand()
    '            ' レコード削除SQL - idが 2 のレコードを削除する。
    '            cmd.CommandText = "DELETE FROM users WHERE id = 2"
    '            cmd.ExecuteNonQuery()
    '        End Using
    '    End Using
    'End Sub

    Private Function GetConnectionString(ByRef Dbpath) As String
        ' 
        Dim builder As SQLiteConnectionStringBuilder = New SQLiteConnectionStringBuilder()
        builder.DataSource = Dbpath

        Return builder.ConnectionString
    End Function
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        DB_CreateTBl(gLab)
        DB_INSERT(gLab)
    End Sub


    Private Sub ListView1_BeforeLabelEdit(sender As Object, e As LabelEditEventArgs) Handles ListView1.BeforeLabelEdit

    End Sub

    Private Sub ListView1_AfterLabelEdit(sender As Object, e As LabelEditEventArgs) Handles ListView1.AfterLabelEdit
        If e.Label = Nothing Then
            Return
        End If
        Proj_Rename(gLab, e.Label)

    End Sub

    Private Sub ListView3_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView3.MouseDoubleClick
        With ListView3
            If .SelectedItems.Count = 0 Then Return
            Element_Act_Item(gLab, .SelectedItems(0).Tag)
        End With
    End Sub

    Private Sub TabControl3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl3.SelectedIndexChanged

        If sender.SelectedTab Is Nothing Then Return
        With gLab
            .DeskID = Val(sender.SelectedTab.name)
            .TabID = sender.Selectedindex

            .DeskSx = .Desk(.TabID).ShiftX
            .DeskSy = .Desk(.TabID).ShiftY
            .DeskZm = .Desk(.TabID).ZM
        End With




        Log_Show(1, "DESK --" & gLab.DeskID.ToString)

    End Sub

    Private Sub TextBox1_MouseLeave(sender As Object, e As EventArgs) Handles TextBox1.MouseLeave

    End Sub

    Private Sub BxText_MouseLeave(sender As Object, e As EventArgs)
        Dim Tbox As TextBox
        Tbox = CType(sender, TextBox)
        Bx_Re_Text(gLab, Val(Tbox.Name), Tbox.Text)

    End Sub

    Private Sub BxDraw_MouseLeave(sender As Object, e As EventArgs)

        Dim DrID As Integer = Val(sender.Tag)

        Dim TEGAKI() As Byte
        TEGAKI = sender.Ink.Save()
        My.Computer.FileSystem.WriteAllBytes(gWORK_Path & "\DDX" & DrID.ToString, TEGAKI, True)
    End Sub

    Private Sub ContextMenuStrip2_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ContextMenuStrip2.ItemClicked
        Select Case e.ClickedItem.Tag
            Case "NEW"
                Desk_Make(gLab)
            Case "RENAME"
                MsgBox("ごめん、まだ作ってない") '  ListView1.SelectedItems(0).BeginEdit()
            Case "DEL"
                Desk_Del(gLab)

            Case Else
        End Select
    End Sub
    Private Sub Desk_Make(ByRef Lab As Lab_Type)
        'デスク追加
        '--------------------
        Desk_Make_New(gLab, ”Desk”, "Desk", False)

    End Sub

    Private Sub Desk_ReSize(sender As Object, e As EventArgs)
        'デスクリサイズ
        '--------------------
        Dim desk As TabPage
        desk = CType(sender, Control)
        For Each box As Panel In sender.Controls
            If box.Tag = "LINE" Then
                box.Width = desk.Width
                Line_ReSize(gLab, Val(box.Name), desk.Width)
            End If

        Next


    End Sub
    Private Sub ContextMenuStrip3_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ContextMenuStrip3.ItemClicked
        Select Case e.ClickedItem.Tag
            Case "COPY"
                MsgBox("いるかな？")
            Case "PASTE"
                MsgBox("いるかな？")
            Case "DEL"
                BX_Del(gLab)

            Case Else
        End Select
    End Sub
    Private Sub Proj_Del(ByRef Lab As Lab_Type)
        'Projの削除(DB修正)
        '-------------------------------------------
        Dim cmd As String
        If Lab.OwnerID <> gX_ownerID Then
            MsgBox（"オーナーでありません"）
            Return
        End If
        With Lab
            'ステート変更
            DBAcs_Update(Lab, "Proj", "State = 0", "ID = " & Lab.ProjID.ToString)

            cmd = " State > 0 "
            cmd &= "AND OwnerID = " & gX_ownerID
            Proj_Show_ListByCmd(gLab, cmd) 'オーナーのレポ表示

        End With
    End Sub
    Private Sub Desk_Del(ByRef Lab As Lab_Type)
        'DESKの削除
        '-------------------------------------------
        Dim cmd As String
        If Lab.OwnerID <> gX_ownerID Then
            MsgBox（"オーナーでありません"）
            Return
        End If
        If Lab.TabID = 0 Then
            MsgBox（"サマリーは削除できません"）
            Return
        End If

        With Lab
            'ステート変更
            DBAcs_Update(Lab, "Desk", "State = 0", "ID = " & Lab.DeskID.ToString)
            TabControl3.TabPages.RemoveAt(Lab.TabID)

        End With


    End Sub
    Private Sub Element_Del(ByRef Lab As Lab_Type)

    End Sub
    Private Sub BX_Del(ByRef Lab As Lab_Type)
        'BOX削除
        '------------------------

        '権限チェック

        'Dim cmd As String
        'If Lab.OwnerID <> gX_ownerID Then
        '    MsgBox（"オーナーでありません"）
        '    Return
        'End If

        gLab.Box.Visible = False

        'ステート変更
        DBAcs_Update(Lab, "Box", "State = 0", "ID = " & Lab.BoxID.ToString)

    End Sub
    Private Function Box_Re_ZM(ByVal WD As Integer, Sx As Integer, SZ As Integer) As Integer
        'BOXのリサイズ
        '----------------------

        If Sx > 0 Then
            WD = WD - (WD * SZ) / 100
            If WD < 10 Then WD = 10
        End If
        If Sx < 0 Then
            WD = WD + (WD * SZ) / 100
        End If
        Return WD
    End Function


    Private Sub Box_UpDown(ByRef Base As Panel, SZ As Integer, Sx As Integer)
        'BOXのZ軸移動
        '----------------------
        ' Dim Wd, He As Integer
        With Base

            If Sx > 0 Then
                .SendToBack()
            End If
            If Sx < 0 Then
                .BringToFront()
            End If

        End With

    End Sub
    Private Function Box_Move_ZM(Dp As Integer, Sp As Integer, Sx As Integer, Sz As Integer) As Integer
        'BOXのリサイズ
        '----------------------
        Select Case Dp - Sp
            Case 0
                Return Dp
            Case > 0
                Return Sp + Box_Re_ZM(Math.Abs(Dp - Sp), Sx, Sz)
            Case < 0
                Return Sp - Box_Re_ZM(Math.Abs(Dp - Sp), Sx, Sz)
        End Select

    End Function

    'Private Sub Box_MoveY(ByRef Base As Panel, Dx As Integer, Sx As Integer, ex As Integer)
    '    'BOXのリサイズ
    '    '----------------------
    '    Dim Y, wx As Integer
    '    With Base
    '        Y = .Top
    '        wx = Dx - Y
    '        If Math.Abs(wx) < 10 Then Return

    '        If ex > 0 Then
    '            Y = Y + (wx * Sx) / 100
    '        End If
    '        If ex < 0 Then
    '            Y = Y - (wx * Sx) / 100
    '        End If
    '        .Top = Y
    '    End With

    'End Sub

    Private Sub Form1_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel
        Dim Wd, He As Integer
        Select Case gLab.Flg_BoxMouse
            Case "OnBOX"
                '  Box_Re_ZM(gLab.Box, 10, e.Delta)
                Box_UpDown(gLab.Box, 10, e.Delta)

            Case "OnDESK"
                'If (Control.MouseButtons And MouseButtons.Right) = MouseButtons.Right Then
                '    With TabControl3
                '        For Each box As Panel In .TabPages(.SelectedIndex).Controls
                '            Box_Move_ZM(box, 10, e.Delta)
                '        Next
                '    End With
                '    Return
                'End If

                If (Control.MouseButtons And MouseButtons.Left) = MouseButtons.Left Then
                    '拡大縮小
                    Dim Dx, Dy, Mx, My, Sx As Integer

                    With TabControl3
                        If e.Delta > 0 Then gLab.DeskZm = gLab.DeskZm - 1
                        If e.Delta < 0 Then gLab.DeskZm = gLab.DeskZm + 1
                        Map_View(gLab, gLab.DeskSx, gLab.DeskSy, gLab.DeskZm)
                        Sx = 10
                        Log_Show(1, gLab.DeskZm)
                        'Dx = .TabPages(.SelectedIndex).Width / 2
                        'Dy = .TabPages(.SelectedIndex).Height / 2Box_Re_ZM
                        Dx = gLab.DeskSx
                        Dy = gLab.DeskSy
                        For Each box As Panel In .TabPages(.SelectedIndex).Controls
                            Select Case box.Tag
                                Case "LINE"

                                Case "OnDESK"
                                    box.Left = Box_Move_ZM(box.Left, Dx, e.Delta, Sx)
                                    box.Top = Box_Move_ZM(box.Top, Dy, e.Delta, Sx)

                                    box.Width = Box_Re_ZM(box.Width, e.Delta, Sx)
                                    box.Height = Box_Re_ZM(box.Height, e.Delta, Sx)
                                    Box_Re_Size(gLab, Val(box.Name), box.Width, box.Height)
                            End Select

                        Next
                    End With

                    Return
                End If
                    Case "OnLINE"
                If (Control.MouseButtons And MouseButtons.Left) = MouseButtons.Left Then
                    '拡大縮小？
                    Return

                End If
                'サクッとスライド
                Dim Ax As Integer
                If e.Delta = 0 Then Return
                If e.Delta < 0 Then Ax = -1
                If e.Delta > 0 Then Ax = 1
                Dim OnLine() As Panel
                Dim Max, Flag_Left, tmpWidth As Integer
                Max = 0
                Flag_Left = 0
                Dim Tag As String = "OnLINE." & gLab.LineID.ToString ' gLab.Line.Name
                Select Case Ax
                    Case -1
                        With TabControl3
                            tmpWidth = .TabPages(.SelectedIndex).Width
                            For Each Box As Panel In .TabPages(.SelectedIndex).Controls
                                If Box.Tag = Tag Then
                                    Max += 1
                                    ReDim Preserve OnLine(Max)
                                    OnLine(Max) = Box
                                    If OnLine(0) Is Nothing Then OnLine(0) = Box

                                    If OnLine(Max).Left > 0 Then
                                        Flag_Left = 1
                                        If OnLine(0).Left < 1 Then OnLine(0) = OnLine(Max)
                                        If OnLine(Max).Left < OnLine(0).Left Then OnLine(0) = OnLine(Max)

                                    End If

                                End If
                            Next
                        End With
                        If OnLine Is Nothing Then Return
                        If Flag_Left = 0 Then Return

                        Dim Mv As Integer
                        Mv = -OnLine(0).Left
                        For L = 1 To Max
                            OnLine(L).Left = OnLine(L).Left + Mv
                            Log_Show(1, OnLine(L).Left.ToString)
                        Next

                    Case 1
                        With TabControl3
                            tmpWidth = .TabPages(.SelectedIndex).Width
                            For Each Box As Panel In .TabPages(.SelectedIndex).Controls
                                If Box.Tag = Tag Then
                                    Max += 1
                                    ReDim Preserve OnLine(Max)
                                    OnLine(Max) = Box
                                    If OnLine(0) Is Nothing Then OnLine(0) = Box

                                    If OnLine(Max).Left < 0 Then
                                        Flag_Left = 1
                                        If OnLine(0).Left >= 0 Then OnLine(0) = OnLine(Max)
                                        If OnLine(Max).Left > OnLine(0).Left Then OnLine(0) = OnLine(Max)

                                    End If

                                End If
                            Next
                        End With
                        If OnLine Is Nothing Then Return

                        Dim Mv As Integer
                        Select Case Flag_Left
                            Case 1
                                Mv = OnLine(0).Left
                                For L = 1 To Max
                                    OnLine(L).Left = OnLine(L).Left - Mv
                                Next
                            Case 2
                            Case 0
                                Flag_Left = 2
                                For L = 1 To Max
                                    If OnLine(L).Left < tmpWidth Then
                                        Flag_Left = 0
                                        If OnLine(0).Left > tmpWidth Then OnLine(0) = OnLine(L)
                                        If OnLine(L).Left > OnLine(0).Left Then OnLine(0) = OnLine(L)
                                    End If
                                Next
                                Mv = tmpWidth - OnLine(0).Left
                                For L = 1 To Max
                                    OnLine(L).Left = OnLine(L).Left + Mv
                                    '  Log_Show(1, OnLine(L).Left.ToString)
                                Next

                            Case Else

                        End Select
                End Select

        End Select


    End Sub
    Private Sub BxPic_DoubleClick(sender As Object, e As EventArgs)
        If sender.tag = "" Then Return

        If InStr(sender.tag, "★"） > 0 Then Return
        System.Diagnostics.Process.Start(sender.tag)


    End Sub

    Private Sub Form1_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        gLab.Flg_BoxMouse = ""
    End Sub

    Private Sub Form1_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        gLab.Flg_BoxMouse = "OnDESK"
    End Sub

    Private Sub SplitContainer1_Panel2_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainer1.Panel2.Paint

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub ListView1_HandleCreated(sender As Object, e As EventArgs) Handles ListView1.HandleCreated

    End Sub

    Private Sub ListView1_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView1.ItemChecked

    End Sub

    Private Sub TabControl1_Resize(sender As Object, e As EventArgs) Handles TabControl1.Resize

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress

    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim cp As Point


        If e.KeyData = (Keys.Control Or Keys.V) Then


            With System.Windows.Forms.Cursor.Position
                cp = New Point(.X, .Y)
            End With
            cp = PointToClient(cp)
            cp.X = cp.X - TabControl3.Left
            cp.Y = cp.Y - TabControl3.Top
            Desk_Paste(cp)
        End If

    End Sub

    Private Sub TabPage11_Click(sender As Object, e As EventArgs) Handles TabPage11.Click

    End Sub

    Private Sub TabPage11_KeyDown(sender As Object, e As KeyEventArgs) Handles TabPage11.KeyDown

    End Sub
    Private Sub Desk_KeyDown(sender As Object, e As KeyEventArgs)

    End Sub

End Class