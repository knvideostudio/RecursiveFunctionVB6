Attribute VB_Name = "iRecursiveFunc"
' *********************************************************************************************************************
' * Author:          Krassimir Nikov
' * Email:           nikov@rokamboll.com
' * Release date:    Apr 03, 2006
' * History:         Apr 03, 2006 - Initial Code
' * Web Site:        www.rokamboll.com/recursivefunc
' * Resume web site: www.rokamboll.com/my_profile.htm
' *********************************************************************************************************************

Option Explicit

Private sQuery As String
Private ObjCnn As ADODB.Connection
Private ObjCnn2 As ADODB.Connection
Private ObjRs As ADODB.Recordset
Private ObjRs3 As ADODB.Recordset
Private x As MSXML2.DOMDocument40
Private root As IXMLDOMElement, rootAdd As IXMLDOMElement
Private rootAtrib As IXMLDOMAttribute
Private node As IXMLDOMNodeList

Private Const STR_CONNECT = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=dbRecursiveCategory;Data Source=(local);Password=12345;"
Global Const FILE_NAME = "RecursiveFunc.xml"


Public Function BuildCategoriesRoot() As String()
 Dim Arr(14) As String
 Dim i As Integer, Affected As Integer
 Dim sTmp As String
 i = 0
 
 Set ObjCnn = New ADODB.Connection
 Set ObjRs = New ADODB.Recordset
 ObjCnn.ConnectionString = STR_CONNECT
 ObjCnn.open

    sQuery = "select tbCategoryMain.UniqueValue as CattRootValue, tbCategoryText.sText as CattRootText, " & _
                " dbo.CountChildren(tbCategoryMain.UniqueValue) as CategoryCount " & _
                " from dbo.tbCategoryMain with (nolock), " & _
                " dbo.tbCategoryText with (nolock) " & _
                " Where tbCategoryMain.UniqueValue = tbCategoryText.CategoryMainUniqueValue " & _
                " and tbCategoryMain.UniqueValue not in " & _
                " (select ChildUniqueValue from dbo.tbCategoryRelation with (nolock)) and " & _
                " tbCategoryMain.UniqueValue in " & _
                " (select ParentUniqueValue from dbo.tbCategoryRelation with (nolock)) " & _
                " order by tbCategoryText.sText "
    
    Set ObjRs = ObjCnn.Execute(sQuery, Affected)
        If ObjRs.EOF = False Then
            ObjRs.MoveFirst
            Do While Not ObjRs.EOF
                'ReDim Preserve strParent(lngTopIndex)
                Arr(i) = ObjRs.fields("CattRootValue").Value
                sTmp = ObjRs.fields("CattRootText").Value
                CreateFirstNodeXMLtext sTmp, Arr(i), "", True

                i = i + 1
            ObjRs.MoveNext
            Loop
        Else
            Arr(0) = "0"
        End If
       
 ObjCnn.Close
BuildCategoriesRoot = Arr
End Function

Private Function SearchForNodeXMLtext(sSearchNode As String, _
        sNewCategoryId As String, _
        sNewCategoryText As String, _
        sXML As String, bool As Boolean)

Dim xSelectedNode As IXMLDOMElement
Dim xQ, s2 As String
Dim sMy As String
Dim strFilePath$

strFilePath = App.Path & "\" & FILE_NAME

    Set x = CreateObject("MSXML2.DOMDocument")
        x.async = False
        x.preserveWhiteSpace = False
        x.setProperty "SelectionLanguage", "XPath"
        
        If bool = True Then
            x.Load strFilePath
        Else
            x.loadXML sXML
        End If
    

        xQ = "//TreeNode[@Id='" & sSearchNode & "']"
        Set xSelectedNode = x.documentElement.selectSingleNode(xQ)
        
        If xSelectedNode Is Nothing Then
            'xSelected.setAttribute "Id", "567"
            ' Create an element in root
            'Set root = x.documentElement
            sMy = "nothing"
        Else
            Set rootAdd = x.createElement("TreeNode")
            
            xSelectedNode.appendChild rootAdd
            
            ' Create an Attribute with Value
            Set rootAtrib = x.createAttribute("Text")
            rootAtrib.Text = sNewCategoryText
            rootAdd.setAttributeNode rootAtrib
            Set rootAtrib = Nothing
    
            Set rootAtrib = x.createAttribute("Id")
            rootAtrib.Text = sNewCategoryId
            rootAdd.setAttributeNode rootAtrib
            Set rootAtrib = Nothing
            
            'x.Save strPath
            If bool = True Then
                x.save strFilePath
                s2 = ""
            Else
                s2 = x.xml
            End If
            
            Set rootAdd = Nothing
        End If
    
    Set x = Nothing
    SearchForNodeXMLtext = s2
End Function

' *********************************************************************************************************************
' *
' *
' *
' *********************************************************************************************************************
Public Function BuildCatheoriesChildren(sRootPKey() As String) As String()
 Dim Affected As Integer
 Dim strTMP As String, CategoryText$
 Dim strNextLevel As String
 Dim i As Integer, k As Integer, b As Integer
 'Dim nDepth As Integer
 Dim ArrCatt() As String
 
 Set ObjCnn = New ADODB.Connection
 Set ObjRs = New ADODB.Recordset
 ObjCnn.ConnectionString = STR_CONNECT
 ObjCnn.open
 
 i = 0
 

 For b = 0 To UBound(sRootPKey)
  sQuery = "Select ParentUniqueValue, ChildUniqueValue, " & _
            " dbo.CountChildren(ParentUniqueValue) As MyCount " & _
            " From dbo.tbCategoryRelation " & _
            " where ParentUniqueValue = '" & sRootPKey(b) & "'"
            
      Set ObjRs = ObjCnn.Execute(sQuery)
        
        If ObjRs.EOF = False Then
            ObjRs.MoveFirst
            i = ObjRs.fields("MyCount").Value
            
            ReDim ArrCatt(i - 1)
            k = 0
            
            Do Until ObjRs.EOF
                strTMP = ObjRs.fields("ChildUniqueValue").Value
                ArrCatt(k) = strTMP
                CategoryText = GetCategoryText(strTMP)
                SearchForNodeXMLtext sRootPKey(b), strTMP, CategoryText, "", True
                k = k + 1
            ObjRs.MoveNext
            Loop
        
            ' recall the same function again
            BuildCatheoriesChildren = BuildCatheoriesChildren(ArrCatt)
        End If
 Next b
 'ObjCnn.Close
End Function

Private Sub LogIT(txtText As String)
Dim sDate As String
sDate = CStr(Month(Now())) & "-" & CStr(Day(Now())) & "-" & CStr(Year(Now()))

     On Error Resume Next

     Open App.Path & "\" & sDate & "_Category_Log.txt" For Append Access Write As #44
     Print #44, txtText
     Close #44

End Sub

Private Function GetCategoryText(sValue As String) As String
Dim sTmp As String
Dim sql As String
sTmp = ""
 Set ObjCnn2 = New ADODB.Connection
 Set ObjRs3 = New ADODB.Recordset
 ObjCnn2.ConnectionString = STR_CONNECT
 ObjCnn2.open
    sql = "select [sText] from tbCategoryText " & _
        " where CategoryMainUniqueValue = '" & sValue & "'"
     Set ObjRs3 = ObjCnn2.Execute(sql)
       If ObjRs3.EOF = False Then sTmp = ObjRs3.fields("sText").Value
 ObjCnn2.Close
 GetCategoryText = sTmp
End Function


' *****************************************************************************
'  Generate first node
' *****************************************************************************
Private Function CreateFirstNodeXMLtext(sParentText As String, _
    sParentId As String, _
    strXML As String, _
    bool As Boolean) As String

Dim strFilePath As String
Dim s As String
    
strFilePath = App.Path & "\" & FILE_NAME

    Set x = CreateObject("MSXML2.DOMDocument")
        x.async = False
        x.preserveWhiteSpace = False
        x.setProperty "SelectionLanguage", "XPath"

        If bool = True Then
            x.Load strFilePath
        Else
            x.loadXML strXML
        End If
          
        ' Create an element in root
        Set root = x.documentElement
        Set rootAdd = x.createElement("TreeNode")
        root.appendChild rootAdd
        
        ' Create an Attribute with Value
        Set rootAtrib = x.createAttribute("Text")
        rootAtrib.Text = sParentText
        rootAdd.setAttributeNode rootAtrib
        Set rootAtrib = Nothing
        
        'Create next attribute
        Set rootAtrib = x.createAttribute("Id")
        rootAtrib.Text = sParentId
        rootAdd.setAttributeNode rootAtrib
        Set rootAtrib = Nothing
        
        If bool = True Then
            x.save strFilePath
            s = ""
        Else
            s = x.xml
        End If
        
        
    Set rootAdd = Nothing
    Set root = Nothing
    Set x = Nothing
CreateFirstNodeXMLtext = s
End Function
