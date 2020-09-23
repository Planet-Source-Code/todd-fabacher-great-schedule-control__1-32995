Attribute VB_Name = "xmlXML"
Option Explicit

Public Sub AppendChild(ParentNode As IXMLDOMNode, ChildNode As IXMLDOMNode, Optional bCloneFirstIfNeeded As Boolean = True)

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: To simplifly the xml AppendChild finction.
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author    Change
'03/01/2002 Todd      Initial Design
'=====================================================================

Dim Node As IXMLDOMNode
    
    'if the child already has a parent and we don't want to remove from the parent
    If Not ChildNode.ParentNode Is Nothing And bCloneFirstIfNeeded Then
        'clone the node
        Set Node = ChildNode.cloneNode(True)
    'either the child has no parent or we don't care if we remove it from its parent
    Else
        'just use the node
        Set Node = ChildNode
    End If
    
    'append the node
    Call ParentNode.AppendChild(Node)
    
End Sub
Public Sub ClearXML(xmlNode As IXMLDOMElement)

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Clears an XML Node
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author    Change
'03/01/2002 Todd      Initial Design
'=====================================================================

    'Load the SGroups if there are any
    If xmlNode Is Nothing Then Exit Sub
    Do Until xmlNode.childNodes.length = 0
        Call xmlNode.removeChild(xmlNode.childNodes.Item(0))
    Loop
    
End Sub
Public Function CreateElement(ParentElement As IXMLDOMElement, sElementName$, Optional bAddToParent As Boolean = True) As IXMLDOMElement
    
'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: To simplifly the xml CreateElement finction.
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author    Change
'03/01/2002 Todd      Initial Design
'=====================================================================

Dim Document As FreeThreadedDOMDocument

    'if no parent element
    If ParentElement Is Nothing Then
        'create a new document
        Set Document = New FreeThreadedDOMDocument
        'can't add to the parent
        bAddToParent = False
    'if the node has no owner document
    ElseIf ParentElement.ownerDocument Is Nothing Then
        'it is the document
        Set Document = ParentElement
    'has an owner document
    Else
        'use it
        Set Document = ParentElement.ownerDocument
    End If
    'create (and return) the element
    Set CreateElement = Document.CreateElement(sElementName)
    
    'if we should add to the parent
    If bAddToParent Then
        'append to the parent
        Call ParentElement.AppendChild(CreateElement)
    End If
    
End Function
Public Function LoadXMLFile(FileName As String) As IXMLDOMElement

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Wrapper for the DOM.Load function
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author    Change
'03/01/2002 Todd      Initial Design
'=====================================================================

Dim Document As New FreeThreadedDOMDocument
Dim xmlError As IXMLDOMElement

    'use the new document
    With Document
        'load synchronously
        .async = False
            
        'load the document and if it fails
        If Not .Load(FileName) Then
            Set xmlError = CreateElement(Nothing, "error")
            xmlError.setAttribute "msg", .parseError.reason
            xmlError.setAttribute "id", .parseError.errorCode
            Set LoadXMLFile = xmlError
            Exit Function
        End If
    End With
    
    'return the root element
    Set LoadXMLFile = Document.documentElement
    
End Function
Public Function xmlAttr(xmlNode As IXMLDOMElement, sAttr As String, Optional Default As String = "") As String

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: To simplifly the "getAttribute" function.
'      Notes: Takes into consideration that there maybe an error.
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author    Change
'03/01/2002 Todd      Initial Design
'=====================================================================

On Error Resume Next

Dim sAttrVal As String

    sAttrVal = "---"
    sAttrVal = xmlNode.getAttribute(sAttr)
    If sAttrVal = "---" Then
        sAttrVal = Default
    End If
    xmlAttr = sAttrVal
    
End Function
Function LoadXML(sXML$) As IXMLDOMElement

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Wrapper for the DOM.LoadXML function
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author    Change
'03/01/2002 Todd      Initial Design
'=====================================================================

Dim Document As New FreeThreadedDOMDocument
Dim xmlError As IXMLDOMElement

    
    'use the new document
    With Document
        'load synchronously
        .async = False
            
        'load the document and if it fails
        If Not .LoadXML(sXML) Then
            Set xmlError = CreateElement(Nothing, "error")
            xmlError.setAttribute "msg", .parseError.reason
            xmlError.setAttribute "id", .parseError.errorCode
            Set LoadXML = xmlError
            Exit Function
        End If
    End With
    
    'return the root element
    Set LoadXML = Document.documentElement
    
End Function

Public Function SetErrorMsg(ByVal ErrNum As Long, ByVal ErrMsg As String) As IXMLDOMElement

    Set SetErrorMsg = CreateElement(Nothing, "error")
    SetErrorMsg.setAttribute "id", Format(ErrNum)
    SetErrorMsg.setAttribute "msg", ErrMsg
        
End Function
