
Imports Contensive.BaseClasses

Namespace Contensive.Addon.PageNavigation2
    Module commonModule
        Private WorkingQueryString As String
        '
        Public Const PageTypeRootChild As Integer = 1
        Public Const PageTypeSibling As Integer = 2
        Public Const PageTypeChild As Integer = 3
        Public Const PageTypeSiblingChild As Integer = 4
        '
        Const ContentNamePageContent = "Page Content"
        '
        Private LoadedPageID As Long
        '
        '=====================================================================================
        ' copy/paste from com version
        '=====================================================================================
        '
        Friend Function GetPageNavigation(cp As CPBaseClass, GivenPageType As Integer, TierMax As Integer) As String
            Try
                '
                Dim Caption As String
                Dim ContentPageStructure As String
                '
                Dim ContentPageStructureArray() As String
                Dim RowPointer As Integer
                Dim RowCount As Integer
                '
                Dim CSPointer As Integer
                '
                Dim ColumnDelimiter As String
                Dim ColumnArray() As String
                Dim ColumnCount As Integer
                Dim ColumnPointer As Integer
                '
                Dim CriteriaString As String
                '
                Dim LinkCaption As String
                Dim Link As String
                Dim CurrentPageID As Integer
                Dim RootPageID As Integer
                Dim ParentPageID As Integer
                Dim ThisPageType As String
                Dim InnerString As String
                Dim RootToCurrentList As String = ""
                Dim SQLNow As String
                '
                Dim SortCriteria As String
                '
                Dim CurrentRecordID As Integer
                '
                Dim BakeName As String
                Dim BakeConent As String
                '
                Dim SecondTest As String
                '
                Dim wrapperClass As String = cp.Doc.GetText("wrapper class")
                Dim listClass As String = cp.Doc.GetText("list class")
                Dim itemClass As String = cp.Doc.GetText("item class")
                Dim firstClass As String = cp.Doc.GetText("first class")
                Dim lastClass As String = cp.Doc.GetText("last class")
                Dim activeClass As String = cp.Doc.GetText("active class")
                Dim itemPtr As Integer = 0
                Dim classAttribute As String
                Dim ul As String = ""
                Dim li As String = ""
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim itemlink As String
                Dim childPageList As String
                '
                SQLNow = cp.Db.EncodeSQLDate(Now())
                ContentPageStructure = cp.Doc.NavigationStructure
                ContentPageStructureArray = Split(ContentPageStructure, vbCrLf)
                RowCount = UBound(ContentPageStructureArray) + 1
                For RowPointer = 0 To RowCount - 1
                    ColumnDelimiter = Left(ContentPageStructureArray(RowPointer), 1)
                    ColumnArray = Split(ContentPageStructureArray(RowPointer), ColumnDelimiter)
                    ColumnCount = UBound(ColumnArray) + 1
                    If ColumnCount > 1 Then
                        'Call Main.TestPoint("ColumnCount: " & ColumnCount)
                        'For ColumnPointer = 0 To ColumnCount - 1
                        'Call Main.TestPoint("ColumnArray(" & ColumnPointer & "): " & ColumnArray(ColumnPointer))
                        If ColumnArray(1) = "0" Then
                            RootPageID = cp.Utils.EncodeInteger(ColumnArray(2))
                        End If
                        If ColumnArray(1) = "2" Then
                            ParentPageID = cp.Utils.EncodeInteger(ColumnArray(3))
                            CurrentPageID = cp.Utils.EncodeInteger(ColumnArray(2))
                        End If
                        If ColumnArray(1) < "3" Then
                            RootToCurrentList = RootToCurrentList & "," & ColumnArray(2)
                        End If
                        'Next
                    End If
                Next
                If RootToCurrentList <> "" Then
                    RootToCurrentList = Mid(RootToCurrentList, 2)
                End If
                '
                If RootPageID = 0 Then
                    RootPageID = CurrentPageID
                End If
                If ParentPageID = 0 Then
                    ParentPageID = cp.Content.GetRecordID(ContentNamePageContent, "Landing Page Content")
                End If
                '
                BakeName = "PageNavigation_Type" & GivenPageType & "_Record" & CurrentPageID
                BakeConent = cp.Cache.Read(BakeName)
                '
                If BakeConent = "" Then
                    CriteriaString = ""
                    SortCriteria = ""
                    Select Case GivenPageType
                        Case PageTypeRootChild
                            SortCriteria = GetChildPageListSortMethod(cp, ContentNamePageContent, RootPageID)
                            CriteriaString = "(ParentID=" & cp.Db.EncodeSQLNumber(RootPageID) & ")"
                        Case PageTypeChild
                            SortCriteria = GetChildPageListSortMethod(cp, ContentNamePageContent, CurrentPageID)
                            CriteriaString = "(ParentID=" & cp.Db.EncodeSQLNumber(CurrentPageID) & ")"
                        Case PageTypeSibling
                            SortCriteria = GetChildPageListSortMethod(cp, ContentNamePageContent, ParentPageID)
                            CriteriaString = "(ParentID=" & cp.Db.EncodeSQLNumber(ParentPageID) & ")"
                        Case PageTypeSiblingChild
                            SortCriteria = GetChildPageListSortMethod(cp, ContentNamePageContent, ParentPageID)
                            CriteriaString = "(ParentID=" & cp.Db.EncodeSQLNumber(RootPageID) & ")"
                    End Select
                    '
                    SecondTest = "" _
                        & "(AllowInMenus<>0)" _
                        & "And((PubDate is null)or(PubDate<" & SQLNow & "))" _
                        & "And((DateArchive is null)or(DateArchive>" & SQLNow & "))" _
                        & "And((dateexpires is null)or(dateexpires>" & SQLNow & "))" _
                        & ""
                    If CriteriaString <> "" Then
                        If cs.Open(ContentNamePageContent, CriteriaString & " AND (" & SecondTest & " )", SortCriteria, True, "ID, Name,MenuHeadline") Then
                            Do
                                childPageList = ""
                                CurrentRecordID = cs.GetInteger("ID")
                                Caption = cs.GetText("menuheadline")
                                itemlink = cp.Content.GetPageLink(CurrentRecordID)
                                If Caption = "" Then
                                    Caption = cs.GetText("name")
                                    If Caption = "" Then
                                        Caption = "Page " & CurrentRecordID
                                    End If
                                End If
                                classAttribute = itemClass
                                If itemPtr = 0 Then
                                    classAttribute &= " " & firstClass
                                End If
                                If CurrentPageID = CurrentRecordID Then
                                    classAttribute &= " " & activeClass
                                End If
                                If (GivenPageType = PageTypeSiblingChild) And IsInDelimitedString(RootToCurrentList, CStr(CurrentRecordID), ",") Then
                                    childPageList = GetChildPageItems(cp, CurrentRecordID, RootToCurrentList, 1, TierMax, itemClass, listClass)
                                    If childPageList <> "" Then
                                        classAttribute = ""
                                        If listClass <> "" Then
                                            classAttribute = " class=""" & listClass & """"
                                        End If
                                        childPageList = vbCrLf & vbTab & "<ul" & classAttribute & ">" & childPageList.Replace(vbCrLf & vbTab, vbCrLf & vbTab & vbTab) & vbCrLf & vbTab & "</ul>"
                                    End If
                                End If
                                Call cs.GoNext()
                                If Not cs.OK Then
                                    classAttribute &= " " & lastClass
                                End If
                                If classAttribute <> "" Then
                                    classAttribute = " class=""" & classAttribute & """"
                                End If
                                ul &= vbCrLf & vbTab & "<li" & classAttribute & "><a href=""" & itemlink & """>" & Caption & "</a> " & childPageList & "</li>"
                                itemPtr += 1
                            Loop While cs.OK
                            classAttribute = ""
                            If listClass <> "" Then
                                classAttribute = " class=""" & listClass & """"
                            End If
                            ul = vbCrLf & vbTab & "<ul" & classAttribute & ">" & ul.Replace(vbCrLf & vbTab, vbCrLf & vbTab & vbTab) & vbCrLf & vbTab & "</ul>"
                        End If
                        Call cs.Close()
                    End If
                    classAttribute = ""
                    If wrapperClass <> "" Then
                        classAttribute = " class=""" & wrapperClass & """"
                    End If
                    ul = vbCrLf & vbTab & "<div" & classAttribute & ">" & ul.Replace(vbCrLf & vbTab, vbCrLf & vbTab & vbTab) & vbCrLf & vbTab & "</div>"
                    If ul <> "" Then
                        Call cp.Cache.Save(BakeName, ul, ContentNamePageContent)
                    End If
                Else
                    ul = BakeConent
                End If
                '
                GetPageNavigation = ul
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex, "exception in " & System.Reflection.MethodBase.GetCurrentMethod.ToString)
            End Try
        End Function
        '
        '   Returns the child sort method for the given page
        '
        Private Function GetChildPageListSortMethod(cp As CPBaseClass, ContentName As String, RecordID As Integer) As String
            GetChildPageListSortMethod = "Name"
            Try
                '
                Dim ChildListSortMethodID As Integer
                Dim cs As CPCSBaseClass = cp.CSNew()
                '
                If cs.Open(ContentName, "ID=" & cp.Db.EncodeSQLNumber(RecordID), "", True, "ChildListSortMethodID") Then
                    ChildListSortMethodID = cs.GetInteger("ChildListSortMethodID")
                End If
                Call cs.Close()
                '
                If cs.Open("Sort Methods", "ID=" & cp.Db.EncodeSQLNumber(ChildListSortMethodID), "", True, "OrderByClause") Then
                    GetChildPageListSortMethod = cs.GetText("OrderByClause")
                End If
                Call cs.Close()
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex, "exception in " & System.Reflection.MethodBase.GetCurrentMethod.ToString)
            End Try
        End Function
        '
        Private Function GetChildPageItems(cp As CPBaseClass, ParentPageID As Integer, RootToCurrentList As String, Tier As Integer, MaxTier As Integer, itemClass As String, listClass As String) As String
            Dim childPageList As String = ""
            Try
                '
                Dim CurrentRecordID As Integer
                Dim MenuHeadline As String
                Dim SubStyleName As String
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim classAttribute As String = ""
                Dim childPageItems As String

                '
                If Tier <= MaxTier Then
                    If cs.Open(ContentNamePageContent, "ParentID=" & cp.Db.EncodeSQLNumber(ParentPageID), GetChildPageListSortMethod(cp, ContentNamePageContent, ParentPageID)) Then
                        Do
                            CurrentRecordID = cs.GetInteger("ID")
                            MenuHeadline = cs.GetText("MenuHeadline")
                            SubStyleName = "Tier" & CStr(Tier)
                            If MenuHeadline <> "" Then
                                '
                                ' display this page navigation
                                '
                                childPageList &= vbCrLf & vbTab & "<li class=""subNav subNav" & SubStyleName & """><a href=""" & cp.Content.GetPageLink(CurrentRecordID) & """>" & MenuHeadline & "</a>"
                                '
                                ' if this page is in the Root-To-Current list, get its child pages also
                                '
                                If IsInDelimitedString(RootToCurrentList, CStr(CurrentRecordID), ",") Then
                                    childPageItems = GetChildPageItems(cp, CurrentRecordID, RootToCurrentList, Tier + 1, MaxTier, itemClass, listClass)
                                    If childPageItems <> "" Then
                                        classAttribute = ""
                                        If listClass <> "" Then
                                            classAttribute = " class=""" & listClass & """"
                                        End If
                                        childPageList = vbCrLf & vbTab & "<ul" & classAttribute & ">" & childPageItems.Replace(vbCrLf & vbTab, vbCrLf & vbTab & vbTab) & vbCrLf & vbTab & "</ul>"
                                    End If
                                End If
                                childPageList &= vbCrLf & vbTab & "</li>"
                            End If
                            Call cs.GoNext()
                        Loop While cs.OK
                    End If
                    Call cs.Close()
                    '
                End If
                '
                'GetChildPageList = childPageList
            Catch ex As Exception
                Call cp.Site.ErrorReport(ex, "exception in " & System.Reflection.MethodBase.GetCurrentMethod.ToString)
            End Try
            Return childPageList
        End Function
        '
        '==========================================================================================
        '   Test if a test string is in a delimited string
        '==========================================================================================
        '
        Public Function IsInDelimitedString(DelimitedString As String, TestString As String, Delimiter As String) As Boolean
            IsInDelimitedString = (0 <> InStr(1, Delimiter & DelimitedString & Delimiter, Delimiter & TestString & Delimiter, vbTextCompare))
        End Function


    End Module

End Namespace