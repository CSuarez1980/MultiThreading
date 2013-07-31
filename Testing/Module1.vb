Module Module1
    Private WithEvents H As New MultiThreading.WorkerHandler

    Sub Main()

        H.Include_Worker(New Custom With {.Plant = "0045", .SAPBox = "L7P"})
        H.Run_Process()

    End Sub


    Friend Class Custom
        Inherits MultiThreading.Worker

        Private _SAPBox As String
        Private _Plant As String

        Public Property SAPBox() As String
            Get
                Return _SAPBox
            End Get
            Set(ByVal value As String)
                _SAPBox = value
            End Set
        End Property

        Public Property Plant() As String
            Get
                Return _Plant
            End Get
            Set(ByVal value As String)
                _Plant = value
            End Set
        End Property

        Public Overrides Sub MyWork()
            Dim T As Integer = 0
            Dim SC As New SAPCOM.SAPConnector
            Dim ErMsg As New List(Of String)

The_Process:
            Dim C As New Object
            C = SC.GetSAPConnection(_SAPBox, "BM4691", "LAT")
            Dim POs As New DataTable

            Dim GetConfirmation As Boolean

            GetConfirmation = False
            If C Is Nothing Then
                If T <= 3 Then
                    GoTo The_Process
                    T += 1
                Else
                    ErMsg.Add(Now.Date.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message: Couldn't get SAP connection.")
                    GoTo The_End
                End If
            End If

            Dim Rep As New SAPCOM.OpenOrders_Report(C)


            ' Rep.RepairsLevel = IncludeRepairs
            Rep.Include_GR_IR = True
            Rep.IncludeDelivDates = True
            Rep.Include_YO_Ref = True
            Rep.IncludePlant(_Plant)

            Rep.ExcludeMatGroup("S731516AW")
            Rep.ExcludeMatGroup("S801416AQ")
            Rep.ExcludeMatGroup("S731516AV")
            MsgBox("getting report")
            Rep.Execute()
            MsgBox("report finished")
            Dim EKES As New SAPCOM.EKES_Report(C)
            Dim EKKO As New SAPCOM.EKKO_Report(C)
            Dim NAST As New SAPCOM.NAST_Report(C)

            If Rep.Success Then
                If Rep.ErrMessage = Nothing Then
                    POs = Rep.Data
                    GetConfirmation = True
                    If POs.Columns.IndexOf("EKKO-WAERS-0219") <> -1 Then
                        POs.Columns.Remove("EKKO-WAERS-0219")
                    End If
                    If POs.Columns.IndexOf("EKPO-ZWERT") <> -1 Then
                        POs.Columns.Remove("EKPO-ZWERT")
                    End If
                    If POs.Columns.IndexOf("EKKO-WAERS-0218") <> -1 Then
                        POs.Columns.Remove("EKKO-WAERS-0218")
                    End If
                    If POs.Columns.IndexOf("EKKO-WAERS-0220") <> -1 Then
                        POs.Columns.Remove("EKKO-WAERS-0220")
                    End If
                    If POs.Columns.IndexOf("EKKO-MEMORYTYPE") <> -1 Then
                        POs.Columns.Remove("EKKO-MEMORYTYPE")
                    End If

                    Dim TN As New DataColumn
                    Dim SB As New DataColumn

                    TN.ColumnName = "Usuario"
                    TN.Caption = "Usuario"
                    TN.DefaultValue = "BM4691"

                    SB.DefaultValue = _SAPBox
                    SB.ColumnName = "SAPBox"
                    SB.Caption = "SAPBox"

                    POs.Columns.Add(TN)
                    POs.Columns.Add(SB)

                    Dim cRow As DataRow
                    For Each cRow In POs.Rows
                        EKKO.IncludeDocument(cRow.Item("Doc Number"))
                        EKES.IncludeDocument(cRow.Item("Doc Number"))
                        NAST.IncludeDocument(cRow.Item("Doc Number"))
                    Next
                Else
                End If

                If GetConfirmation Then

                    EKES.Execute()
                    If EKES.Success Then
                        Dim SBE As New DataColumn
                        SBE.DefaultValue = _SAPBox
                        SBE.ColumnName = "SAPBox"
                        SBE.Caption = "SAPBox"
                        EKES.Data.Columns.Add(SBE)

                        If EKES.Data.Columns.IndexOf("OA") <> -1 Then
                            EKES.Data.Columns.Remove("OA")
                        End If

                        If EKES.Data.Columns.IndexOf("O Reference") <> -1 Then
                            EKES.Data.Columns.Remove("O Reference")
                        End If

                    Else
                        ErMsg.Add(Now.Date.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message:" & EKES.ErrMessage)
                    End If

                    EKKO.Execute()
                    If EKKO.Success Then
                        Dim ESB As New DataColumn
                        ESB.DefaultValue = _SAPBox
                        ESB.ColumnName = "SAPBox"
                        ESB.Caption = "SAPBox"

                        If EKKO.Data.Columns.IndexOf("Company Code") <> -1 Then
                            EKKO.Data.Columns.Remove("Company Code")
                        End If
                        If EKKO.Data.Columns.IndexOf("Doc Type") <> -1 Then
                            EKKO.Data.Columns.Remove("Doc Type")
                        End If
                        If EKKO.Data.Columns.IndexOf("Created On") <> -1 Then
                            EKKO.Data.Columns.Remove("Created On")
                        End If
                        If EKKO.Data.Columns.IndexOf("Created By") <> -1 Then
                            EKKO.Data.Columns.Remove("Created By")
                        End If
                        If EKKO.Data.Columns.IndexOf("Vendor") <> -1 Then
                            EKKO.Data.Columns.Remove("Vendor")
                        End If
                        If EKKO.Data.Columns.IndexOf("Language") <> -1 Then
                            EKKO.Data.Columns.Remove("Language")
                        End If
                        If EKKO.Data.Columns.IndexOf("POrg") <> -1 Then
                            EKKO.Data.Columns.Remove("POrg")
                        End If
                        If EKKO.Data.Columns.IndexOf("PGrp") <> -1 Then
                            EKKO.Data.Columns.Remove("PGrp")
                        End If
                        If EKKO.Data.Columns.IndexOf("Currency") <> -1 Then
                            EKKO.Data.Columns.Remove("Currency")
                        End If
                        If EKKO.Data.Columns.IndexOf("Doc Date") <> -1 Then
                            EKKO.Data.Columns.Remove("Doc Date")
                        End If
                        If EKKO.Data.Columns.IndexOf("Validity Start") <> -1 Then
                            EKKO.Data.Columns.Remove("Validity Start")
                        End If
                        If EKKO.Data.Columns.IndexOf("Validity End") <> -1 Then
                            EKKO.Data.Columns.Remove("Validity End")
                        End If
                        If EKKO.Data.Columns.IndexOf("Y Refer") <> -1 Then
                            EKKO.Data.Columns.Remove("Y Refer")
                        End If
                        If EKKO.Data.Columns.IndexOf("SalesPerson") <> -1 Then
                            EKKO.Data.Columns.Remove("SalesPerson")
                        End If
                        If EKKO.Data.Columns.IndexOf("Telephone") <> -1 Then
                            EKKO.Data.Columns.Remove("Telephone")
                        End If

                        EKKO.Data.Columns.Add(ESB)
                        For Each r As DataRow In EKKO.Data.Rows
                            If r("O Reference").ToString.ToUpper.IndexOf("Y") <> -1 Then
                                r("O Reference") = "Manual"
                            Else
                                r("O Reference") = ""
                            End If
                        Next

                        NAST.Show_All_Records = True
                        NAST.AddCustomField("AENDE", "Chance")

                        If NAST.IsReady Then
                            NAST.Execute()
                            If NAST.Success Then
                                Dim NSB As New DataColumn
                                NSB.DefaultValue = _SAPBox
                                NSB.ColumnName = "SAPBox"
                                NSB.Caption = "SAPBox"
                                NAST.Data.Columns.Add(NSB)
                            Else
                                ErMsg.Add(Now.Date.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message:" & NAST.ErrMessage)
                            End If
                        End If

                    Else
                        ErMsg.Add(Now.Date.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message:" & EKKO.ErrMessage)
                    End If
                End If
            Else
                ErMsg.Add(Now.Date.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message:" & Rep.ErrMessage)
            End If


The_End:
            ErMsg.Add(Now.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message: Finished.")

        End Sub
    End Class
End Module
