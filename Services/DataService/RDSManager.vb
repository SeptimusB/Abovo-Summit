Namespace Abovo
    Public Class RDSManager

        Private RDSs() As RDSInstance
        Private LiveRDSCount As Integer = 0
        Private LiveRDSIndex As Integer = -1
        Private EventIDIndex As Integer = -1
        Private ModelID As Integer

        Public Sub New(SetModelID As Integer)

            ModelID = SetModelID

        End Sub

        Private Function GetNewEventID() As Integer

            EventIDIndex += 1
            Return EventIDIndex

        End Function


        Public Function ManageNRManipulationStart(NR As String) As Integer

            If LiveRDSCount = 0 Then Return -1

            Dim ReturnValue As Boolean = -1
            Dim EventSet As Boolean = False
            Dim CurrentEventID As Integer = GetNewEventID()

            For Each RDS As RDSInstance In RDSs

                If RDS.LiveNR = NR Then

                    If RDS.CurrRDSStatus = 1 Then

                        RDS.DisconnectRDS(CurrentEventID)
                        ReturnValue = 1

                    End If

                End If

            Next RDS

            Return ReturnValue

        End Function

        Public Function ManageNRManipulationEnd(NR As String) As Integer

            If LiveRDSCount = 0 Then Return -1

            Dim ReturnValue As Boolean = -1
            Dim EventSet As Boolean = False
            Dim CurrentEventID As Integer = -1

            For Each RDS As RDSInstance In RDSs

                If RDS.LiveNR = NR Then

                    If EventSet = False Then
                        EventSet = True

                        RDS.ReConnectRDS(CurrentEventID)

                    End If

                End If

            Next RDS

            Return ReturnValue

        End Function






    End Class
    <Serializable>
    Class RDSInstance

        Public EventID As Integer
        Public RDSID As Integer
        Public LiveNR As String
        Private RDSStatus As Integer '0 = Paused/Unlocked, 1 = Live/In Use, 2 = Locked for Edit
        Public AttachedGridControl As DevExpress.XtraGrid.GridControl
        Public AttachedGridView As DevExpress.XtraGrid.Views.Grid.GridView
        Public AttachedGridColView As DevExpress.XtraGrid.Views.Base.ColumnView
        Public MasterRDS As AbovoRangeDataSource
        Public RDSCopy As AbovoRangeDataSource
        Public TempRDS As AbovoRangeDataSource
        Public GridViewTopRowIndex As Integer

        Public Sub DisconnectRDS(SetEventID As Integer)

            EventID = SetEventID
            RDSCopy = New AbovoRangeDataSource
            RDSCopy.Range = MasterRDS.Range
            RDSCopy.DSTag = MasterRDS.DSTag
            AttachedGridControl.DataSource = Nothing
            MasterRDS.Range = Nothing
            MasterRDS = Nothing
            Try

                GridViewTopRowIndex = AttachedGridView.TopRowIndex

            Catch ex As Exception

                GridViewTopRowIndex = 0

            End Try


        End Sub
        Public Sub ReConnectRDS(CalledEventID As Integer)

            'EventID = CalledEventID
            RDSCopy = New AbovoRangeDataSource
            RDSCopy.Range = MasterRDS.Range
            RDSCopy.DSTag = MasterRDS.DSTag

        End Sub
        Public Property CurrRDSStatus As Integer
            Get
                Return RDSStatus
            End Get
            Set(Setvalue As Integer)
                RDSStatus = Setvalue
            End Set
        End Property
        Public Sub New(SetRDSID As Integer, SetNR As String, SetRDS As AbovoRangeDataSource, SetAttachedGrid As DevExpress.XtraGrid.GridControl)

            AttachedGridControl = SetAttachedGrid
            AttachedGridView = AttachedGridControl.FocusedView
            MasterRDS = SetRDS
            Me.RDSID = RDSID
            RDSStatus = 1
            Me.LiveNR = SetNR

        End Sub


    End Class
End Namespace