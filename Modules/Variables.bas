Attribute VB_Name = "Variables"
Public adoConn As ADODB.Connection
Public scadaConn As ADODB.Connection
Public TotalShift As Variant
Public TotalDaily As Variant
Public showName As Boolean
Public showProperties As Boolean
Public showCustomer As Boolean
Public showMachine As Boolean
Public showUnitRatio As Boolean
Public showTotalIndex As Boolean
Public showComments As Boolean
Public pUnit As String
Public sUnit As String
Public headerAddress As String
Public Const regPath As String = "HKEY_CURRENT_USER\Software\Prod_sched\"
Public roastingBatches As Collection
Public scheduleMode As Integer '1-roasting, 2-grinding, 3-packing, 4-palety
Public currentSchedule As clsSchedule
Public comparativeSchedule As clsSchedule
