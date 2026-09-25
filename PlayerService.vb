Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.Threading
Imports System.Threading.Tasks

Public Class PlayerService

    ''' <summary>التحقق من وجود اللاعب مسبقاً</summary>
    Public Async Function PlayerExistsAsync(playerID As String) As Task(Of Boolean)
        Dim n As Long
        If Not Long.TryParse(playerID, n) Then Return False
        Dim result As Object = Await DatabaseModule.ExecuteScalarAsync(
            $"SELECT COUNT(*) FROM {AppConstants.Table_Players} WHERE {AppConstants.Col_PlayerId} = @id",
            New SQLiteParameter("@id", n))
        Return Convert.ToInt32(result) > 0
    End Function

    ''' <summary>إضافة لاعب جديد لقاعدة البيانات</summary>
    Public Async Function InsertPlayerAsync(
        playerIdNum As Long, playerName As String, birthDate As Object,
        fatherName As String, fatherId As Object, wifeName As String, wifeId As Object,
        fone As String, altFone As String, address As String, dateIn As Object,
        notes As String, gender As String, maritalStatus As String,
        team As String, rolle As String, jobTitle As String, photo As Object,
        modifiedBy As String) As Task(Of Integer)

        If Not UserSession.CanEdit Then Return 0

        Dim parameters As SQLiteParameter() = {
            New SQLiteParameter("@Playerid", playerIdNum),
            New SQLiteParameter("@Playername", If(String.IsNullOrWhiteSpace(playerName), DBNull.Value, CObj(playerName.Trim()))),
            New SQLiteParameter("@birthdate", If(birthDate Is Nothing, DBNull.Value, birthDate)),
            New SQLiteParameter("@fathername", If(String.IsNullOrWhiteSpace(fatherName), DBNull.Value, CObj(fatherName.Trim()))),
            New SQLiteParameter("@fatherid", If(fatherId Is Nothing, DBNull.Value, fatherId)),
            New SQLiteParameter("@wifename", If(String.IsNullOrWhiteSpace(wifeName), DBNull.Value, CObj(wifeName.Trim()))),
            New SQLiteParameter("@wifeid", If(wifeId Is Nothing, DBNull.Value, wifeId)),
            New SQLiteParameter("@fone", If(String.IsNullOrWhiteSpace(fone), DBNull.Value, CObj(fone.Trim()))),
            New SQLiteParameter("@altfone", If(String.IsNullOrWhiteSpace(altFone), DBNull.Value, CObj(altFone.Trim()))),
            New SQLiteParameter("@address", If(String.IsNullOrWhiteSpace(address), DBNull.Value, CObj(address.Trim()))),
            New SQLiteParameter("@datein", If(dateIn Is Nothing, DBNull.Value, dateIn)),
            New SQLiteParameter("@Notes", If(String.IsNullOrWhiteSpace(notes), DBNull.Value, CObj(notes.Trim()))),
            New SQLiteParameter("@gender", If(String.IsNullOrWhiteSpace(gender), DBNull.Value, CObj(gender.Trim()))),
            New SQLiteParameter("@MaritalStatus", If(String.IsNullOrWhiteSpace(maritalStatus), DBNull.Value, CObj(maritalStatus.Trim()))),
            New SQLiteParameter("@Team", If(String.IsNullOrWhiteSpace(team), DBNull.Value, CObj(team.Trim()))),
            New SQLiteParameter("@Rolle", If(String.IsNullOrWhiteSpace(rolle), DBNull.Value, CObj(rolle.Trim()))),
            New SQLiteParameter("@JobTitle", If(String.IsNullOrWhiteSpace(jobTitle), DBNull.Value, CObj(jobTitle.Trim()))),
            New SQLiteParameter("@playerphoto", If(photo Is Nothing, DBNull.Value, photo)),
            New SQLiteParameter("@LastModifiedBy", If(String.IsNullOrWhiteSpace(modifiedBy), DBNull.Value, CObj(modifiedBy)))
        }

        Dim sql As String =
            $"INSERT INTO {AppConstants.Table_Players} ({AppConstants.Col_PlayerId}, {AppConstants.Col_PlayerName}, {AppConstants.Col_BirthDate}, {AppConstants.Col_FatherName}, {AppConstants.Col_FatherId}, {AppConstants.Col_WifeName}, {AppConstants.Col_WifeId}, " &
            $"{AppConstants.Col_Fone}, {AppConstants.Col_AltFone}, {AppConstants.Col_Address}, {AppConstants.Col_DateIn}, {AppConstants.Col_Notes}, {AppConstants.Col_Gender}, {AppConstants.Col_MaritalStatus}, {AppConstants.Col_Team}, {AppConstants.Col_Rolle}, {AppConstants.Col_JobTitle}, {AppConstants.Col_PlayerPhoto}, {AppConstants.Col_LastModifiedBy}) " &
            "VALUES (@Playerid, @Playername, @birthdate, @fathername, @fatherid, @wifename, @wifeid, " &
            "@fone, @altfone, @address, @datein, @Notes, @gender, @MaritalStatus, @Team, @Rolle, @JobTitle, @playerphoto, @LastModifiedBy)"

        Dim affected As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql, parameters)

        ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
        If affected > 0 Then
            Dim audit As New AuditService()
            Await audit.LogAsync(AuditService.Act_PlayerAdd, "لاعب",
                playerIdNum.ToString(System.Globalization.CultureInfo.InvariantCulture), $"{playerName}")
        End If

        Return affected
    End Function


    ''' <summary>تحديث بيانات لاعب موجود (مع دعم تغيير رقم الهوية)</summary>
    Public Async Function UpdatePlayerAsync(
        newPlayerIdNum As Long, oldPlayerIdNum As Long, playerName As String, birthDate As Object,
        fatherName As String, fatherId As Object, wifeName As String, wifeId As Object,
        fone As String, altFone As String, address As String, dateIn As Object,
        notes As String, gender As String, maritalStatus As String,
        team As String, rolle As String, jobTitle As String,
        photo As Object, imageChanged As Boolean,
        modifiedBy As String) As Task(Of Integer)

        If Not UserSession.CanEdit Then Return 0

        Dim parameters As New List(Of SQLiteParameter) From {
            New SQLiteParameter("@NewPlayerid", newPlayerIdNum), ' 🌟 الرقم الجديد
            New SQLiteParameter("@Playername", If(String.IsNullOrWhiteSpace(playerName), DBNull.Value, CObj(playerName.Trim()))),
            New SQLiteParameter("@birthdate", If(birthDate Is Nothing, DBNull.Value, birthDate)),
            New SQLiteParameter("@fathername", If(String.IsNullOrWhiteSpace(fatherName), DBNull.Value, CObj(fatherName.Trim()))),
            New SQLiteParameter("@fatherid", If(fatherId Is Nothing, DBNull.Value, fatherId)),
            New SQLiteParameter("@wifename", If(String.IsNullOrWhiteSpace(wifeName), DBNull.Value, CObj(wifeName.Trim()))),
            New SQLiteParameter("@wifeid", If(wifeId Is Nothing, DBNull.Value, wifeId)),
            New SQLiteParameter("@fone", If(String.IsNullOrWhiteSpace(fone), DBNull.Value, CObj(fone.Trim()))),
            New SQLiteParameter("@altfone", If(String.IsNullOrWhiteSpace(altFone), DBNull.Value, CObj(altFone.Trim()))),
            New SQLiteParameter("@address", If(String.IsNullOrWhiteSpace(address), DBNull.Value, CObj(address.Trim()))),
            New SQLiteParameter("@datein", If(dateIn Is Nothing, DBNull.Value, dateIn)),
            New SQLiteParameter("@Notes", If(String.IsNullOrWhiteSpace(notes), DBNull.Value, CObj(notes.Trim()))),
            New SQLiteParameter("@gender", If(String.IsNullOrWhiteSpace(gender), DBNull.Value, CObj(gender.Trim()))),
            New SQLiteParameter("@MaritalStatus", If(String.IsNullOrWhiteSpace(maritalStatus), DBNull.Value, CObj(maritalStatus.Trim()))),
            New SQLiteParameter("@Team", If(String.IsNullOrWhiteSpace(team), DBNull.Value, CObj(team.Trim()))),
            New SQLiteParameter("@Rolle", If(String.IsNullOrWhiteSpace(rolle), DBNull.Value, CObj(rolle.Trim()))),
            New SQLiteParameter("@JobTitle", If(String.IsNullOrWhiteSpace(jobTitle), DBNull.Value, CObj(jobTitle.Trim()))),
            New SQLiteParameter("@LastModifiedBy", If(String.IsNullOrWhiteSpace(modifiedBy), DBNull.Value, CObj(modifiedBy))),
            New SQLiteParameter("@OldPlayerid", oldPlayerIdNum) ' 🌟 الرقم القديم للبحث
        }

        Dim sql As String =
            $"UPDATE {AppConstants.Table_Players} SET " &
            $"Playerid=@NewPlayerid, Playername=@Playername, birthdate=@birthdate, fathername=@fathername, " &
            "fatherid=@fatherid, wifename=@wifename, wifeid=@wifeid, fone=@fone, " &
            "altfone=@altfone, address=@address, datein=@datein, Notes=@Notes, " &
            "gender=@gender, MaritalStatus=@MaritalStatus, Team=@Team, Rolle=@Rolle, " &
            $"JobTitle=@JobTitle, {AppConstants.Col_LastModifiedBy}=@LastModifiedBy "

        If imageChanged Then
            sql &= ", playerphoto=@playerphoto "
            parameters.Add(New SQLiteParameter("@playerphoto", If(photo Is Nothing, DBNull.Value, photo)))
        End If

        ' 🌟 البحث بناءً على رقم الهوية القديم
        sql &= "WHERE Playerid=@OldPlayerid"

        ' تغيير الهوية يجب أن يحدّث ربط الإيرادات في المعاملة نفسها.
        ' تنفيذ تحديث اللاعب وحده كان يترك الدفعات مرتبطة بالهوية القديمة.
        If newPlayerIdNum <> oldPlayerIdNum Then
            Dim queries As New List(Of Tuple(Of String, SQLiteParameter()))
            queries.Add(Tuple.Create(
                sql,
                parameters.ToArray()))
            queries.Add(Tuple.Create(
                $"UPDATE {AppConstants.Table_Payments} SET PlayerId = @NewPlayerid WHERE PlayerId = @OldPlayerid",
                New SQLiteParameter() {
                    New SQLiteParameter("@NewPlayerid", newPlayerIdNum),
                    New SQLiteParameter("@OldPlayerid", oldPlayerIdNum)
                }))

            Dim committed As Boolean = Await Task.Run(Function() DatabaseModule.ExecuteTransaction(queries))
            If Not committed Then Return 0

            ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق (مسار تغيير الهوية)
            Dim auditChange As New AuditService()
            Await auditChange.LogAsync(AuditService.Act_PlayerUpdate, "لاعب",
                newPlayerIdNum.ToString(System.Globalization.CultureInfo.InvariantCulture),
                $"{playerName} — تم تغيير الهوية من {oldPlayerIdNum}")

            Dim updatedPlayer As Object = Await DatabaseModule.ExecuteScalarAsync(
                $"SELECT COUNT(*) FROM {AppConstants.Table_Players} WHERE {AppConstants.Col_PlayerId} = @id",
                New SQLiteParameter("@id", newPlayerIdNum))
            Return If(updatedPlayer IsNot Nothing AndAlso Not IsDBNull(updatedPlayer) AndAlso Convert.ToInt32(updatedPlayer) = 1, 1, 0)
        End If

        Dim affected As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql, parameters.ToArray())

        ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
        If affected > 0 Then
            Dim audit As New AuditService()
            Await audit.LogAsync(AuditService.Act_PlayerUpdate, "لاعب",
                newPlayerIdNum.ToString(System.Globalization.CultureInfo.InvariantCulture), $"{playerName}")
        End If

        Return affected
    End Function

    ''' <summary>نقل لاعب إلى الأرشيف دون استبدال سجل موجود</summary>
    Public Function MoveToArchive(playerID As String) As Boolean
        Try
            If Not UserSession.CanDelete Then Return False

            Dim playerIdNum As Long
            If Not Long.TryParse(playerID, playerIdNum) Then Return False

            Dim sourceCount As Object = DatabaseModule.ExecuteScalar(
                $"SELECT COUNT(*) FROM {AppConstants.Table_Players} WHERE {AppConstants.Col_PlayerId} = @id",
                New SQLiteParameter("@id", playerIdNum))
            If Convert.ToInt32(sourceCount) <> 1 Then Return False

            Dim archiveCount As Object = DatabaseModule.ExecuteScalar(
                $"SELECT COUNT(*) FROM {AppConstants.Table_Archive} WHERE {AppConstants.Col_PlayerId} = @id",
                New SQLiteParameter("@id", playerIdNum))
            If Convert.ToInt32(archiveCount) > 0 Then Return False

            Dim queries As New List(Of Tuple(Of String, SQLiteParameter()))

            queries.Add(Tuple.Create(
                $"INSERT INTO {AppConstants.Table_Archive} (Playerid, Playername, birthdate, fathername, fatherid, wifename, wifeid, fone, altfone, address, datein, Notes, gender, MaritalStatus, Team, Rolle, JobTitle, playerphoto, DeletedDate, {AppConstants.Col_LastModifiedBy}) " &
                $"SELECT Playerid, Playername, birthdate, fathername, fatherid, wifename, wifeid, fone, altfone, address, datein, Notes, gender, MaritalStatus, Team, Rolle, JobTitle, playerphoto, date('now'), {AppConstants.Col_LastModifiedBy} FROM {AppConstants.Table_Players} WHERE {AppConstants.Col_PlayerId} = @id",
                New SQLiteParameter() {New SQLiteParameter("@id", playerIdNum)}))

            queries.Add(Tuple.Create(
                $"DELETE FROM {AppConstants.Table_Players} WHERE {AppConstants.Col_PlayerId} = @id",
                New SQLiteParameter() {New SQLiteParameter("@id", playerIdNum)}))

            Dim committed As Boolean = DatabaseModule.ExecuteTransaction(queries)

            ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
            If committed Then
                Dim audit As New AuditService()
                audit.Log(AuditService.Act_PlayerArchive, "لاعب", playerID, "")
            End If

            Return committed
        Catch ex As Exception
            DatabaseModule.LogError("PlayerService.MoveToArchive", ex)
            Return False
        End Try
    End Function

    ''' <summary>استعادة لاعب من الأرشيف دون استبدال لاعب نشط</summary>
    Public Function RestoreFromArchive(playerID As String) As Boolean
        Try
            If Not UserSession.CanEdit Then Return False

            Dim playerIdNum As Long
            If Not Long.TryParse(playerID, playerIdNum) Then Return False

            Dim archiveCount As Object = DatabaseModule.ExecuteScalar(
                $"SELECT COUNT(*) FROM {AppConstants.Table_Archive} WHERE {AppConstants.Col_PlayerId} = @id",
                New SQLiteParameter("@id", playerIdNum))
            If Convert.ToInt32(archiveCount) <> 1 Then Return False

            Dim activeCount As Object = DatabaseModule.ExecuteScalar(
                $"SELECT COUNT(*) FROM {AppConstants.Table_Players} WHERE {AppConstants.Col_PlayerId} = @id",
                New SQLiteParameter("@id", playerIdNum))
            If Convert.ToInt32(activeCount) > 0 Then Return False

            Dim queries As New List(Of Tuple(Of String, SQLiteParameter()))

            queries.Add(Tuple.Create(
                $"INSERT INTO {AppConstants.Table_Players} (Playerid, Playername, birthdate, fathername, fatherid, wifename, wifeid, fone, altfone, address, datein, Notes, gender, MaritalStatus, Team, Rolle, JobTitle, playerphoto, {AppConstants.Col_LastModifiedBy}) " &
                $"SELECT Playerid, Playername, birthdate, fathername, fatherid, wifename, wifeid, fone, altfone, address, datein, Notes, gender, MaritalStatus, Team, Rolle, JobTitle, playerphoto, {AppConstants.Col_LastModifiedBy} " &
                $"FROM {AppConstants.Table_Archive} WHERE {AppConstants.Col_PlayerId} = @id",
                New SQLiteParameter() {New SQLiteParameter("@id", playerIdNum)}))

            queries.Add(Tuple.Create(
                $"DELETE FROM {AppConstants.Table_Archive} WHERE {AppConstants.Col_PlayerId} = @id",
                New SQLiteParameter() {New SQLiteParameter("@id", playerIdNum)}))

            Dim committed As Boolean = DatabaseModule.ExecuteTransaction(queries)

            ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
            If committed Then
                Dim audit As New AuditService()
                audit.Log(AuditService.Act_PlayerRestore, "لاعب", playerID, "")
            End If

            Return committed
        Catch ex As Exception
            DatabaseModule.LogError("PlayerService.RestoreFromArchive", ex)
            Return False
        End Try
    End Function

    ''' <summary>جلب اسم اللاعب بواسطة رقم الهوية (للتحقق من التكرار)</summary>
    Public Async Function GetPlayerNameByIdAsync(playerID As String) As Task(Of String)
        Try
            Dim playerIdNum As Long
            If Not Long.TryParse(playerID, playerIdNum) Then Return ""
            Dim result As Object = Await DatabaseModule.ExecuteScalarAsync(
                $"SELECT {AppConstants.Col_PlayerName} FROM {AppConstants.Table_Players} WHERE {AppConstants.Col_PlayerId} = @id",
                New SQLiteParameter("@id", playerIdNum))
            Return If(result IsNot Nothing AndAlso Not IsDBNull(result), result.ToString(), "")
        Catch ex As Exception
            DatabaseModule.LogError("PlayerService.GetPlayerNameByIdAsync", ex)
            Return ""
        End Try
    End Function

    ''' <summary>جلب صورة اللاعب من قاعدة البيانات باستخدام دالة GetBytes القياسية</summary>
    Public Async Function GetPlayerPhotoAsync(playerID As String, source As String) As Task(Of Byte())
        Try
            If String.IsNullOrWhiteSpace(playerID) Then Return Nothing

            Dim playerIdNum As Long
            If Not Long.TryParse(playerID.Trim(), playerIdNum) Then Return Nothing

            Dim tableName As String = If(source = AppConstants.Source_Archive, AppConstants.Table_Archive, AppConstants.Table_Players)

            Dim sql As String = $"SELECT {AppConstants.Col_PlayerPhoto} FROM {tableName} WHERE {AppConstants.Col_PlayerId} = @id"

            Return Await Task.Run(Function()
                                      Using conn As SQLiteConnection = DatabaseModule.GetConnection()
                                          Using cmd As New SQLiteCommand(sql, conn)
                                              cmd.Parameters.AddWithValue("@id", playerIdNum)
                                              Using reader As SQLiteDataReader = cmd.ExecuteReader(CommandBehavior.SequentialAccess)
                                                  If reader.Read() AndAlso Not reader.IsDBNull(0) Then
                                                      Dim bufferSize As Long = reader.GetBytes(0, 0, Nothing, 0, 0)
                                                      If bufferSize > 0 Then
                                                          Dim bytes(CInt(bufferSize - 1)) As Byte
                                                          reader.GetBytes(0, 0, bytes, 0, CInt(bufferSize))
                                                          Return bytes
                                                      End If
                                                  End If
                                              End Using
                                          End Using
                                      End Using
                                      Return Nothing
                                  End Function)
        Catch ex As Exception
            DatabaseModule.LogError("PlayerService.GetPlayerPhotoAsync", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>تحميل كل اللاعبين النشطين</summary>
    Public Async Function GetAllPlayersAsync() As Task(Of DataTable)
        Dim sql As String =
            "SELECT CAST(Playerid AS TEXT) AS Playerid, Playername, birthdate, fathername, CAST(fatherid AS TEXT) AS fatherid, " &
            "wifename, CAST(wifeid AS TEXT) AS wifeid, fone, altfone, address, datein, Notes, " &
            "gender, MaritalStatus, Team, Rolle, JobTitle, " &
            $"{AppConstants.Col_LastModifiedBy}, " &
            $"'{AppConstants.Source_Main}' AS {AppConstants.Grid_DataSource} FROM {AppConstants.Table_Players} ORDER BY Playername"
        Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql)
        UtilityModule.FormatDateColumnsForDisplay(dt, AppConstants.Col_BirthDate, AppConstants.Col_DateIn)
        Return dt
    End Function

    ''' <summary>تحميل كل اللاعبين المؤرشفين</summary>
    Public Async Function GetAllArchivedPlayersAsync() As Task(Of DataTable)
        Dim sql As String =
            "SELECT CAST(Playerid AS TEXT) AS Playerid, Playername, birthdate, fathername, CAST(fatherid AS TEXT) AS fatherid, " &
            "wifename, CAST(wifeid AS TEXT) AS wifeid, fone, altfone, address, datein, Notes, " &
            "gender, MaritalStatus, Team, Rolle, JobTitle, " &
            $"DeletedDate, '{AppConstants.Source_Archive}' AS {AppConstants.Grid_DataSource}, {AppConstants.Col_LastModifiedBy} FROM {AppConstants.Table_Archive} ORDER BY Playername"
        Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql)
        UtilityModule.FormatDateColumnsForDisplay(dt, AppConstants.Col_BirthDate, AppConstants.Col_DateIn, AppConstants.Col_DeletedDate)
        Return dt
    End Function

    ''' <summary>البحث المطور في اللاعبين والأرشيف معاً</summary>
    Public Async Function SearchPlayersAsync(
        playerID As String, playerName As String, birthDate As String,
        fatherName As String, fatherID As String, wifeName As String, wifeID As String,
        fone As String, altFone As String, address As String, dateIn As String, notes As String,
        gender As String, maritalStatus As String, team As String, rolle As String, jobTitle As String,
        ct As CancellationToken) As Task(Of DataTable)

        Dim bothConditions As New List(Of String)
        Dim mainOnlyConditions As New List(Of String)
        Dim parameters As New List(Of SQLiteParameter)

        ' 🌟 1. حقول تُبحث في كلا الجدولين (اللاعبين والأرشيف)
        AddNumericCondition(bothConditions, parameters, AppConstants.Col_PlayerId, playerID)
        AddLikeCondition(bothConditions, parameters, AppConstants.Col_PlayerName, playerName)
        AddLikeCondition(bothConditions, parameters, AppConstants.Col_FatherName, fatherName)
        AddNumericCondition(bothConditions, parameters, AppConstants.Col_FatherId, fatherID)
        AddLikeCondition(bothConditions, parameters, AppConstants.Col_WifeName, wifeName)
        AddNumericCondition(bothConditions, parameters, AppConstants.Col_WifeId, wifeID)

        ' 🌟 2. حقول تُبحث في جدول اللاعبين النشطين فقط
        AddLikeCondition(mainOnlyConditions, parameters, AppConstants.Col_BirthDate, birthDate)
        AddLikeCondition(mainOnlyConditions, parameters, AppConstants.Col_Fone, fone)
        AddLikeCondition(mainOnlyConditions, parameters, AppConstants.Col_AltFone, altFone)
        AddLikeCondition(mainOnlyConditions, parameters, AppConstants.Col_Address, address)
        AddLikeCondition(mainOnlyConditions, parameters, AppConstants.Col_DateIn, dateIn)
        AddLikeCondition(mainOnlyConditions, parameters, AppConstants.Col_Notes, notes)
        AddExactCondition(mainOnlyConditions, parameters, AppConstants.Col_Gender, gender)
        AddExactCondition(mainOnlyConditions, parameters, AppConstants.Col_MaritalStatus, maritalStatus)
        AddExactCondition(mainOnlyConditions, parameters, AppConstants.Col_Team, team)
        AddExactCondition(mainOnlyConditions, parameters, AppConstants.Col_Rolle, rolle)
        AddExactCondition(mainOnlyConditions, parameters, AppConstants.Col_JobTitle, jobTitle)

        ' دمج الشروط لجدول النشطين
        Dim allMainConditions = bothConditions.Concat(mainOnlyConditions).ToList()
        Dim whereMain As String = If(allMainConditions.Count > 0, " WHERE " & String.Join(" AND ", allMainConditions), "")

        ' إذا لم يتم إدخال أي حقل خاص بالنشطين فقط، نبحث في الأرشيف أيضاً باستخدام شروط الحقول المشتركة
        Dim searchArchive As Boolean = (mainOnlyConditions.Count = 0)
        Dim whereArchive As String = ""

        If searchArchive AndAlso bothConditions.Count > 0 Then
            whereArchive = " WHERE " & String.Join(" AND ", bothConditions)
        End If

        Dim baseSelect As String =
            "CAST(Playerid AS TEXT) AS Playerid, Playername, birthdate, fathername, CAST(fatherid AS TEXT) AS fatherid, " &
            "wifename, CAST(wifeid AS TEXT) AS wifeid, fone, altfone, address, datein, Notes, " &
            $"gender, MaritalStatus, Team, Rolle, JobTitle, {AppConstants.Col_LastModifiedBy} "

        Dim sql As String = ""

        If searchArchive Then
            ' 🌟 البحث في كلا الجدولين
            sql = $"SELECT {baseSelect}, NULL AS DeletedDate, '{AppConstants.Source_Main}' AS {AppConstants.Grid_DataSource} FROM {AppConstants.Table_Players}{whereMain} " &
                  $"UNION ALL " &
                  $"SELECT {baseSelect}, DeletedDate, '{AppConstants.Source_Archive}' AS {AppConstants.Grid_DataSource} FROM {AppConstants.Table_Archive}{whereArchive} " &
                  "ORDER BY Playername"
        Else
            ' 🌟 البحث في جدول النشطين فقط (لأن المستخدم أدخل حقل مخصص للنشطين مثل الجنس أو الهاتف)
            sql = $"SELECT {baseSelect}, NULL AS DeletedDate, '{AppConstants.Source_Main}' AS {AppConstants.Grid_DataSource} FROM {AppConstants.Table_Players}{whereMain} " &
                  "ORDER BY Playername"
        End If

        Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, ct, parameters.ToArray())
        UtilityModule.FormatDateColumnsForDisplay(dt, AppConstants.Col_BirthDate, AppConstants.Col_DateIn, AppConstants.Col_DeletedDate)
        Return dt
    End Function

    ' 🌟 دوال مساعدة لبناء شروط البحث
    Private Sub AddLikeCondition(ByRef conds As List(Of String), ByRef params As List(Of SQLiteParameter), fieldName As String, value As String)
        If String.IsNullOrWhiteSpace(value) Then Return
        conds.Add($"{fieldName} LIKE @{fieldName}")
        params.Add(New SQLiteParameter($"@{fieldName}", "%" & value.Trim() & "%"))
    End Sub

    Private Sub AddExactCondition(ByRef conds As List(Of String), ByRef params As List(Of SQLiteParameter), fieldName As String, value As String)
        If String.IsNullOrWhiteSpace(value) Then Return
        conds.Add($"{fieldName} = @{fieldName}")
        params.Add(New SQLiteParameter($"@{fieldName}", value.Trim()))
    End Sub

    Private Sub AddNumericCondition(ByRef conds As List(Of String), ByRef params As List(Of SQLiteParameter), fieldName As String, value As String)
        If String.IsNullOrWhiteSpace(value) Then Return
        conds.Add($"CAST({fieldName} AS TEXT) LIKE @{fieldName}")
        params.Add(New SQLiteParameter($"@{fieldName}", "%" & value.Trim() & "%"))
    End Sub
    ''' <summary>
    ''' توليد رقم هوية مؤقت تلقائي من نطاق 900000000+ (غير قابل لتصادم مع هويات حقيقية) —
    ''' يعيد أول رقم متاح، أو 0 إذا فشل
    ''' </summary>
    Public Async Function GenerateTempPlayerIdAsync() As Task(Of Long)
        Try
            Const tempRangeStart As Long = 900000000L

            ' 🌟 جلب كل الأرقام المؤقتة المستخدمة حالياً (استعلام واحد بسيط)
            Dim sql As String =
                $"SELECT Playerid FROM {AppConstants.Table_Players} WHERE Playerid >= @start " &
                $"UNION ALL SELECT Playerid FROM {AppConstants.Table_Archive} WHERE Playerid >= @start"

            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql,
                New SQLiteParameter("@start", tempRangeStart))

            Dim used As New HashSet(Of Long)
            If dt IsNot Nothing Then
                For Each row As DataRow In dt.Rows
                    used.Add(Convert.ToInt64(row("Playerid")))
                Next
            End If

            ' 🌟 أول رقم حر من بداية النطاق — الحلقة واقعياً تعطي نتيجة من أول محاولة
            Dim candidate As Long = tempRangeStart + 1
            While used.Contains(candidate)
                candidate += 1
            End While
            Return candidate

        Catch ex As Exception
            DatabaseModule.LogError("PlayerService.GenerateTempPlayerIdAsync", ex)
            Return 0
        End Try
    End Function
End Class
