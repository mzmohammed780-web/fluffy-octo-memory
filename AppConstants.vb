Public Module AppConstants

    ' ── أسماء الجداول ───────────────────────────────────────────
    Public Const Table_Players As String = "Players"
    Public Const Table_Archive As String = "archive"
    Public Const Table_Expenses As String = "Expenses"
    Public Const Table_Payments As String = "Revenues"   ' 🌟 أُعيدت التسمية من Payment — اسم أدق (إيرادات)
    Public Const Table_Users As String = "Users"


    ' ── مصادر البيانات (DataSource) ──────────────────────────────
    Public Const Source_Main As String = "Main"
    Public Const Source_Archive As String = "Archive"

    ' ── أعمدة مخصصة في الداتا جريد ──────────────────────────────
    Public Const Grid_SelectColumn As String = "SelectColumn"
    Public Const Grid_SeqColumn As String = "SeqColumn"
    Public Const Grid_DataSource As String = "DataSource"

    ' ── أسماء أعمدة جدول اللاعبين ───────────────────────────────
    Public Const Col_PlayerId As String = "Playerid"
    Public Const Col_PlayerName As String = "Playername"
    Public Const Col_BirthDate As String = "birthdate"
    Public Const Col_FatherName As String = "fathername"
    Public Const Col_FatherId As String = "fatherid"
    Public Const Col_WifeName As String = "wifename"
    Public Const Col_WifeId As String = "wifeid"
    Public Const Col_Fone As String = "fone"
    Public Const Col_AltFone As String = "altfone"
    Public Const Col_Address As String = "address"
    Public Const Col_DateIn As String = "datein"
    Public Const Col_Notes As String = "Notes"
    Public Const Col_Gender As String = "gender"
    Public Const Col_MaritalStatus As String = "MaritalStatus"
    Public Const Col_Team As String = "Team"
    Public Const Col_Rolle As String = "Rolle"
    Public Const Col_JobTitle As String = "JobTitle"
    Public Const Col_PlayerPhoto As String = "playerphoto"
    Public Const Col_DeletedDate As String = "DeletedDate"
    Public Const Col_LastModifiedBy As String = "LastModifiedBy"

    ' ── أسماء أعمدة جدول المصروفات ──────────────────────────────
    Public Const Col_Exp_ID As String = "ID"
    Public Const Col_Exp_VoucherNumber As String = "VoucherNumber"
    Public Const Col_Exp_Description As String = "Description"
    Public Const Col_Exp_Amount As String = "Amount"
    Public Const Col_Exp_CurrencyType As String = "CurrencyType"
    Public Const Col_Exp_ExpenseDate As String = "ExpenseDate"
    Public Const Col_Exp_Category As String = "Category"
    Public Const Col_Exp_Notes As String = "Notes"
    Public Const Col_Exp_LastModifiedBy As String = "LastModifiedBy"

    ' ── أسماء أعمدة جدول المدفوعات ──────────────────────────────
    Public Const Col_Pay_ID As String = "ID"
    Public Const Col_Pay_PlayerName As String = "Playername"
    Public Const Col_Pay_PaymentDate As String = "PaymentDate"
    Public Const Col_Pay_VoucherNumber As String = "VoucherNumber"
    Public Const Col_Pay_Amount As String = "Amount"
    Public Const Col_Pay_PaymentMethod As String = "PaymentMethod"
    Public Const Col_Pay_TransferorName As String = "TransferorName"
    Public Const Col_Pay_Notes As String = "Notes"

    Public Const Col_Pay_DueDate As String = "DueDate"
    Public Const Col_Pay_CurrencyType As String = "CurrencyType"
    Public Const Col_Pay_CreatedDate As String = "CreatedDate"
    Public Const Col_Pay_LastModifiedBy As String = "LastModifiedBy"
    Public Const Col_Pay_PlayerId As String = "PlayerId"
    ' ── أسماء أعمدة جدول المستخدمين ─────────────────────────────

    Public Const Col_Username As String = "Username"
    Public Const Col_PasswordHash As String = "PasswordHash"
    Public Const Col_Salt As String = "Salt"
    Public Const Col_CanEdit As String = "CanEdit"
    Public Const Col_CanDelete As String = "CanDelete"
    Public Const Col_MustChangePassword As String = "MustChangePassword"

    ' ── ثوابت العملات الموحدة ───────────────────────────────────
    Public Const Currency_ILS As String = "شيكل"
    Public Const Currency_USD As String = "دولار"
    Public Const Currency_JOD As String = "دينار"



    ' ── ثوابت طرق الدفع ─────────────────────────────────────────
    Public Const PayMethod_Cash As String = "نقداً"
    Public Const PayMethod_Transfer As String = "تحويل بنكي"
    Public Const PayMethod_Cheque As String = "شيك"


End Module
