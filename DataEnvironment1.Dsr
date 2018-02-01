VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DataEnvironment1 
   ClientHeight    =   9495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15615
   _ExtentX        =   27543
   _ExtentY        =   16748
   FolderFlags     =   1
   TypeLibGuid     =   "{0653B042-9D5C-4A2B-89B7-5573BC1A28B9}"
   TypeInfoGuid    =   "{588176EF-C36B-4FE9-883A-4AED5B8BEA94}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Connection1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=db_YouthFestival;Data Source=HP-PC"
      Expanded        =   -1  'True
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   10
   BeginProperty Recordset1 
      CommandName     =   "cmdSchool"
      CommDispId      =   1002
      RsDispId        =   1007
      CommandText     =   "select * from tbl_District"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "district_id"
         Caption         =   "district_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "district_name"
         Caption         =   "district_name"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "cmdSchool1"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "select * from tbl_School"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdSchool"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "school_id"
         Caption         =   "school_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "district_id"
         Caption         =   "district_id"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "school_name"
         Caption         =   "school_name"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "address"
         Caption         =   "address"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "contact_no1"
         Caption         =   "contact_no1"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "contact_no2"
         Caption         =   "contact_no2"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   130
         Name            =   "email"
         Caption         =   "email"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "district_id"
         ChildField      =   "district_id"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "cmdProgram"
      CommDispId      =   1008
      RsDispId        =   1013
      CommandText     =   "select * from tbl_program"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_id"
         Caption         =   "program_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "program_name"
         Caption         =   "program_name"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "program_description"
         Caption         =   "program_description"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_time"
         Caption         =   "program_time"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "cmdJudges"
      CommDispId      =   1014
      RsDispId        =   1020
      CommandText     =   "select * from tbl_judges,tbl_program where tbl_judges.program_id=tbl_program.program_id"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   12
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "judge_id"
         Caption         =   "judge_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_id"
         Caption         =   "program_id"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "name"
         Caption         =   "name"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "address"
         Caption         =   "address"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "qualification"
         Caption         =   "qualification"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "experience"
         Caption         =   "experience"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "contact_no"
         Caption         =   "contact_no"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "email"
         Caption         =   "email"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_id"
         Caption         =   "program_id"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "program_name"
         Caption         =   "program_name"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "program_description"
         Caption         =   "program_description"
      EndProperty
      BeginProperty Field12 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_time"
         Caption         =   "program_time"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "cmdMemberSchoolwise"
      CommDispId      =   1021
      RsDispId        =   1026
      CommandText     =   "select * from tbl_member_registration where school_id=?"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "member_registration_id"
         Caption         =   "member_registration_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "school_id"
         Caption         =   "school_id"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "name"
         Caption         =   "name"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "address"
         Caption         =   "address"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "contact_no"
         Caption         =   "contact_no"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dob"
         Caption         =   "dob"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "gender"
         Caption         =   "gender"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "member_category"
         Caption         =   "member_category"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "g_name"
         Caption         =   "g_name"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "relationship"
         Caption         =   "relationship"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   4
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "cmdMemberPgmwise"
      CommDispId      =   1027
      RsDispId        =   1032
      CommandText     =   $"DataEnvironment1.dsx":0000
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   14
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "member_registration_id"
         Caption         =   "member_registration_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "school_id"
         Caption         =   "school_id"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "name"
         Caption         =   "name"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "address"
         Caption         =   "address"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "contact_no"
         Caption         =   "contact_no"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dob"
         Caption         =   "dob"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "gender"
         Caption         =   "gender"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "member_category"
         Caption         =   "member_category"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "g_name"
         Caption         =   "g_name"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "relationship"
         Caption         =   "relationship"
      EndProperty
      BeginProperty Field11 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "MemberPgm_Id"
         Caption         =   "MemberPgm_Id"
      EndProperty
      BeginProperty Field12 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "MemberReg_Id"
         Caption         =   "MemberReg_Id"
      EndProperty
      BeginProperty Field13 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Program_Id"
         Caption         =   "Program_Id"
      EndProperty
      BeginProperty Field14 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "DistLevelPrizeCat_Id"
         Caption         =   "DistLevelPrizeCat_Id"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   4
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "cmdScorePgmwise"
      CommDispId      =   1033
      RsDispId        =   1043
      CommandText     =   $"DataEnvironment1.dsx":009A
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   16
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "score_entering_id"
         Caption         =   "score_entering_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_id"
         Caption         =   "program_id"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "student_id"
         Caption         =   "student_id"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "judge_id"
         Caption         =   "judge_id"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "score"
         Caption         =   "score"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "member_registration_id"
         Caption         =   "member_registration_id"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "school_id"
         Caption         =   "school_id"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "name"
         Caption         =   "name"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "address"
         Caption         =   "address"
      EndProperty
      BeginProperty Field10 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "contact_no"
         Caption         =   "contact_no"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dob"
         Caption         =   "dob"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "gender"
         Caption         =   "gender"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "member_category"
         Caption         =   "member_category"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "g_name"
         Caption         =   "g_name"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "relationship"
         Caption         =   "relationship"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "judge"
         Caption         =   "judge"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   4
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "cmdScore"
      CommDispId      =   1044
      RsDispId        =   1049
      CommandText     =   "select * from tbl_program"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_id"
         Caption         =   "program_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "program_name"
         Caption         =   "program_name"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "program_description"
         Caption         =   "program_description"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_time"
         Caption         =   "program_time"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset9 
      CommandName     =   "cmdScore1"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"DataEnvironment1.dsx":01E8
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdScore"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   16
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "score_entering_id"
         Caption         =   "score_entering_id"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_id"
         Caption         =   "program_id"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "student_id"
         Caption         =   "student_id"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "judge_id"
         Caption         =   "judge_id"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "score"
         Caption         =   "score"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "member_registration_id"
         Caption         =   "member_registration_id"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "school_id"
         Caption         =   "school_id"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "name"
         Caption         =   "name"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "address"
         Caption         =   "address"
      EndProperty
      BeginProperty Field10 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "contact_no"
         Caption         =   "contact_no"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dob"
         Caption         =   "dob"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "gender"
         Caption         =   "gender"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "member_category"
         Caption         =   "member_category"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "g_name"
         Caption         =   "g_name"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "relationship"
         Caption         =   "relationship"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "judge"
         Caption         =   "judge"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "program_id"
         ChildField      =   "program_id"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset10 
      CommandName     =   "cmdPgmTiming"
      CommDispId      =   1050
      RsDispId        =   1056
      CommandText     =   $"DataEnvironment1.dsx":0311
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   19
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "PgmTimingId"
         Caption         =   "PgmTimingId"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Pgm_id"
         Caption         =   "Pgm_id"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Member_id"
         Caption         =   "Member_id"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Schedule_id"
         Caption         =   "Schedule_id"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Time"
         Caption         =   "Time"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_id"
         Caption         =   "program_id"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "program_name"
         Caption         =   "program_name"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "program_description"
         Caption         =   "program_description"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "program_time"
         Caption         =   "program_time"
      EndProperty
      BeginProperty Field10 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "member_registration_id"
         Caption         =   "member_registration_id"
      EndProperty
      BeginProperty Field11 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "school_id"
         Caption         =   "school_id"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "name"
         Caption         =   "name"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "address"
         Caption         =   "address"
      EndProperty
      BeginProperty Field14 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "contact_no"
         Caption         =   "contact_no"
      EndProperty
      BeginProperty Field15 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dob"
         Caption         =   "dob"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "gender"
         Caption         =   "gender"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "member_category"
         Caption         =   "member_category"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "g_name"
         Caption         =   "g_name"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "relationship"
         Caption         =   "relationship"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DataEnvironment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
