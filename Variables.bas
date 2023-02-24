Option Compare Database

'ログインユーザID
Public lngLoginUserID As Long

'ユーザ種類
Public Enum EnumUserType
    Admin = 1
    Normal = 2
End Enum

'ユーザ種類
Public Enum EnumAttendanceType
    Attendance = 1
    Leaving = 2
End Enum

'ロックされているか
Public Enum EnumIsLock
    NotLock = 0
    Locked = 1
End Enum

'承認フラグ
Public Enum EnumIsApploval
    NotApploval = 0
    Applovaled = 1
End Enum
