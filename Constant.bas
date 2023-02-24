Option Compare Database

'メッセージ
Public Const USERNAME_LENGTH_OVER = "ユーザ名は1～255文字以内で指定してください。"
Public Const PASSWORD_LENGTH_OVER = "パスワードは1～255文字以内で指定してください。"
Public Const USER_IS_ADMIN = "既に管理者権限を保有しています。"
Public Const USER_IS_NORMAL = "既に一般ユーザです。"
Public Const USER_IS_LOCKED = "アカウントがロックされています。管理者にお問い合わせください。"
Public Const LOGIN_FAILED = "ログインに失敗しました。"
Public Const USER_REGISTER_FAILED = "ユーザ登録に失敗しました。"
Public Const USER_DELETE_FAILED = "ユーザ登録に失敗しました。"
Public Const USER_UPDATE_FAILED = "ユーザ更新に失敗しました。"
Public Const AUTHORITY_UPDATE_FAILED = "権限更新に失敗しました。"
Public Const LOCK_UPDATE_FAILED = "ロック状態の更新に失敗しました。"
Public Const ATTENDANCE_INSERT_FAILED = "勤怠打刻に失敗しました。"
Public Const CSV_OUTPUT_FAILED = "CSV出力に失敗しました。"
Public Const USER_REGISTER_SUCCESS = "ユーザ登録に成功しました。"
Public Const USER_UPDATE_SUCCESS = "ユーザ更新に成功しました。"
Public Const AUTHORITY_UPDATE_SUCCESS = "権限更新に成功しました。"
Public Const LOCK_UPDATE_SUCCESS = "ロック状態の更新に成功しました。"
Public Const ATTENDANCE_INSERT_SUCCESS = "勤怠打刻に成功しました。"
Public Const CSV_OUTPUT_SUCCESS = "CSV出力に成功しました。"

'フォーム
Public Const MAIN_FORM = "F_Main"
Public Const LOGIN_FORM = "F_Login"
Public Const USER_REGISTER_FORM = "F_UserRegister"
Public Const USERS_FORM = "F_Users"
Public Const USER_UPDATE_FORM = "F_UserUpdate"
Public Const LOCK_USERS_FORM = "F_LockUsers"
Public Const ATTENDANCES_FORM = "F_Attendances"

'その他
Public Const USERNAME_MAX_LENGTH = 255
Public Const PASSWORD_MAX_LENGTH = 255
Public Const CONFIRM = "確認"
Public Const LOCK_COUNT = 3
Public Const ATTENDANCE_TABLE_NAME = "Attendances"
Public Const CSV_FILE_NAME = "勤怠リスト"
