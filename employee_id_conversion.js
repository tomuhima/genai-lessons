// =============================
// 社員マスター設定（LINE UserID → 社員名）
// 社員が増えた場合はここに追加してください
// =============================
const EMPLOYEE_MAP = {
  "Uxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx": "田中太郎",
  "Uyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy": "山田花子",
  "Uc1e1e958fb952752c3d34628a17585a1": "野添優"
  // 追加例:
  // "Uzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz": "鈴木一郎",
};
// =============================

const userId = $json.userId;
const employeeName = EMPLOYEE_MAP[userId];

if (!employeeName) {
  return [{
    json: {
      ...$json,
      valid: false,
      errorMessage: `登録されていないLINEアカウントです。\n管理者に連絡してください。\n(UserID: ${userId})`
    }
  }];
}

return [{
  json: {
    ...$json,
    employeeName
  }
}];
