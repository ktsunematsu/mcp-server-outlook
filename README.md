# MCP Server for Outlook Calendar

Windows専用のMCPサーバーで、ローカルのMicrosoft Outlookアプリケーションのカレンダーに直接アクセスします。COM Interopを使用しているため、Azure ADの設定やインターネット接続は不要です。

## 特徴

- ✅ **Azure不要** - COM Interop経由でローカルOutlookに直接アクセス
- ✅ **完全オフライン** - インターネット接続不要
- ✅ **認証不要** - ローカルのOutlookアプリを使用
- ⚠️ **Windows専用** - COM技術を使用するためWindowsでのみ動作
- ⚠️ **Outlookインストール必須** - Microsoft Outlookがインストールされている必要があります

## 必要環境

- **OS**: Windows 10/11
- **ソフトウェア**:
  - Microsoft Outlook (デスクトップ版)
  - Node.js 18以上
  - PowerShell 5.0以上 (Windows標準搭載)

## インストール

```bash
cd mcp-server-outlook
npm install
npm run build
```

## 設定方法

### GitHub Copilot / Claude Desktop での設定

設定ファイルに以下を追加:

**Windows**
- Claude Desktop: `%APPDATA%\Claude\claude_desktop_config.json`
- GitHub Copilot: VSCode設定

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["C:\\path\\to\\mcp-server-outlook\\build\\index.js"]
    }
  }
}
```

**注意**: パスは絶対パスで指定してください。

## 利用可能なツール

### 1. `outlook_list_events`
カレンダーの予定一覧を取得

**パラメータ**:
- `startDate` (オプション): 開始日時 (ISO 8601形式)
- `endDate` (オプション): 終了日時 (ISO 8601形式)

**例**:
```
今週の予定を表示して
```

### 2. `outlook_get_event`
特定の予定の詳細を取得

**パラメータ**:
- `eventId` (必須): イベントのEntryID

**例**:
```
このイベントの詳細を教えて: [eventId]
```

### 3. `outlook_create_event`
新しい予定を作成

**パラメータ**:
- `subject` (必須): 件名
- `start` (必須): 開始日時
- `end` (必須): 終了日時
- `body` (オプション): 本文
- `location` (オプション): 場所
- `attendees` (オプション): 参加者のメールアドレス配列
- `isAllDay` (オプション): 終日イベントかどうか

**例**:
```
明日の14時から15時まで「チームミーティング」という予定を作成して
```

### 4. `outlook_update_event`
既存の予定を更新

**パラメータ**:
- `eventId` (必須): イベントのEntryID
- `subject` (オプション): 新しい件名
- `start` (オプション): 新しい開始日時
- `end` (オプション): 新しい終了日時
- `body` (オプション): 新しい本文
- `location` (オプション): 新しい場所

**例**:
```
このイベントの場所を「会議室A」に変更して
```

### 5. `outlook_delete_event`
予定を削除

**パラメータ**:
- `eventId` (必須): イベントのEntryID

**例**:
```
このイベントを削除して: [eventId]
```

### 6. `outlook_search_events`
キーワードで予定を検索

**パラメータ**:
- `query` (必須): 検索キーワード

**例**:
```
「ミーティング」を含む予定を検索して
```

## トラブルシューティング

### Outlookに接続できない

**症状**: "Failed to connect to Outlook"エラー

**解決策**:
1. Outlookアプリが起動しているか確認
2. Outlookが正しくインストールされているか確認
3. PowerShellの実行ポリシーを確認:
   ```powershell
   Get-ExecutionPolicy
   ```
   必要に応じて変更:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

### PowerShellスクリプトが実行できない

**症状**: "実行ポリシー"エラー

**解決策**:
管理者権限でPowerShellを開き、以下を実行:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### イベントIDが見つからない

**症状**: "Event not found"エラー

**解決策**:
- `outlook_list_events`で最新のイベント一覧を取得
- 正しいEntryIDを使用しているか確認

## プロジェクト構造

```
mcp-server-outlook/
├── src/
│   ├── index.ts                    # MCPサーバーのエントリポイント
│   └── outlook/
│       ├── powershell-bridge.ts    # PowerShell実行モジュール
│       └── calendar-client.ts      # Outlookカレンダークライアント
├── scripts/
│   └── outlook-calendar.ps1        # PowerShell COM操作スクリプト
├── package.json
├── tsconfig.json
└── README.md
```

## 技術詳細

### アーキテクチャ

1. **Node.js (MCP Server)** → `child_process.spawn()`
2. **PowerShell Script** → COM Interop
3. **Outlook COM Object** → ローカルカレンダー

### 制限事項

- ✗ Windows以外のOSでは動作しません
- ✗ Outlookがインストールされていない場合は使用不可
- ✗ Outlook Web版 (outlook.com) には対応していません
- ✗ オフライン時、Outlookが起動していない場合はエラーになる可能性があります

## ライセンス

MIT

## 比較: ryaker/outlook-mcp との違い

| 機能 | mcp-server-outlook (本実装) | ryaker/outlook-mcp |
|------|---------------------------|-------------------|
| 認証 | 不要 (COM) | Azure AD必須 |
| Email機能 | ✗ | ✅ |
| カレンダー機能 | ✅ | ✅ |
| オフライン動作 | ✅ | ✗ |
| Azure設定 | 不要 | 必須 |
| 対応OS | Windows のみ | クロスプラットフォーム |
| 実装難易度 | 中 | 低 |

**こちらを選ぶべき場合**:
- Azure Portalへのアクセス権がない
- 完全にローカルで動作させたい
- カレンダー機能のみで十分

**ryaker/outlook-mcpを選ぶべき場合**:
- Email機能も必要
- Azure設定が可能
- クロスプラットフォーム対応が必要
