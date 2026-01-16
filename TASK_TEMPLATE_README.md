# タスクテンプレートシステム

統合タスクハンドラーにより、task3.py, task5.py, task6.py, task7.pyを1つのファイルで処理できるようになりました。

## ファイル構成

- **task_handler.py**: 統合タスクハンドラー（全タスクタイプを処理）
- **task_config.json**: タスク設定ファイル（テンプレートとマッピング定義）

## 設定構造 (task_config.json)

各タスクタイプには以下の設定が可能：

```json
{
  "タスク番号": {
    "name": "タスク名",
    "template": "テンプレートファイル名.xlsx",
    "title_replace": "タイトルから削除する文字列",
    "output_suffix": "出力ファイル名のサフィックス",
    "cell_mappings": {
      "セル位置": {
        "type": "マッピングタイプ",
        "key": "フィールド名",
        "format": "フォーマッタ名"
      }
    },
    "custom_processor": "カスタム処理メソッド名",
    "operation_records_start": 操作記録開始行,
    "operation_records_style": "simple/detailed",
    "operation_records_title": "操作記録のタイトル",
    "print_area": true/false,
    "print_mode": "a4/a4_singleside/conditional"
  }
}
```

### マッピングタイプ

- `business_id`: 審批編号
- `title_modified`: タイトル（title_replace適用後）
- `originator_dept`: 部門名
- `field`: フォームフィールドの値

### フォーマッタ

- `goods_detail`: 商品明細フォーマット
- `invoice_attachments`: 発票添付ファイルフォーマット

### 操作記録スタイル

- `simple`: 同意のみ表示
- `detailed`: 全操作を詳細表示

## 新しいタスクタイプの追加方法

1. **task_config.json**に新しいタスク設定を追加
2. 特殊な処理が必要な場合は**task_handler.py**にカスタムメソッドを追加
3. **main.py**の`ALL_TASKS`に新しいタスクを登録

## 使用例

```python
import task_handler

# タスク処理
result = task_handler.handle_task(
    pid="task_id",
    task_type=3,
    status="COMPLETED",
    title="タスクタイトル"
)
```

## 利点

1. **コードの重複削減**: 共通処理を1箇所で管理
2. **保守性向上**: 新しいタスクタイプを設定ファイルで追加可能
3. **柔軟性**: カスタム処理を必要に応じて追加可能
4. **統一性**: 全タスクで一貫した処理フロー
