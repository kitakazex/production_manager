# Ver.1.0α Phase 1 データ構造

## production_orders

生産計画を「製番 × 組立工程」単位で管理します。
一意キーは `(production_no, assembly_process)` です。

## parts

部品の設計・構成情報を管理します。

- production_order_id
- part_no
- part_name
- quantity（整数）
- drawing_no
- material
- notes

一意キーは `(production_order_id, part_no)` です。

## procurements

部品ごとの材料手配・加工手配・購入品手配を同じテーブルで管理します。

- part_id
- process_order（10, 20, 30...）
- procurement_type（材料、加工、購入品など）
- processing_details
- vendor
- due_date
- issue_date
- actual_received_date
- status
- notes

一意キーは `(part_id, process_order)` です。

## 旧データの移行

旧 `bom_items` の各行は次のように変換されます。

- 部品情報 → `parts`
- 発注先・納期・状態 → `procurements`
- 工程順 → 10
- 手配区分 → 加工
- 加工内容 → 従来手配
