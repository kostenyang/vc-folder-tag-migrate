# vCenter 中繼資料搬家（Folder / Tag / Custom Attribute / Notes）

把 **vCenter A** 的 Folder 結構、Tag、Custom Attribute、VM Notes 匯出成 CSV（落地本機），
再匯入 **vCenter B**。兩支腳本，PowerCLI 13.5，`pwsh` 執行。

```
Export-VcMeta.ps1   來源 vC  →  CSV
Import-VcMeta.ps1   CSV      →  目標 vC
```

## 1. 匯出（來源 vCenter）

```bash
pwsh -File .\Export-VcMeta.ps1 -Server <來源vC> -User administrator@vsphere.local -Password '<pw>' -OutDir .\export-A
```

參數：
- `-Datacenter`：只匯出指定 Datacenter（可多個）
- `-Folder 'Linux','MGMT/Prod'`：**只搬指定 VM 資料夾(含子樹)**與裡面的 VM；連帶只帶出這些物件用到的 tag/屬性定義（`-AllDefinitions` 可全帶）
- `-Include`：`Folders,VMPlacement,Tags,CustomAttributes,Notes`（預設全要；用 `-Command` 呼叫才能傳多值）

產出 CSV（UTF-8 with BOM，Excel 可直接開、可手動編修後再匯入）：

| 檔案 | 內容 |
|---|---|
| `folders.csv` | Folder 樹（Datacenter / 類型 VM,HostAndCluster,Datastore,Network / 相對路徑） |
| `vm-placement.csv` | 每台 VM/範本 在哪個 Folder（含 InstanceUuid、IsTemplate、InVApp） |
| `tag-categories.csv` | Tag 分類（Cardinality、可套用的 EntityType） |
| `tags.csv` | 標籤（分類 / 名稱 / 說明） |
| `tag-assignments.csv` | 誰被貼了什麼標籤 |
| `custom-attributes.csv` | 自訂屬性定義（含 TargetType，Global 代表全域） |
| `custom-attribute-values.csv` | 每個物件的自訂屬性值 |
| `notes.csv` | VM Notes（附註） |

## 2. 匯入（目標 vCenter）

**一定先跑 `-DryRun`**，看報告確認比對結果再正式套用。

```bash
# 試跑
pwsh -Command "& .\Import-VcMeta.ps1 -Server <目標vC> -User administrator@vsphere.local -Password '<pw>' -InDir .\export-A -DryRun"

# 正式（不搬 VM，只建資料夾/標籤/屬性/Notes）
pwsh -Command "& .\Import-VcMeta.ps1 -Server <目標vC> -User administrator@vsphere.local -Password '<pw>' -InDir .\export-A"

# 連 VM 也搬進對應資料夾
pwsh -Command "& .\Import-VcMeta.ps1 -Server <目標vC> -User administrator@vsphere.local -Password '<pw>' -InDir .\export-A -MoveVMs"
```

參數：
- `-DatacenterMap 'DC-A=DC-B'`：兩邊 Datacenter 名稱不同時做對應（可多組）
- `-Include`：只做某幾類（預設 `Folders,Tags,CustomAttributes,Notes`，**不含**搬 VM）
- `-Folder 'Linux'`：同上，**用整台 vC 的匯出檔也能只匯入某個資料夾**；範圍外的 VM Notes / 指派會被濾掉
- `-MoveVMs`：把 VM/範本 `Move-VM`/`Move-Template` 到對應資料夾（僅改 inventory 位置，不動儲存/運算）
- `-ReportPath`：明細報告位置（預設寫在 `-InDir` 下 `import-report-<時間>.csv`）

## 3. 比對規則

| 物件 | 比對方式 |
|---|---|
| VM / 範本 | 先 `InstanceUuid`，找不到再用**名稱**；名稱重複則跳過並記錄 |
| Host / Datastore / Cluster / RP / DS Cluster / DPortgroup | 名稱 |
| Folder | Datacenter + 類型 + 相對路徑 |
| Datacenter | 名稱（可用 `-DatacenterMap` 改對應） |

全程 **idempotent**：已存在的資料夾/標籤/屬性/值不會重建，重複跑安全。

## 4. 已知限制

- **vApp 內的 VM**：`InVApp=True`，不會被搬進資料夾（vApp 本身要另外處理）；但它的 Tag / 屬性 / Notes 照樣會套。
- **Notes 是覆蓋**，不是附加；目標端已有相同內容則跳過。
- Tag 分類的 `Cardinality`／`EntityType` 若目標端已有同名分類，**沿用目標端既有設定**、不會改。
- 只搬 inventory 中繼資料，**不搬** 權限(Permission)、Role、Alarm、Resource Pool 結構、DRS/HA 規則。
- 來源若有無法存取的 datastore，`Get-TagAssignment` 一次撈會中斷 → 腳本已改成分批 / 逐一 fallback。

## 5. 驗證紀錄

**回測（同機）**：2026-09-02 對 `10.0.0.101`（vCenter 8.0.3）匯出 → 同機 DryRun 匯入，
16 folders / 17 tags / 13 分類 / 6 指派 / 21 屬性 / 21 Notes / 29 VM 位置 **全部比對命中、零差異**。

**正式搬移**：`10.0.0.101` → `10.0.1.19`（`-DatacenterMap 'Datacenter=m01-dc01'`）：
建 15 folders、13 tag 分類、17 tags、13 自訂屬性、1 屬性值；重跑 DryRun 全為 `Exists/AlreadySet`（冪等確認）。

**端到端功能測試**：來源建 `zz-migtest` 一組物件 → 匯出 → 匯入目標 → 獨立查詢驗證 → 雙邊清除，7/7 PASS：

| 測項 | 結果 |
|---|---|
| 兩層資料夾 `ZZ-MigTest/Level2`（含父子關係） | PASS |
| Tag 分類（Cardinality=Single、EntityType=All、中文說明） | PASS |
| 標籤 + 中文說明 | PASS |
| **標籤貼在資料夾上**（用路徑比對，不是名稱） | PASS |
| 自訂屬性定義（Global） | PASS |
| 屬性值中文 `測試值-2026` | PASS |
| VM Notes：**手動加一列 CSV、不填 UUID**（測名稱 fallback） | PASS |

清除後雙邊複驗均 `無殘留`。

## 6. 自我測試 `Test-VcMeta.ps1`

一鍵端到端測試：來源建測試物件 → 匯出 → DryRun → 匯入 → **獨立查詢驗證** → 兩邊清除。
全過 exit code 0，任何一項失敗 exit 1（可掛 CI）。

```bash
pwsh -Command "& .\Test-VcMeta.ps1 -SourceServer <來源vC> -SourcePassword '<pw>' -SourceDatacenter <來源DC> -TargetServer <目標vC> -TargetPassword '<pw>' -TargetDatacenter <目標DC> -NotesTestVM <目標端一台測試VM>"
```

- 只會動 `-Prefix`（預設 `zz-migtest`）開頭的物件，以及 `-NotesTestVM` 那台 VM 的 Notes（測完清空）
- `-KeepTestObjects`：保留測試物件不清，方便到 UI 上看
- 21 項檢查：匯出 4、**範圍匯出 4**、DryRun 2、**範圍匯入 2**、匯入 2、目標端獨立驗證 5、清除複驗 2

實測結果（2026-09-02，vCenter 8.0.3 → 9.1）：**21/21 ALL PASS**。

## 7. 以資料夾為單位搬（by folder）

只搬某個資料夾子樹、連同裡面 VM 的 tag / 自訂屬性 / Notes：

```bash
pwsh -Command "& .\Export-VcMeta.ps1 -Server <來源vC> -User administrator@vsphere.local -Password '<pw>' -OutDir .\export-linux -Folder 'Linux'"
```

```bash
pwsh -Command "& .\Import-VcMeta.ps1 -Server <目標vC> -User administrator@vsphere.local -Password '<pw>' -InDir .\export-linux -DatacenterMap '<來源DC>=<目標DC>' -Include Folders,VMPlacement,Tags,CustomAttributes,Notes -MoveVMs -DryRun"
```

範圍規則：

| 項目 | `-Folder` 下的行為 |
|---|---|
| 資料夾 | 指定路徑本身 + 所有子資料夾（VM 類型） |
| VM / 範本 | 只有這棵子樹裡的 |
| Notes | 只有範圍內 VM 的 |
| Tag 指派 / 屬性值 | 只有範圍內的**資料夾**與 **VM**；host / datastore / cluster 等一律不帶 |
| Tag 分類、標籤、屬性定義 | 只帶「範圍內物件真的用到」的那些（`-AllDefinitions` 可全帶） |

`-Folder` 兩邊都能用：匯出時縮範圍，或者已經有整台 vC 的匯出檔、匯入時才縮範圍（分批一個資料夾一個資料夾搬時很好用）。

## 8. 從 datastore 把 VM 註冊回 inventory `Register-VmxFromDatastore.ps1`

datastore 搬到新 vCenter 後檔案都在、inventory 是空的——掃 `.vmx` / `.vmtx` 把它們註冊回來。

```bash
pwsh -Command "& .\Register-VmxFromDatastore.ps1 -Server <目標vC> -User administrator@vsphere.local -Password '<pw>' -Datastore ds01 -Cluster cl01 -DryRun"
```

```bash
pwsh -Command "& .\Register-VmxFromDatastore.ps1 -Server <目標vC> -User administrator@vsphere.local -Password '<pw>' -Datastore ds01 -Cluster cl01 -PlacementCsv .\export-A\vm-placement.csv -CreateFolders"
```

參數：
- `-Datastore`：可多個；`-Cluster` / `-VMHost` 決定註冊到哪（叢集內有掛該 datastore 的主機輪流用）
- `-PlacementCsv`：拿 `Export-VcMeta` 的 `vm-placement.csv`，註冊完直接依 VMName 放進原本的資料夾（`-CreateFolders` 不存在就建）
- `-Folder`：沒 placement 對照時的預設資料夾（預設 DC 根）
- `-Include` / `-Exclude`：名稱 wildcard（預設排除 `vCLS*`）
- `-NoTemplates`：略過 `.vmtx`；`-NameFromFile`：用檔名當 VM 名（預設抓 vmx 內的 `displayName`）
- `-DryRun`：只掃、不註冊

規則：已註冊的 vmx（比對 `[datastore] 路徑`）一律跳過，重跑安全；`.vmtx` 註冊成範本；**只註冊不開機**（開機時的 moved/copied 詢問要自己回）。

自我測試 `Test-RegisterVmx.ps1`：建丟棄式 VM + 範本 → unregister → 註冊回來 → 驗證（路徑一致、範本仍是範本、placement 資料夾、沒開機、重跑冪等）→ 清除。實測 **13/13 ALL PASS**（vCenter 9.1.1，VMFS）。

### 整套「datastore 搬家」流程

```
vC A   Export-VcMeta.ps1                          → export-A\*.csv
       把 datastore 從 vC A 卸載、掛到 vC B（儲存端操作）
vC B   Register-VmxFromDatastore.ps1 -PlacementCsv export-A\vm-placement.csv   → VM 回 inventory + 進原資料夾
vC B   Import-VcMeta.ps1 -InDir export-A          → tag / 自訂屬性 / Notes 補上（VM 已在，全部對得到）
```
