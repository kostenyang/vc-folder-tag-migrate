# vCenter 中繼資料搬家（Folder / Tag / Custom Attribute / Notes）

把 **vCenter A** 的 Folder 結構、Tag、Custom Attribute、VM Notes 搬到 **vCenter B**，
以及 datastore 搬過去之後把 VM 註冊回 inventory 並照舊 vC 補齊。PowerCLI 13.5，`pwsh` 執行。

| 腳本 | 用途 |
|---|---|
| `Copy-VcCustomAttributes.ps1` | **一支搞定**：自訂屬性（定義 + 值）直接 A → B |
| `Copy-VcFolders.ps1` | **一支搞定**：資料夾樹直接建到 B（可連 VM 一起放進去） |
| `Copy-VcMeta.ps1` | 通用版直接 A → B，`-Include` 選要搬哪幾類 |
| `Unregister-VmFromOldVc.ps1` | 舊 vC 端：依匯出清單把 VM 移出 inventory（檔案留著），寫 `unregistered.csv` 給新 vC 對帳 |
| `Register-VmxFromDatastore.ps1` | 新 vC 端：掃 datastore 把 vmx/vmtx 註冊回來；`-MetaDir` 就**照舊 vC 把資料夾 / 屬性 / Notes / tag 一起做好**，結尾**對帳**列出還沒過來的 |
| `Move-VmToNewVc.ps1` | 兩台 vC **同時連得到**時用：每台「新 vC 先註冊 → 成功才從舊 vC 拿掉」，沒有看不到的空窗；可選自動關機 / 開機回答 moved |
| `Export-VcMeta.ps1` / `Import-VcMeta.ps1` | 底層兩步式（先落 CSV、可人工編修、再匯入）；上面幾支都是包這兩支 |
| `Test-VcMeta.ps1` / `Test-RegisterVmx.ps1` | 端到端自我測試（建測試物件 → 跑 → 獨立驗證 → 清除） |

## 0. 主線：datastore 搬家（兩台 vC 不必同時在線，中間靠落地檔）

```
舊 vC  ① Export-VcMeta.ps1            → export-A\      落地：資料夾 / VM 位置 / tag / 屬性 / Notes（這份 = 清單 = 真相源）
舊 vC  ② Unregister-VmFromOldVc.ps1    → 依清單 unregister（開著的不動），寫 export-A\unregistered.csv
       …datastore 卸載、搬到新 vC…（隔幾小時、幾天都可以）
新 vC  ③ Register-VmxFromDatastore.ps1 -MetaDir export-A
          → 註冊 + 放回原資料夾 + 屬性 + Notes + tag，結尾對帳：unregistered.csv 裡誰還沒過來
```

```bash
pwsh -Command "& .\Export-VcMeta.ps1 -Server <舊vC> -User administrator@vsphere.local -Password '<pw>' -OutDir .\export-A -Folder 'Linux'"
```
```bash
pwsh -Command "& .\Unregister-VmFromOldVc.ps1 -Server <舊vC> -Password '<pw>' -MetaDir .\export-A -Datastore ds01 -DryRun"
```
```bash
pwsh -Command "& .\Register-VmxFromDatastore.ps1 -Server <新vC> -Password '<pw>' -Datastore ds01 -Cluster cl01 -MetaDir .\export-A -DatacenterMap '<舊DC>=<新DC>' -DryRun"
```

- ② 的安全機制：沒有 ① 的匯出檔不給做；清單的 vmx 路徑要跟現在一致（防舊匯出檔）；開著的 VM 跳過（或 `-ShutdownFirst`）；不接受「全部」，要給 `-Datastore` / `-Folder` / `-VM`。
- ③ 可以分批跑很多次，每次都會對帳；VM 看不到的那段時間，`unregistered.csv` 就是你的清單。
- 兩台 vC 同時連得到、想完全沒有空窗 → 用 `Move-VmToNewVc.ps1`（見 §10）。

## 0b. 只搬中繼資料（VM 已經在新 vC）

**只搬自訂屬性**
```bash
pwsh -Command "& .\Copy-VcCustomAttributes.ps1 -SourceServer <舊vC> -SourcePassword '<pw>' -TargetServer <新vC> -TargetPassword '<pw>' -DatacenterMap '<舊DC>=<新DC>' -DryRun"
```

**只搬資料夾樹（新 vC 已有同名/同 UUID 的 VM 就順便搬進去）**
```bash
pwsh -Command "& .\Copy-VcFolders.ps1 -SourceServer <舊vC> -SourcePassword '<pw>' -TargetServer <新vC> -TargetPassword '<pw>' -DatacenterMap '<舊DC>=<新DC>' -MoveVMs -DryRun"
```


都先 `-DryRun` 看報告，確認後拿掉再跑。`-Folder 'Linux'` 可縮到單一資料夾子樹。

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

自我測試 `Test-RegisterVmx.ps1`：建丟棄式 VM + 範本 → unregister → 註冊回來 → 驗證（路徑一致、範本仍是範本、placement 資料夾、沒開機、重跑冪等）→ 清除。舊版 13 項已併入下方「照舊 vC」的 19 項測試（vCenter 9.1.1，VMFS）。

### 照舊 vC 一次做好：`-SourceServer` / `-MetaDir`

加上其中一個，註冊完會自動：**放回原本的資料夾**（缺的資料夾自動建）→ **補自訂屬性定義與值** → **補 Notes**（其實 vmx 自帶，會回報 `AlreadySet`）→ **重建 tag 分類/標籤並指派**。只補剛註冊的那幾台，不會動到新 vC 上其他東西。

```bash
# 直接連舊 vC 抓（內部先跑 Export-VcMeta 到 -MetaDir，預設 .\meta-<舊vC>-<時間>）
pwsh -Command "& .\Register-VmxFromDatastore.ps1 -Server <新vC> -Password '<pw>' -Datastore ds01 -Cluster cl01 -SourceServer <舊vC> -SourcePassword '<pw>' -DatacenterMap '<舊DC>=<新DC>'"
```

```bash
# 舊 vC 已連不到：用先前 Export-VcMeta 匯出的目錄
pwsh -Command "& .\Register-VmxFromDatastore.ps1 -Server <新vC> -Password '<pw>' -Datastore ds01 -Cluster cl01 -MetaDir .\export-A -DatacenterMap '<舊DC>=<新DC>'"
```

報告有兩份：`register-report-*.csv`（註冊/放資料夾）與 `*-meta.csv`（屬性/Notes/tag）。

自我測試 `Test-RegisterVmx.ps1` 就是照這條路測的：建 VM（放兩層資料夾 + 中文屬性值 + Notes + tag）→ 匯出當舊 vC 快照 → unregister 並把資料夾/屬性定義/tag 分類全刪 → `-MetaDir` 註冊 → 獨立驗證每一項照舊 → 冪等 → 清除。實測 **19/19 ALL PASS**（vCenter 9.1.1）。

### 整套「datastore 搬家」流程（順序很重要）

```
vC A   ① Export-VcMeta.ps1 → export-A\        ← 🔴 一定先做：VM 一 unregister，A 端的 tag 指派 / 屬性值就沒了
       ② unregister VM（或直接把 datastore 卸載）
       ③ datastore 掛到 vC B（儲存端操作）
vC B   ④ Register-VmxFromDatastore.ps1 -MetaDir export-A -DatacenterMap 'A=B'
          → VM 回 inventory + 原資料夾 + 自訂屬性 + Notes + tag，一次做完
```

`-SourceServer`（現場連舊 vC 抓）只有在舊 vC 的 VM **還在 inventory** 時才抓得到（例如 datastore 直接拔掉、VM 在 A 變 orphaned 那種情況）。已經 unregister 就只能靠 ①的匯出檔——所以預設就用 `-MetaDir`。

### 實機跨 vC 驗證（2026-09-14）

不是模擬，是真的兩台 vCenter + 一顆共用 NFS：

| 步驟 | 內容 | 結果 |
|---|---|---|
| 準備 | vcd-nfs01 加 export → 掛成 NFS datastore 到 vC A 的 host 與 vC B 的 host | 兩邊都看到同一顆 191GB |
| vC A | 建 `zz-real-vm01`（1GB vmdk）放 `ZZ-Real/Sub`，屬性 `zz-real-owner=基礎架構組-Kosten`，Notes，tag `zz-real-env/prod` | |
| vC A | `Export-VcMeta -Folder ZZ-Real` → 8 個 CSV | 資料夾 2 / VM 1 / 屬性 1 / Notes 1 / tag 1 |
| vC A | unregister | inventory 不見 |
| vC B | `Register-VmxFromDatastore -MetaDir ... -DryRun` | 預告 1 台 → `ZZ-Real/Sub` |
| vC B | 正式跑 | Registered 1 / CreatedFolder 2 / Placed 1 / CA Created+Set / TagCategory+Tag Created / Assigned |
| vC B | 獨立查詢核對 | 名稱、vmx、資料夾、屬性值、Notes、tag **全部與 A 端一致**；PoweredOff；**InstanceUuid 跨 vC 保留** |
| 拆除 | VM(含檔)、資料夾、屬性、tag、兩邊 datastore 卸載、NFS export 移除 | 兩邊複驗 clean |

順帶驗到一件事：`InstanceUuid` 註冊到另一台 vCenter 後不會變，所以 Import 的 UUID 比對在真實搬移也有效，不必只靠名稱。
## 9. 直接 A → B 的單一腳本

不想分兩步、也不需要改 CSV 時用這些。內部就是 `Export-VcMeta → Import-VcMeta`，CSV 與報告留在 `-WorkDir`（預設 `.\copy-<來源>-<時間>`）。

| 腳本 | 等同 |
|---|---|
| `Copy-VcCustomAttributes.ps1` | `Copy-VcMeta.ps1 -Include CustomAttributes` |
| `Copy-VcFolders.ps1` | `Copy-VcMeta.ps1 -Include Folders`（加 `-MoveVMs` 則含 `VMPlacement`） |
| `Copy-VcMeta.ps1 -Include Tags,Notes` | 其他組合自己選 |

共同參數：`-SourceServer/-SourceUser/-SourcePassword`、`-TargetServer/-TargetUser/-TargetPassword`、`-DatacenterMap`、`-Folder`、`-Datacenter`、`-DryRun`、`-WorkDir`。

## 10. 兩台 vC 同時連得到：`Move-VmToNewVc.ps1`（零空窗交接）

前提：VM 所在 datastore 兩邊都掛得到。每台 VM：確認關機（或 `-ShutdownFirst`）→ **新 vC 先註冊** → 成功才從舊 vC unregister（失敗就不動舊 vC）→ 放回原資料夾 + 屬性 + Notes + tag → `-PowerOnAfter` 在新 vC 開機並自動回答「I moved it」。匯出在一開始自動做（VM 還在舊 vC）。

```bash
pwsh -Command "& .\Move-VmToNewVc.ps1 -SourceServer <舊vC> -SourcePassword '<pw>' -TargetServer <新vC> -TargetPassword '<pw>' -Folder 'Linux' -TargetCluster cl01 -DatacenterMap '<舊DC>=<新DC>' -DryRun"
```

實機驗證（2026-09-14）：`zz-real-vm03` 從 vC A `ZZ-Real/Sub` 交接到 vC B，Moved → 開機 → 屬性 / tag / 資料夾一致，舊 vC 上已不存在。

## 11. 落地三段式實機驗證（2026-09-14）

| 步驟 | 動作 | 結果 |
|---|---|---|
| ① 舊 vC | 兩台 VM（vm02 故意開著）`Export -Folder ZZ-Real` | 8 個 CSV，`vm-placement.csv` 帶 vmx 路徑與電源狀態 |
| ② 舊 vC | `Unregister -Datastore vc-migtest -DryRun` | vm01 WouldUnregister、**vm02 SkippedPoweredOn** |
| ② 舊 vC | 關掉 vm02 後正式跑 | 2 台 Unregistered，`unregistered.csv` 2 筆 |
| ③ 新 vC | `Register -MetaDir -Include zz-real-vm01`（故意只做一台） | vm01 註冊 + 資料夾 + 屬性 + tag；**對帳：還沒過來 1 台 = vm02** |
| ③ 新 vC | 再跑一次不篩 | vm01 AlreadyRegistered、vm02 註冊補齊；對帳 0 台 |
| 核對 | 兩台在 `ZZ-Real/Sub`，屬性 / Notes / tag 與舊 vC 一致，PoweredOff | ✅ |

實跑抓到並修掉的 bug：PowerShell `-like "[ds] *"` 會把 `[ds]` 當字元集合 wildcard，datastore 篩選全部落空 → 改 `StartsWith`。
