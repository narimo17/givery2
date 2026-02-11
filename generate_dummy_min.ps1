param([int]$Count=3000,[string]$OutDir=".\data",[int]$Seed=42)
$ErrorActionPreference="Stop"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$rng=[Random]::new($Seed)

$Operators=@("岩里","篠原","不明","松本","Nishinetei Yuiki Hara","不明","松本","松本","小室","角屋","屋下","いわさ","不明","武田","川畑","中村","松本","山下","岩田","岩沙")
$Purposes=@("障害/不具合（TV/NET/電話）","機器/設定サポート（Wi‑Fi/メール/リモコン/STB等）","料金/請求","契約変更","解約","引越し","工事/訪問手配","営業案内","サービス案内","その他")

$Header=@("通話日時","オペレーター名","顧客属性(推定)","契約状況(推測)","入電目的(大分類)","具体的課題(Trigger)","解決プロセス(Action)","ライフイベント・予兆","解約リスク判定","解約リスク判定_詳細","評価_Soft_点数","評価_Soft_内容","評価_Soft_根拠","評価_Hard基本_点数","評価_Hard基本_内容","評価_Hard基本_根拠","評価_Hard専門_点数","評価_Hard専門_内容","評価_Hard専門_根拠","評価_判断提案_点数","評価_判断提案_内容","評価_判断提案_根拠","会話要約","会話全文(整形済)")

$start=Get-Date "2023-01-01"; $end=Get-Date "2023-12-31"
function Pick($a){ $a[$rng.Next(0,$a.Count)] }
function D([datetime]$d){ $d.ToString("yyyy年M月d日") }

$rows = New-Object System.Collections.Generic.List[object]
for($i=0;$i -lt $Count;$i++){
  $dt=$start.AddDays($rng.Next(0, ($end-$start).Days+1))
  $op=Pick $Operators
  $purpose=Pick $Purposes
  $contract = if($purpose -in @("営業案内","サービス案内")){"既加入者"}else{ Pick @("既加入者","既加入者","不明") }
  $attr = Pick @("不明","不明","既加入者")
  $risk = if($purpose -eq "解約"){"高"}elseif($purpose -eq "料金/請求" -and $rng.Next(0,100) -lt 25){"中"}else{"低"}
  $riskD = if($risk -eq "高"){"顧客から解約意向の発言あり"}elseif($risk -eq "中"){"料金見直し・他社比較の示唆"}else{"顧客からの解約に関する発言なし"}

  $soft=2+$rng.Next(0,4); $hb=2+$rng.Next(0,4); $hp=1+$rng.Next(0,5); $prop=2+$rng.Next(0,4)
  $trigger = switch($purpose){
    "障害/不具合（TV/NET/電話）" {"ネット不通/TV視聴不良の相談"}
    "機器/設定サポート（Wi‑Fi/メール/リモコン/STB等）" {"Wi‑Fi/機器設定の相談"}
    "料金/請求" {"請求内訳の確認"}
    "契約変更" {"プラン/オプション変更相談"}
    "解約" {"解約手続き相談"}
    "引越し" {"移転手続き相談"}
    "工事/訪問手配" {"訪問サポート依頼"}
    "営業案内" {"セット割/モバイル案内"}
    "サービス案内" {"サービスのご案内のための連絡"}
    default {"その他問い合わせ"}
  }
  $action = if($purpose -in @("営業案内","サービス案内")){"電話完結"}else{"要件確認→案内→終話"}
  $life = if($purpose -in @("解約","引越し")){"引越し/費用見直し"}elseif($purpose -eq "料金/請求"){"家計見直し"}else{"特になし"}

  $summary="オペレーターの$opが$contractの顧客に対して$trigger。対応は$action。解約リスクは$risk。"
  $conv="OP: お世話になります。愛媛ケーブルテレビの$opです。 CU: はい。 OP: $trigger の件でご案内します。 CU: はい。 OP: $action。 CU: わかりました。"

  $o=[ordered]@{}
  $o["通話日時"]=D $dt; $o["オペレーター名"]=$op; $o["顧客属性(推定)"]=$attr; $o["契約状況(推測)"]=$contract
  $o["入電目的(大分類)"]=$purpose; $o["具体的課題(Trigger)"]=$trigger; $o["解決プロセス(Action)"]=$action
  $o["ライフイベント・予兆"]=$life; $o["解約リスク判定"]=$risk; $o["解約リスク判定_詳細"]=$riskD
  $o["評価_Soft_点数"]=$soft; $o["評価_Soft_内容"]="丁寧な対応（ダミー）"; $o["評価_Soft_根拠"]="お世話になります／ありがとうございます（ダミー）"
  $o["評価_Hard基本_点数"]=$hb; $o["評価_Hard基本_内容"]="基本案内（ダミー）"; $o["評価_Hard基本_根拠"]="要件確認の発話（ダミー）"
  $o["評価_Hard専門_点数"]=$hp; $o["評価_Hard専門_内容"]="専門案内（ダミー）"; $o["評価_Hard専門_根拠"]="切り分け/手続き（ダミー）"
  $o["評価_判断提案_点数"]=$prop; $o["評価_判断提案_内容"]="次アクション提示（ダミー）"; $o["評価_判断提案_根拠"]="希望日確認/折返し（ダミー）"
  $o["会話要約"]=$summary; $o["会話全文(整形済)"]=$conv
  $rows.Add([pscustomobject]$o) | Out-Null
}

# TSV
$tsv=Join-Path $OutDir "voc.tsv"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$sb=New-Object System.Text.StringBuilder
[void]$sb.AppendLine(($Header -join "`t"))
foreach($r in $rows){
  $vals = foreach($h in $Header){ ([string]$r.$h).Replace("`t"," ").Replace("`r"," ").Replace("`n"," ") }
  [void]$sb.AppendLine(($vals -join "`t"))
}
[IO.File]::WriteAllText($tsv, $sb.ToString(), $utf8NoBom)

# JSON
$rows | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $OutDir "voc.json") -Encoding UTF8

Write-Host "DONE -> $OutDir\voc.tsv / voc.json"
