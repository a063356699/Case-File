param(
  [string]$JsonPath = "C:\Users\User\Downloads\物件總表_2026-08-01_0943_土地欄位核對修正.json",
  [string]$MasterPath = "C:\Users\User\Downloads\物件總表.xlsm",
  [string]$SummaryPath = "C:\Users\User\Downloads\統整表.xlsx",
  [string]$WorkbookRoot = "C:\Users\User\Downloads\售屋資料表+不動產說明書\售屋資料表+不動產說明書",
  [string]$OutputCsv = "C:\Users\User\Downloads\物件編輯內容_三來源差異確認清單.csv"
)
$ErrorActionPreference = 'Stop'

function N([string]$s) { if(!$s){return ''}; return (($s-replace '[\s　\-—_｜|\[\]【】()（）,，、.。之號市縣區]','')-replace '臺','台').ToLower() }
function V([object]$v) { if($null-eq$v){return ''}; return ([string]$v).Trim() }
function Canon([object]$v) {
  $s=V $v
  if($s-match '^[-－—–_／/、.．無未填沒有]*$'){return ''}
  return (($s-replace '[\s　]','')-replace '臺','台'-replace '％','%'-replace '\*',''-replace '(公?米|坪)$','').ToLower()
}
function RowMap($sheet,$row){$h=@{};$cols=$sheet.UsedRange.Columns.Count;for($c=1;$c-le$cols;$c++){$k=V $sheet.Cells.Item(1,$c).Text;if($k){$h[$k]=V $sheet.Cells.Item($row,$c).Text}};return $h}
function GetLike($map,[string[]]$names){foreach($n in $names){foreach($k in $map.Keys){if($k -eq $n -or $k -like "*$n*"){return V $map[$k]}}};return ''}
function SplitPair([string]$s,[string]$first,[string]$second){$a='';$b='';if($s-match "$first(?:寬|度)?\s*([^\r\n]+)"){$a=V $Matches[1]};if($s-match "$second(?:寬|度)?\s*([^\r\n]+)"){$b=V $Matches[1]};return @($a,$b)}

$doc=Get-Content -Raw -Encoding UTF8 -LiteralPath $JsonPath|ConvertFrom-Json
$active=@($doc.records|Where-Object{-not$_.archived-and$_.status-eq'委託中'})
$master=@{};$summary=@{};$excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false
try{
  $w=$excel.Workbooks.Open($MasterPath,0,$true);$s=$w.Worksheets.Item('物件總表');for($r=2;$r-le$s.UsedRange.Rows.Count;$r++){$no=V $s.Cells.Item($r,1).Text;if($no-match '^[A-Z]{2}\d+'){$master[$no]=RowMap $s $r}};$w.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($s);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
  $w=$excel.Workbooks.Open($SummaryPath,0,$true);$s=$w.Worksheets.Item('統整表');for($r=2;$r-le$s.UsedRange.Rows.Count;$r++){$no=(V $s.Cells.Item($r,2).Text).Trim();if($no-match '^[A-Z]{2}\d+'){$summary[$no]=RowMap $s $r}};$w.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($s);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)

  $files=@(Get-ChildItem -LiteralPath $WorkbookRoot -Recurse -File|Where-Object{$_.Extension-match'^\.xls[mx]?$'-and$_.Name-notlike'~$*'}|ForEach-Object{[pscustomobject]@{File=$_;Name=N $_.BaseName}})
  function FindBook($rec){$a=N $rec.address;$c=N $rec.caseName;$x=@($files|Where-Object{($a.Length -ge 5 -and $_.Name.Contains($a)) -or ($c.Length -ge 6 -and $_.Name.Contains($c))});if(!$x.Count){$token=switch($rec.propertyNo){'EG0495413'{'安平路3047樓12'}'EG0507776'{'新義南路20'}'EG0522916'{'中華路一段323巷6'}default{''}};if($token){$x=@($files|Where-Object{$_.Name.Contains($token)})}};if(!$x.Count){return $null};return $x|Sort-Object @{Expression={$p=$_.File.FullName;$n=0;if($p -match '複製|檔案與文件|南北-其他區'){$n+=10000};if($_.File.Extension -ne '.xlsm'){$n+=1000};$n+$p.Length}}|Select-Object -First 1}

  $fieldDefs=@(
    @{f='案名';j='caseName';m=@('案名');u=@('案名');x=@('案名')},@{f='地址';j='address';m=@('地址');u=@('地址');x=@('物件(完整)地址')},
    @{f='契種';j='contractType';u=@('契種')},@{f='委託開始';j='entrustStart';u=@('委託開始');x=@('委託開始 日期')},@{f='委託結束';j='entrustEnd';u=@('委託結束')},
    @{f='物件狀態';j='status';u=@('物件狀態')},@{f='開價';j='price';m=@('總價');u=@('開價');x=@('契約開價')},@{f='種類';j='type';x=@('物件型態')},
    @{f='朝向';j='direction';m=@('朝向');x=@('朝向 [房屋朝]','朝向 [大門朝]','朝向 [土地朝]')},@{f='建築完成日期';j='completionDate';x=@('建築完成日期')},
    @{f='權狀樓層';j='titleFloor';x=@('權狀層數')},@{f='現況樓層';j='currentFloor';x=@('透天請寫')},@{f='樓層顯示';j='floor';m=@('樓層')},@{f='格局';j='layout';m=@('格局')},
    @{f='室內坪';j='indoorPing';m=@('室內坪');x=@('室內坪')},@{f='建坪';j='buildingPing';m=@('建坪');x=@('總建坪')},@{f='地坪';j='landPing';m=@('地坪');x=@('地坪')},
    @{f='車位';j='parking';m=@('車位');x=@('車位')},@{f='車位編號';j='parkingNo';x=@('車位編號')},@{f='管理方式';j='managementMethod';x=@('警衛管理')},@{f='管理費';j='managementFee';m=@('管理費');x=@('管理費')},
    @{f='鑰匙';j='key';m=@('鑰匙');x=@('鑰匙位置')},@{f='現況';j='currentState';m=@('現況');x=@('(物件)現況')},@{f='臨路';j='road';m=@('臨路');x=@('臨路')},
    @{f='使用分區';j='zoning';m=@('使用分區');x=@('使用分區')},@{f='開發';j='developer';m=@('開發');u=@('開發');x=@('開發１/開發２')},
    @{f='備註欄';j='notes';m=@('備註欄');u=@('註記')},@{f='進案報件日期';j='reportDate';m=@('進案報件日期')},@{f='更新日期';j='updateDate';m=@('更新日期')},@{f='團看日期';j='groupViewDate';m=@('團看日期')},
    @{f='物件本';j='bookLocationDate';u=@('物件本')},@{f='銷售本';j='salesBook';m=@('銷售本')},@{f='照片';j='photoInfo';u=@('物件照片')},
    @{f='591';j='platform591';u=@('591')},@{f='YES319';j='yes319';u=@('YES319')},@{f='HOUSEINFOR';j='houseinfor';u=@('HOUSEINFOR')},@{f='櫥窗廣告';j='windowAd';u=@('櫥窗廣告')},@{f='LED';j='led';u=@('LED')},@{f='我家網';j='homeWeb';u=@('我家網')},@{f='5168';j='price5168';u=@('5168')},@{f='黃金曝光';j='goldExposure';u=@('黃金曝光')}
  )
  $out=New-Object System.Collections.Generic.List[object]
  foreach($rec in $active){$no=V $rec.propertyNo;$mm=$master[$no];$uu=$summary[$no];$xx=@{};$cand=FindBook $rec;if($cand){$wb=$null;$sh=$null;try{$wb=$excel.Workbooks.Open($cand.File.FullName,0,$true);try{$sh=$wb.Worksheets.Item('回應')}catch{$sh=$null};if($sh){$xx=RowMap $sh 2}}catch{}finally{if($sh){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($sh)};if($wb){$wb.Close($false);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)}}}
    foreach($d in $fieldDefs){$local=V $rec.($d.j);$mv=if($d.m){GetLike $mm $d.m}else{''};$uv=if($d.u){GetLike $uu $d.u}else{''};$xv=if($d.x){GetLike $xx $d.x}else{''};$sources=@($mv,$uv,$xv)|Where-Object{Canon $_};$distinct=@($sources|ForEach-Object{Canon $_}|Select-Object -Unique);$localDiff=$sources.Count -and ($sources|Where-Object{(Canon $_)-ne(Canon $local)}).Count;$sourceConflict=$distinct.Count -gt 1;if($localDiff -or $sourceConflict){$out.Add([pscustomobject]@{物件編號=$no;案名=$rec.caseName;欄位=$d.f;單機內容=$local;物件總表=$mv;統整表=$uv;物件Excel=$xv;判斷=if($sourceConflict){'來源互相衝突，請確認'}else{'單機與來源不同，可恢復'}})}}
    if($mm){$pair=SplitPair (GetLike $mm @('面寬','深度')) '面' '深';$cov=(GetLike $mm @('建蔽','容積'))-split'[/／]';foreach($z in @(@('面寬','frontage',$pair[0]),@('深度','depth',$pair[1]),@('建蔽率','coverage',$(if($cov.Count){$cov[0]}else{''})),@('容積率','far',$(if($cov.Count -gt 1){$cov[1]}else{''})))){if((Canon $z[2])-ne(Canon $rec.($z[1]))){$out.Add([pscustomobject]@{物件編號=$no;案名=$rec.caseName;欄位=$z[0];單機內容=$rec.($z[1]);物件總表=$z[2];統整表='';物件Excel='';判斷='單機與來源不同，可恢復'})}}}
  }
  $out|Sort-Object 物件編號,欄位|Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $OutputCsv
  [pscustomobject]@{委託中=$active.Count;差異列=$out.Count;來源衝突=@($out|Where-Object{$_.判斷-like'來源*'}).Count;可直接恢復=@($out|Where-Object{$_.判斷-like'單機*'}).Count;清單=$OutputCsv}|Format-List
}finally{$excel.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
